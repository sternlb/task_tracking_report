"""Tests for M6: foreign templates saved by real PowerPoint (R-OUT-8, T-6.1 … T-6.3).

The fixtures in tests/fixtures/templates come from tools/make_templates.ps1:
German layout names, a .potx, 4:3, renamed layouts, existing slides with
picture, notes, sections and a custom show.
"""

from __future__ import annotations

import os
import re
import sys
import unittest

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, ROOT)

from statusreport.calc import compute  # noqa: E402
from statusreport.model import Settings, Task  # noqa: E402
from statusreport.pptx import opc  # noqa: E402
from statusreport.pptx.deck import Deck  # noqa: E402
from statusreport.pptx.lint import lint_package  # noqa: E402
from statusreport.pptx.shapes import Box, Cell, Para, Row, Run, inch, table_frame  # noqa: E402
from statusreport.render import LEGEND_H, metrics, render  # noqa: E402
from statusreport.report import build_report  # noqa: E402

TPL = os.path.join(ROOT, "tests", "fixtures", "templates")
TEMPLATES = ["office_16x9.pptx", "office_16x9.potx", "facet_4x3.pptx", "renamed_cust.pptx"]
EXPECTED_LAYOUTS = {  # (title slide, content) chosen without a matching name
    "office_16x9.pptx": ("Titelfolie", "Nur Titel"),
    "office_16x9.potx": ("Titelfolie", "Nur Titel"),
    "facet_4x3.pptx": ("Titelfolie", "Nur Titel"),
    "renamed_cust.pptx": ("Cover", "Content"),
}


def deck(name: str) -> Deck:
    return Deck.from_file(os.path.join(TPL, name))


def render_demo(d: Deck, n_tasks: int = 20, footer: str = "", **kw) -> opc.Package:
    s = Settings(project="Falcon", customer="ACME", report_date=Settings().report_date)
    for k, v in kw.items():
        setattr(s, k, v)
    ts = [Task(id=f"T-{i:02d}", title=f"Task number {i} with a fairly long title for the column", version="1.0",
               progress_manual=float(i * 4), comment="a comment that is longer than the comment column") for i in range(n_tasks)]
    versions, _ = compute(ts, s)
    r = build_report(s, versions, None, "2026-W46", ["CW 45", "CW 46"], [("1.0", [10.0, 20.0])])
    render(r, d, s.title_layout, s.content_layout, s.rows_per_slide, footer)
    return opc.Package.from_bytes(d.to_bytes())


class StripTest(unittest.TestCase):
    def test_existing_slides_and_their_parts_are_removed(self):
        for name in TEMPLATES:
            d = deck(name)
            leftovers = [p for p in d.pkg.parts if p.startswith(("ppt/slides/", "ppt/notesSlides/", "ppt/media/", "ppt/charts/"))]
            self.assertEqual(leftovers, [], name)
            pres = d.pkg.parts[d.pres].decode("utf-8")
            for tag in ("sldIdLst", "custShowLst", "sectionLst"):
                self.assertNotIn(tag, pres, f"{name}: {tag}")
            self.assertEqual(lint_package(d.pkg), [], name)

    def test_potx_becomes_a_presentation(self):
        pkg = render_demo(deck("office_16x9.potx"))
        self.assertEqual(pkg.content_type(pkg.main_part()), opc.CT_PRESENTATION)

    def test_other_files_are_rejected(self):
        xlsx = os.path.join(ROOT, "tests", "fixtures", "excel_filled.xlsx")
        with self.assertRaisesRegex(ValueError, "not a PowerPoint"):
            Deck.from_file(xlsx)


class LayoutTest(unittest.TestCase):
    def test_fallback_by_layout_type(self):
        for name, (title, content) in EXPECTED_LAYOUTS.items():
            d = deck(name)
            self.assertEqual(d.layout("Title Slide", "title").name, title, name)
            self.assertEqual(d.layout("Title Only", "content").name, content, name)
            self.assertEqual(len(d.warnings), 2, name)
            d.layout("Title Only", "content")
            self.assertEqual(len(d.warnings), 2, "one warning per name and role")
            self.assertIn("Settings", d.warnings[0])

    def test_name_from_settings_wins_without_warning(self):
        d = deck("office_16x9.pptx")
        self.assertEqual(d.layout("titel und inhalt", "content").name, "Titel und Inhalt")
        self.assertEqual(d.warnings, [])

    def test_content_area_stays_inside_the_slide_and_above_the_footer(self):
        for name in TEMPLATES:
            d = deck(name)
            lay = d.layout("Title Only", "content")
            area = d.content_area(lay)
            footer_top = min(p.box.y for p in lay.placeholders if p.type in ("dt", "ftr", "sldNum") and p.box)
            self.assertGreaterEqual(area.x, 0, name)
            self.assertLessEqual(area.x + area.w, d.width, name)
            self.assertLessEqual(area.y + area.h, footer_top, name)
            title = lay.find("title")
            self.assertGreaterEqual(area.y, title.box.y + title.box.h, name)


class FitTest(unittest.TestCase):
    def test_table_fits_the_content_area(self):
        for name in TEMPLATES + [None]:
            d = deck(name) if name else Deck.default()
            area = d.content_area(d.layout("Title Only", "content"))
            m = metrics(area, 12)
            needed = m.header_h + (m.per_page + 1) * m.row_h + LEGEND_H
            self.assertLessEqual(needed, area.h, name)
            self.assertGreaterEqual(m.per_page, 3, name)

    def test_small_areas_get_fewer_rows_and_smaller_fonts(self):
        base = metrics(Deck.default().content_area(Deck.default().layout("Title Only")), 12)
        facet_deck = deck("facet_4x3.pptx")
        facet = metrics(facet_deck.content_area(facet_deck.layout("Title Only", "content")), 12)
        self.assertEqual((base.per_page, base.scale), (12, 1.0))
        self.assertLess(facet.per_page, base.per_page)
        self.assertLess(facet.scale, 1.0)
        self.assertLess(facet.title_chars, base.title_chars)

    def test_rows_per_slide_is_a_maximum(self):
        m = metrics(Deck.default().content_area(Deck.default().layout("Title Only")), 5)
        self.assertEqual(m.per_page, 5)

    def test_more_pages_on_small_templates(self):
        for name in TEMPLATES:
            pkg = render_demo(deck(name), n_tasks=20)
            self.assertEqual(lint_package(pkg), [], name)
            slides = pkg.related(pkg.main_part(), opc.RT_SLIDE)
            per_page = metrics(deck(name).content_area(deck(name).layout("Title Only", "content")), 12).per_page
            expected_pages = -(-20 // per_page)
            self.assertEqual(len(slides), 2 + expected_pages, name)


class FooterTest(unittest.TestCase):
    def test_slide_numbers_on_content_slides_only(self):
        pkg = render_demo(deck("office_16x9.pptx"), n_tasks=3)
        slides = pkg.related(pkg.main_part(), opc.RT_SLIDE)
        first = pkg.parts[slides[0]].decode()
        self.assertNotIn('type="sldNum"', first)
        for n, part in enumerate(slides[1:], start=2):
            xml = pkg.parts[part].decode()
            self.assertIn('type="sldNum"', xml)
            self.assertRegex(xml, rf'type="slidenum"><a:rPr lang="en-US"/><a:t>{n}</a:t>')
            self.assertNotIn('type="dt"', xml)

    def test_footer_text(self):
        pkg = render_demo(deck("facet_4x3.pptx"), n_tasks=3, footer="ACME confidential")
        xml = pkg.parts["ppt/slides/slide2.xml"].decode()
        self.assertIn('type="ftr"', xml)
        self.assertIn("ACME confidential", xml)
        plain = render_demo(deck("facet_4x3.pptx"), n_tasks=3)
        self.assertNotIn('type="ftr"', plain.parts["ppt/slides/slide2.xml"].decode())  # empty footer left out

    def test_built_in_base_has_no_footer_placeholders(self):
        pkg = render_demo(Deck.default(), n_tasks=3)
        self.assertNotIn("sldNum", pkg.parts["ppt/slides/slide2.xml"].decode())


class MergedCellTest(unittest.TestCase):
    def test_grid_span(self):
        cells = [Cell([Para([Run("Version total")])], span=2), Cell([Para([Run("x")])]), Cell([Para([Run("y")])])]
        xml = table_frame(2, "T", Box(0, 0, 400, 100), [100, 100, 100, 100], [Row(cells, inch(0.3))])
        self.assertIn('<a:tc gridSpan="2">', xml)
        self.assertEqual(xml.count('<a:tc hMerge="1">'), 1)
        self.assertEqual(len(re.findall(r"<a:tc[ >]", xml)), 4)  # one per grid column


if __name__ == "__main__":
    unittest.main()
