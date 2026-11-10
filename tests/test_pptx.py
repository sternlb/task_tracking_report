"""Tests for the OPC layer, the package lint and deck assembly (M1)."""

from __future__ import annotations

import io
import os
import sys
import unittest
import zipfile

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, ROOT)
sys.path.insert(0, os.path.join(ROOT, "tools"))

from statusreport.pptx import opc  # noqa: E402
from statusreport.pptx.deck import Deck  # noqa: E402
from statusreport.pptx.lint import lint_package  # noqa: E402
from statusreport.pptx.shapes import Para, Run, color_xml  # noqa: E402
from statusreport.xlsx_write import col_letter  # noqa: E402

import spike  # noqa: E402


def spike_bytes() -> bytes:
    return spike.build(None).to_bytes()


class OpcTest(unittest.TestCase):
    def test_rels_name_roundtrip(self):
        for part in ("ppt/presentation.xml", "ppt/slides/slide1.xml", ""):
            self.assertEqual(opc.source_of_rels(opc.rels_name(part)), part)
        self.assertEqual(opc.rels_name("ppt/slides/slide1.xml"), "ppt/slides/_rels/slide1.xml.rels")

    def test_resolve_and_relative(self):
        src = "ppt/slides/slide1.xml"
        self.assertEqual(opc.resolve(src, "../slideLayouts/slideLayout2.xml"), "ppt/slideLayouts/slideLayout2.xml")
        self.assertEqual(opc.resolve(src, "/ppt/media/image%201.png"), "ppt/media/image 1.png")
        self.assertEqual(opc.relative(src, "ppt/charts/chart1.xml"), "../charts/chart1.xml")
        self.assertEqual(opc.relative("", "ppt/presentation.xml"), "ppt/presentation.xml")

    def test_drop_unreachable(self):
        pkg = opc.Package()
        pkg.add_part("a.xml", b"<a/>")
        pkg.add_part("b.xml", b"<b/>")
        pkg.add_part("orphan.xml", b"<o/>")
        pkg.add_rel("", opc.RT_OFFICE_DOCUMENT, "a.xml")
        pkg.add_rel("a.xml", opc.RT_THEME, "b.xml")
        pkg.add_rel("orphan.xml", opc.RT_THEME, "b.xml")
        self.assertEqual(pkg.drop_unreachable(), ["orphan.xml"])
        self.assertNotIn("_rels/orphan.xml.rels", pkg.parts)
        self.assertIn("b.xml", pkg.parts)

    def test_output_is_deterministic(self):
        self.assertEqual(spike_bytes(), spike_bytes())


class LintTest(unittest.TestCase):
    def test_base_is_clean(self):
        self.assertEqual(lint_package(Deck.default().pkg), [])

    def test_spike_is_clean(self):
        self.assertEqual(lint_package(opc.Package.from_bytes(spike_bytes())), [])

    def test_detects_dangling_relationship(self):
        pkg = opc.Package.from_bytes(spike_bytes())
        pkg.parts.pop("ppt/charts/chart1.xml")
        problems = lint_package(pkg)
        self.assertTrue(any("missing part ppt/charts/chart1.xml" in p for p in problems), problems)

    def test_detects_missing_content_type(self):
        pkg = opc.Package.from_bytes(spike_bytes())
        pkg.parts["ppt/media/strange.bin"] = b"x"
        self.assertTrue(any("strange.bin: no content type" in p for p in lint_package(pkg)))

    def test_detects_broken_xml_and_duplicate_ids(self):
        pkg = opc.Package.from_bytes(spike_bytes())
        slide = pkg.parts["ppt/slides/slide1.xml"]
        pkg.parts["ppt/slides/slide1.xml"] = slide.replace(b'id="3"', b'id="2"')
        pkg.parts["ppt/slides/slide2.xml"] = b"<p:sld"
        problems = lint_package(pkg)
        self.assertTrue(any("slide1.xml: duplicate shape ids" in p for p in problems), problems)
        self.assertTrue(any("slide2.xml: XML not well-formed" in p for p in problems), problems)


class DeckTest(unittest.TestCase):
    def test_spike_structure(self):
        pkg = opc.Package.from_bytes(spike_bytes())
        slides = pkg.related(pkg.main_part(), opc.RT_SLIDE)
        self.assertEqual(len(slides), 3)
        charts = pkg.related("ppt/slides/slide2.xml", opc.RT_CHART)
        self.assertEqual(charts, ["ppt/charts/chart1.xml"])
        book = pkg.related(charts[0], opc.RT_PACKAGE)[0]
        with zipfile.ZipFile(io.BytesIO(pkg.parts[book])) as zf:
            self.assertIn("xl/worksheets/sheet1.xml", zf.namelist())

    def test_template_slides_are_removed(self):
        # a finished deck used as a "customer template": its slides must vanish
        deck = Deck(spike_bytes())
        self.assertEqual(deck.slides, [])
        self.assertFalse([n for n in deck.pkg.parts if n.startswith(("ppt/slides/", "ppt/charts/", "ppt/embeddings/"))])
        self.assertEqual(set(deck.layouts), {"Title Slide", "Title Only", "Title and Content"})
        s = deck.add_slide("Title Only")
        s.set_title([Para([Run("again")])])
        self.assertEqual(lint_package(opc.Package.from_bytes(deck.to_bytes())), [])

    def test_unknown_layout_falls_back_with_warning(self):
        deck = Deck.default()
        layout = deck.layout("Does Not Exist")
        self.assertIsNotNone(layout.find("title", "ctrTitle"))
        self.assertEqual(len(deck.warnings), 1)

    def test_layout_lookup_ignores_case(self):
        self.assertEqual(Deck.default().layout("title only").name, "Title Only")

    def test_content_area_below_title(self):
        deck = Deck.default()
        area = deck.content_area(deck.layout("Title Only"))
        title = deck.layout("Title Only").find("title").box
        self.assertGreater(area.y, title.y + title.h)
        self.assertLessEqual(area.y + area.h, deck.height)


class HelpersTest(unittest.TestCase):
    def test_col_letter(self):
        self.assertEqual([col_letter(i) for i in (0, 25, 26, 51, 52, 701, 702)], ["A", "Z", "AA", "AZ", "BA", "ZZ", "AAA"])

    def test_color_xml(self):
        self.assertEqual(color_xml("#2e9e44"), '<a:srgbClr val="2E9E44"/>')
        self.assertEqual(color_xml("accent1"), '<a:schemeClr val="accent1"/>')
        self.assertIn('<a:lumMod val="20000"/><a:lumOff val="80000"/>', color_xml("accent1/20"))


if __name__ == "__main__":
    unittest.main()
