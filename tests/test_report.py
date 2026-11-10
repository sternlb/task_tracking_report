"""Tests for M5: report model, rendering, output naming and the full CLI run (R-OUT-*)."""

from __future__ import annotations

import contextlib
import datetime as dt
import io
import os
import re
import shutil
import sys
import tempfile
import unittest

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, ROOT)
sys.path.insert(0, os.path.join(ROOT, "tools"))

from statusreport import history, labels, output  # noqa: E402
from statusreport.__main__ import main  # noqa: E402
from statusreport.calc import compute  # noqa: E402
from statusreport.model import Settings, Task  # noqa: E402
from statusreport.pptx import opc  # noqa: E402
from statusreport.pptx.deck import Deck  # noqa: E402
from statusreport.pptx.lint import lint_package  # noqa: E402
from statusreport.render import render  # noqa: E402
from statusreport.report import build_report, paginate, prev_text, shorten  # noqa: E402

D = dt.date
ON = D(2026, 11, 10)


def settings(**kw) -> Settings:
    s = Settings(project="Falcon", customer="ACME", report_date=ON)
    for k, v in kw.items():
        setattr(s, k, v)
    return s


def tasks(n: int, version: str = "1.0") -> list:
    return [Task(id=f"T-{i:02d}", title=f"Task {i}", version=version, progress_manual=float(i)) for i in range(n)]


def make_prev(**task_progress) -> history.Snapshot:
    snap = history.Snapshot("2026-W45", D(2026, 11, 3))
    for tid, p in task_progress.items():
        snap.tasks[tid.replace("_", "-")] = history.TaskSnap("1.0", "", p, "orange")
    snap.versions["1.0"] = history.VersionSnap(5.0, None, "green")
    return snap


def slide_texts(pkg: opc.Package, part: str) -> str:
    return "".join(re.findall(r"<a:t>([^<]*)</a:t>", pkg.parts[part].decode("utf-8")))


class ReportModelTest(unittest.TestCase):
    def test_prev_text(self):
        snap = make_prev()
        self.assertEqual(prev_text(None, False, None), "(–)")  # first report ever
        self.assertEqual(prev_text(snap, False, None), "(new)")  # not in last week's report
        self.assertEqual(prev_text(snap, True, None), "(–)")  # was there, without a value
        self.assertEqual(prev_text(snap, True, 49.6), "(50 %)")

    def test_titles_and_subtitle(self):
        versions, _ = compute(tasks(2), settings())
        r = build_report(settings(), versions, None, "2026-W46", ["CW 46"], [("1.0", [1.0])])
        self.assertEqual(r.title, "Falcon")
        self.assertEqual(r.subtitle, "Status report · ACME · CW 46 / 2026 · 10 Nov 2026")
        self.assertEqual(r.trend_note, labels.TREND_FIRST_WEEK)  # only one week so far
        r2 = build_report(settings(project="", customer=""), versions, None, "2026-W46", ["CW 45", "CW 46"], [])
        self.assertEqual((r2.title, r2.subtitle), ("Software status", "Status report · CW 46 / 2026 · 10 Nov 2026"))
        self.assertEqual(r2.trend_note, "")

    def test_pages_split_and_repeat_the_total_row(self):
        versions, _ = compute(tasks(25), settings())
        (table,) = build_report(settings(), versions, None, "2026-W46", [], []).tables
        pages = paginate(table, 12)
        self.assertEqual([p.title for p in pages], ["Version 1.0 (1/3)", "Version 1.0 (2/3)", "Version 1.0 (3/3)"])
        self.assertEqual([len(p.rows) for p in pages], [12, 12, 1])
        self.assertTrue(all(p.total.is_total and p.total.id == "Version total" for p in pages))
        self.assertEqual([p.title for p in paginate(table, 30)], ["Version 1.0"])

    def test_previous_values_and_light_changes(self):
        ts = tasks(3)
        versions, _ = compute(ts, settings())
        prev = make_prev(T_00=0.0, T_01=None)
        (table,) = build_report(settings(), versions, prev, "2026-W46", [], []).tables
        (page,) = paginate(table, 12)
        rows = {r.id: r for r in page.rows}
        self.assertEqual(rows["T-00"].prev_text, "(0 %)")
        self.assertEqual(rows["T-01"].prev_text, "(–)")
        self.assertEqual(rows["T-02"].prev_text, "(new)")
        self.assertEqual(rows["T-00"].light, "grey")
        self.assertEqual(rows["T-00"].prev_light, "orange")  # changed -> shown
        self.assertIsNone(rows["T-02"].prev_light)  # new task: nothing to compare
        self.assertEqual(page.total.prev_text, "(5 %)")
        self.assertEqual((page.total.light, page.total.prev_light), ("grey", "green"))  # no plan now, green last week

    def test_shorten(self):
        self.assertEqual(shorten("short", 10), "short")
        self.assertEqual(shorten("a  b\nc", 10), "a b c")
        self.assertEqual(shorten("x" * 20, 10), "x" * 9 + "…")


class RenderTest(unittest.TestCase):
    def build(self, n=14, weeks=3) -> opc.Package:
        s = settings(rows_per_slide=12)
        versions, _ = compute(tasks(n) + tasks(2, "2.0"), s)
        labels_ = [f"CW {44 + i}" for i in range(weeks)]
        series = [(v.name, [1.0] * weeks) for v in versions]
        r = build_report(s, versions, make_prev(T_00=0.0), "2026-W46", labels_, series)
        deck = render(r, Deck.default(), s.title_layout, s.content_layout, s.rows_per_slide)
        return opc.Package.from_bytes(deck.to_bytes())

    def test_structure(self):
        pkg = self.build()
        self.assertEqual(lint_package(pkg), [])
        slides = pkg.related(pkg.main_part(), opc.RT_SLIDE)
        self.assertEqual(len(slides), 1 + 1 + 2 + 1)  # title, trend, 1.0 (2 pages), 2.0
        self.assertEqual(len(pkg.related(slides[1], opc.RT_CHART)), 1)
        texts = [slide_texts(pkg, s) for s in slides]
        self.assertIn("Falcon", texts[0])
        self.assertIn("CW 46 / 2026", texts[0])
        self.assertIn("Progress over time", texts[1])
        self.assertIn("Version 1.0 (1/2)", texts[2])
        self.assertIn("Progress (prev. week)", texts[2])
        self.assertIn("(0 %)", texts[2])
        self.assertIn("(new)", texts[2])
        self.assertNotIn("Owner", texts[2] + texts[3])  # R-OUT-3
        self.assertNotIn("Due", texts[2] + texts[3])
        self.assertIn("behind by more than 25 pp or overdue", texts[2])

    def test_table_size(self):
        pkg = self.build()
        slide = pkg.parts["ppt/slides/slide3.xml"].decode()
        self.assertEqual(slide.count("<a:tr "), 1 + 1 + 12)  # header, total, 12 tasks
        self.assertEqual(slide.count("<a:gridCol "), 5)

    def test_first_week_note(self):
        pkg = self.build(n=2, weeks=1)
        self.assertIn(labels.TREND_FIRST_WEEK, slide_texts(pkg, "ppt/slides/slide2.xml"))

    def test_deterministic(self):
        self.assertEqual(self.build().to_bytes(), self.build().to_bytes())


class OutputTest(unittest.TestCase):
    def test_default_name(self):
        self.assertEqual(output.default_name("Falcon", "2026-W46"), "Falcon Status 2026-W46.pptx")
        self.assertEqual(output.default_name('A/B: "C"?', "2026-W46"), "AB C Status 2026-W46.pptx")
        self.assertEqual(output.default_name("  ", "2026-W46"), "Status 2026-W46.pptx")

    def test_free_path(self):
        with tempfile.TemporaryDirectory() as d:
            p = os.path.join(d, "x.pptx")
            self.assertEqual(output.free_path(p), p)
            open(p, "w").close()
            self.assertEqual(output.free_path(p), os.path.join(d, "x (2).pptx"))
            open(os.path.join(d, "x (2).pptx"), "w").close()
            self.assertEqual(output.free_path(p), os.path.join(d, "x (3).pptx"))
            self.assertEqual(output.free_path(p, force=True), p)

    def test_template_path(self):
        inp = os.path.join("C:\\", "work", "Tasks.xlsx") if os.name == "nt" else "/work/Tasks.xlsx"
        self.assertIsNone(output.template_path(inp, "", None))
        self.assertEqual(output.template_path(inp, "", "cli.pptx"), "cli.pptx")
        self.assertEqual(output.template_path(inp, "t/customer.pptx", None), os.path.join(os.path.dirname(inp), "t/customer.pptx"))


class CliDeckTest(unittest.TestCase):
    def setUp(self):
        self.dir = tempfile.mkdtemp()
        self.input = os.path.join(self.dir, "Tasks.xlsx")
        shutil.copy(os.path.join(ROOT, "tests", "fixtures", "excel_filled.xlsx"), self.input)
        self.deck = os.path.join(self.dir, "Falcon Status 2026-W46.pptx")

    def tearDown(self):
        shutil.rmtree(self.dir)

    def run_cli(self, *args) -> int:
        with contextlib.redirect_stdout(io.StringIO()) as out:
            code = main([self.input, *args])
        self.output = out.getvalue()
        return code

    def test_weekly_run_writes_deck_and_snapshot(self):
        self.assertEqual(self.run_cli(), 1)  # the fixture has row errors; the deck is still written
        self.assertTrue(os.path.exists(self.deck))
        self.assertIn("Deck written", self.output)
        self.assertTrue(os.path.exists(os.path.join(self.dir, "history", "2026-W46.json")))
        self.assertEqual(lint_package(opc.Package.open(self.deck)), [])

    def test_second_run_does_not_overwrite(self):
        self.run_cli("--no-save")
        self.run_cli("--no-save")
        self.assertTrue(os.path.exists(os.path.join(self.dir, "Falcon Status 2026-W46 (2).pptx")))
        self.run_cli("--no-save", "--force")
        self.assertFalse(os.path.exists(os.path.join(self.dir, "Falcon Status 2026-W46 (3).pptx")))
        self.assertFalse(os.path.exists(os.path.join(self.dir, "history")))

    def test_out_and_template(self):
        tpl = os.path.join(self.dir, "customer.pptx")
        with open(tpl, "wb") as fh:
            fh.write(Deck.default().to_bytes())
        out = os.path.join(self.dir, "custom.pptx")
        self.assertEqual(self.run_cli("--no-save", "--template", tpl, "--out", out), 1)
        self.assertTrue(os.path.exists(out))
        self.assertEqual(self.run_cli("--no-save", "--template", os.path.join(self.dir, "missing.pptx")), 2)
        self.assertIn("cannot use the template", self.output)

    def test_check_writes_no_deck(self):
        self.run_cli("--check")
        self.assertEqual(os.listdir(self.dir), ["Tasks.xlsx"])


class DemoTest(unittest.TestCase):
    def test_demo_runs_end_to_end(self):
        import make_demo

        with tempfile.TemporaryDirectory() as d:
            folder = os.path.join(d, "demo")
            sys.argv = ["make_demo.py", folder]
            with contextlib.redirect_stdout(io.StringIO()):
                make_demo.main()
                code = main([os.path.join(folder, "Tasks.xlsx")])
            self.assertEqual(code, 0)
            deck = os.path.join(folder, "Project Falcon Status 2026-W46.pptx")
            pkg = opc.Package.open(deck)
            self.assertEqual(lint_package(pkg), [])
            self.assertEqual(len(pkg.related(pkg.main_part(), opc.RT_SLIDE)), 6)
            self.assertEqual(len(os.listdir(os.path.join(folder, "history"))), 7)


if __name__ == "__main__":
    unittest.main()
