"""Tests for M7: the packaged tool, Python 3.9 compatibility, performance (R-OP-1/2, C-1, NF-1)."""

from __future__ import annotations

import ast
import contextlib
import datetime as dt
import io
import os
import re
import shutil
import subprocess
import sys
import tempfile
import time
import unittest
import zipfile

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, ROOT)
sys.path.insert(0, os.path.join(ROOT, "tools"))

from statusreport import __version__, history, labels  # noqa: E402
from statusreport.__main__ import main  # noqa: E402
from statusreport.model import Task, VersionStatus  # noqa: E402
from statusreport.xlsx_write import Sheet, workbook_bytes  # noqa: E402

PACKAGE = os.path.join(ROOT, "statusreport")
STDLIB = set(getattr(sys, "stdlib_module_names", ())) or None  # Python 3.10+


def sources():
    for folder, _, files in os.walk(PACKAGE):
        for name in files:
            if name.endswith(".py"):
                path = os.path.join(folder, name)
                with open(path, encoding="utf-8") as fh:
                    yield os.path.relpath(path, ROOT), fh.read()


class CompatibilityTest(unittest.TestCase):
    """C-1: standard library only, Python 3.9 syntax (checked without a 3.9 interpreter)."""

    def test_python39_grammar(self):
        for name, text in sources():
            with self.subTest(name):
                ast.parse(text, filename=name, feature_version=(3, 9))

    def test_future_annotations_where_annotations_are_used(self):
        for name, text in sources():
            if re.search(r"^\s*def .*\)\s*->|^\s*\w+: ", text, re.M):
                self.assertIn("from __future__ import annotations", text, name)

    def test_no_apis_newer_than_39(self):
        newer = [r"\bzip\([^)]*strict=", r"\bpairwise\b", r"\bbit_count\(", r"slots=True", r"kw_only=",
                 r"\btomllib\b", r"datetime\.UTC\b", r"\bExceptionGroup\b", r"typing import .*\bSelf\b",
                 r"\bTypeAlias\b", r"\bParamSpec\b", r"\baiter\(", r"\banext\("]
        for name, text in sources():
            for pattern in newer:
                self.assertIsNone(re.search(pattern, text), f"{name}: {pattern}")

    def test_only_standard_library_imports(self):
        if STDLIB is None:
            self.skipTest("needs sys.stdlib_module_names (Python 3.10+)")
        for name, text in sources():
            for node in ast.walk(ast.parse(text)):
                if isinstance(node, ast.Import):
                    mods = [a.name for a in node.names]
                elif isinstance(node, ast.ImportFrom) and node.level == 0:
                    mods = [node.module or ""]
                else:
                    continue
                for mod in mods:
                    top = mod.split(".")[0]
                    self.assertTrue(top in STDLIB or top == "statusreport" or top == "__future__", f"{name}: import {mod}")


class DeliveryTest(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        import build_pyz

        cls.tmp = tempfile.mkdtemp()
        with contextlib.redirect_stdout(io.StringIO()):
            cls.archive = build_pyz.main(os.path.join(cls.tmp, "dist"))
        cls.folder = os.path.join(cls.tmp, "dist", "sw-status-report")
        cls.pyz = os.path.join(cls.folder, "statusreport.pyz")

    @classmethod
    def tearDownClass(cls):
        shutil.rmtree(cls.tmp)

    def run_pyz(self, *args, cwd=None):
        return subprocess.run(
            [sys.executable, "-I", self.pyz, *args], cwd=cwd or self.tmp,
            capture_output=True, text=True, encoding="utf-8", errors="replace", timeout=120,
        )

    def test_contents(self):
        self.assertEqual(
            sorted(os.listdir(self.folder)),
            ["README.md", "Tasks.xlsx", "Tutorial (Deutsch).html", "make_report.cmd", "statusreport.pyz"],
        )
        with open(os.path.join(self.folder, "Tutorial (Deutsch).html"), encoding="utf-8") as fh:
            tutorial = fh.read()
        self.assertTrue(tutorial.startswith("<!doctype html>"))
        self.assertNotIn('src="img/', tutorial)  # images embedded, works offline
        self.assertNotIn('="video/', tutorial)
        self.assertEqual(tutorial.count("data:image/png;base64,"), 6)  # 5 slides + video poster
        self.assertEqual(tutorial.count("data:video/mp4;base64,"), 1)
        with zipfile.ZipFile(self.pyz) as zf:
            names = zf.namelist()
        self.assertIn("__main__.py", names)
        self.assertIn("statusreport/pptx/base.pptx", names)
        self.assertFalse([n for n in names if "__pycache__" in n or n.endswith(".pyc")])
        self.assertTrue(self.archive.endswith(f"sw-status-report-{__version__}.zip"))
        with zipfile.ZipFile(self.archive) as zf:
            self.assertIn("sw-status-report/statusreport.pyz", zf.namelist())

    def test_pyz_runs_isolated_from_the_sources(self):
        r = self.run_pyz("--version")
        self.assertEqual((r.returncode, r.stdout.strip()), (0, f"statusreport {__version__}"))
        work = os.path.join(self.tmp, "work")
        os.makedirs(work, exist_ok=True)
        shutil.copy(os.path.join(ROOT, "tests", "fixtures", "excel_filled.xlsx"), os.path.join(work, "Tasks.xlsx"))
        r = self.run_pyz(os.path.join(work, "Tasks.xlsx"), cwd=work)
        self.assertEqual(r.returncode, 1, r.stdout)  # fixture has row errors; the deck is still written
        self.assertIn("Deck written", r.stdout)
        self.assertTrue(os.path.exists(os.path.join(work, "Falcon Status 2026-W46.pptx")))
        self.assertTrue(os.path.exists(os.path.join(work, "history", "2026-W46.json")))

    def test_bundled_template_workbook_is_the_init_template(self):
        from statusreport.template import template_bytes
        from statusreport.xlsx_read import Workbook

        a = Workbook.open(os.path.join(self.folder, "Tasks.xlsx"))
        b = Workbook(template_bytes())
        self.assertEqual([a.rows(s) for s in a.sheet_names], [b.rows(s) for s in b.sheet_names])

    @unittest.skipUnless(os.name == "nt", "make_report.cmd is Windows only")
    def test_make_report_cmd(self):
        work = os.path.join(self.tmp, "cmd")
        os.makedirs(work, exist_ok=True)
        shutil.copy(os.path.join(ROOT, "tests", "fixtures", "excel_with_import.xlsx"), os.path.join(work, "Tasks.xlsx"))
        env = dict(os.environ, SWSR_NO_PAUSE="1")
        cmd = os.path.join(self.folder, "make_report.cmd")
        r = subprocess.run(["cmd", "/c", cmd, os.path.join(work, "Tasks.xlsx")], capture_output=True, text=True,
                           encoding="utf-8", errors="replace", env=env, timeout=120)
        self.assertEqual(r.returncode, 0, r.stdout + r.stderr)
        self.assertIn("Done.", r.stdout)
        r = subprocess.run(["cmd", "/c", cmd], capture_output=True, text=True, encoding="utf-8", errors="replace", env=env, timeout=60)
        self.assertEqual(r.returncode, 2)
        self.assertIn("Drag your input workbook", r.stdout)


class PerformanceTest(unittest.TestCase):
    """NF-1: 500 tasks and 52 weeks of history in under 5 s."""

    def test_500_tasks_52_weeks(self):
        with tempfile.TemporaryDirectory() as d:
            header = [f.header for f in labels.FIELDS]
            rows = [header]
            start = dt.date(2026, 9, 1)
            for i in range(500):
                rows.append(task_row(i, start))
            settings = [list(labels.SETTINGS_HEADERS), ["Project", "Load"], ["Report date", dt.date(2026, 11, 10)]]
            path = os.path.join(d, "Tasks.xlsx")
            with open(path, "wb") as fh:
                fh.write(workbook_bytes([Sheet(labels.SHEET_TASKS, rows), Sheet(labels.SHEET_SETTINGS, settings)]))
            hist = os.path.join(d, "history")
            for w in range(52):
                day = dt.date(2026, 11, 3) - dt.timedelta(weeks=w)
                tasks = [Task(id=f"L-{i:03d}", title="t", version=f"{i % 5}.0", progress=float(w), light="green") for i in range(500)]
                vss = [VersionStatus(f"{v}.0", [t for t in tasks if t.version == f"{v}.0"], progress=float(w)) for v in range(5)]
                history.save(hist, history.snapshot_from(vss, day))
            t0 = time.perf_counter()
            with contextlib.redirect_stdout(io.StringIO()):
                code = main([path, "--no-save"])
            elapsed = time.perf_counter() - t0
            self.assertEqual(code, 0)
            self.assertLess(elapsed, 5.0, f"{elapsed:.2f} s")


def task_row(i: int, start: dt.date) -> list:
    values = {
        "id": f"L-{i:03d}", "title": f"Load test task {i}", "version": f"{i % 5}.0", "owner": "x",
        "start": start, "due": start + dt.timedelta(days=60 + i % 30), "estimate_h": 40, "spent_h": i % 40,
        "state": "done" if i % 17 == 0 else "in progress", "comment": "c" * (i % 50),
    }
    return [values.get(f.key) for f in labels.FIELDS]


if __name__ == "__main__":
    unittest.main()
