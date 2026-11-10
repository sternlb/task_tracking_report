"""Tests for M4: weekly snapshots, previous week, trend data (R-HI-1 … R-HI-4)."""

from __future__ import annotations

import contextlib
import datetime as dt
import io
import json
import os
import shutil
import sys
import tempfile
import unittest

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, ROOT)

from statusreport import history  # noqa: E402
from statusreport.__main__ import main  # noqa: E402
from statusreport.model import Task, VersionStatus  # noqa: E402

D = dt.date


def snap(week_day: dt.date, **versions) -> history.Snapshot:
    s = history.Snapshot(history.week_key(week_day), week_day)
    for name, progress in versions.items():  # keyword v1_2 -> version "1.2"
        s.versions[name[1:].replace("_", ".")] = history.VersionSnap(progress, None, "green")
    return s


class WeekTest(unittest.TestCase):
    def test_week_key(self):
        cases = {
            D(2026, 11, 10): "2026-W46", D(2026, 1, 1): "2026-W01", D(2026, 12, 31): "2026-W53",
            D(2027, 1, 3): "2026-W53", D(2027, 1, 4): "2027-W01", D(2021, 1, 3): "2020-W53",
        }
        for day, key in cases.items():
            self.assertEqual(history.week_key(day), key, day)

    def test_labels(self):
        self.assertEqual(history.week_labels(["2026-W45", "2026-W46"]), ["CW 45", "CW 46"])
        self.assertEqual(history.week_labels(["2026-W53", "2027-W01"]), ["CW 53/26", "CW 01/27"])

    def test_history_dir_is_next_to_the_input(self):
        path = os.path.join("some", "dir", "Tasks.xlsx")
        self.assertEqual(history.history_dir(path), os.path.join(os.path.abspath(os.path.join("some", "dir")), "history"))


class SnapshotFileTest(unittest.TestCase):
    def setUp(self):
        self.dir = tempfile.mkdtemp()

    def tearDown(self):
        shutil.rmtree(self.dir)

    def make(self) -> history.Snapshot:
        t = Task(id="PRJ-1", title="Login", version="1.2.0", progress=41.66666, light="green")
        vs = VersionStatus("1.2.0", [t], progress=41.66666, expected=32.0, light="green")
        return history.snapshot_from([vs], D(2026, 11, 10))

    def test_roundtrip(self):
        s = self.make()
        path = history.save(self.dir, s, dt.datetime(2026, 11, 10, 9, 12, 44))
        self.assertTrue(path.endswith("2026-W46.json"))
        with open(path, encoding="utf-8") as fh:
            data = json.load(fh)
        self.assertEqual(data["schema"], 1)
        self.assertEqual(data["generated_at"], "2026-11-10T09:12:44")
        self.assertEqual(data["tasks"]["PRJ-1"], {"version": "1.2.0", "title": "Login", "progress": 41.67, "light": "green"})
        self.assertEqual(data["versions"]["1.2.0"], {"progress": 41.67, "expected": 32.0, "light": "green"})
        loaded = history.load_all(self.dir)["2026-W46"]
        self.assertEqual(loaded.report_date, D(2026, 11, 10))
        self.assertEqual(loaded.tasks["PRJ-1"].progress, 41.67)

    def test_same_week_is_replaced_and_no_temp_files_remain(self):
        s = self.make()
        history.save(self.dir, s)
        s.versions["1.2.0"].progress = 99.0
        history.save(self.dir, s)
        self.assertEqual(os.listdir(self.dir), ["2026-W46.json"])
        self.assertEqual(history.load_all(self.dir)["2026-W46"].versions["1.2.0"].progress, 99.0)

    def test_broken_and_foreign_files(self):
        history.save(self.dir, self.make())
        for name, text in (
            ("2026-W45.json", "{not json"),
            ("2026-W44.json", json.dumps({"schema": 99, "week": "2026-W44", "report_date": "2026-10-27"})),
            ("2026-W43.json", json.dumps({"schema": 1, "week": "2026-W40", "report_date": "2026-10-20"})),
            ("notes.json", "{}"),
            ("2026-W42.json.bak", "x"),
        ):
            with open(os.path.join(self.dir, name), "w", encoding="utf-8") as fh:
                fh.write(text)
        issues = []
        snaps = history.load_all(self.dir, issues)
        self.assertEqual(list(snaps), ["2026-W46"])
        self.assertEqual(sorted(i.message.split(" ")[0] for i in issues), ["2026-W43.json", "2026-W44.json", "2026-W45.json"])
        self.assertTrue(all(i.level == "warning" for i in issues))

    def test_missing_folder(self):
        self.assertEqual(history.load_all(os.path.join(self.dir, "nope")), {})


class QueryTest(unittest.TestCase):
    def setUp(self):
        self.snaps = {
            s.week: s
            for s in (
                snap(D(2026, 10, 20), v1_2=10),  # W43
                snap(D(2026, 10, 27), v1_2=20, v2_0=0),  # W44
                snap(D(2026, 11, 10), v1_2=40, v2_0=5),  # W46 (W45 skipped)
                snap(D(2026, 11, 24), v1_2=70),  # W48, later than the report below
            )
        }

    def test_previous_skips_missing_weeks(self):
        self.assertEqual(history.previous(self.snaps, "2026-W47").week, "2026-W46")
        self.assertEqual(history.previous(self.snaps, "2026-W46").week, "2026-W44")
        self.assertEqual(history.previous(self.snaps, "2026-W44").week, "2026-W43")
        self.assertIsNone(history.previous(self.snaps, "2026-W43"))
        self.assertIsNone(history.previous({}, "2026-W46"))

    def test_trend(self):
        current = snap(D(2026, 11, 11), v1_2=45, v3_0=1)  # W46 again, in memory
        keys, labels, series = history.trend(self.snaps, current, 12, ["1.2", "2.0", "3.0"])
        self.assertEqual(keys, ["2026-W43", "2026-W44", "2026-W46"])  # W48 is after the report week
        self.assertEqual(labels, ["CW 43", "CW 44", "CW 46"])
        self.assertEqual(dict(series), {"1.2": [10, 20, 45], "2.0": [None, 0, None], "3.0": [None, None, 1]})

    def test_trend_limits_the_number_of_weeks(self):
        current = snap(D(2026, 11, 17), v1_2=50)  # W47
        keys, _, series = history.trend(self.snaps, current, 2, ["1.2"])
        self.assertEqual(keys, ["2026-W46", "2026-W47"])
        self.assertEqual(series, [("1.2", [40, 50])])

    def test_first_report_has_one_week(self):
        keys, labels, series = history.trend({}, snap(D(2026, 11, 10), v1_2=5), 12, ["1.2"])
        self.assertEqual((keys, labels, series), (["2026-W46"], ["CW 46"], [("1.2", [5])]))


class CliHistoryTest(unittest.TestCase):
    def setUp(self):
        self.dir = tempfile.mkdtemp()
        self.input = os.path.join(self.dir, "Tasks.xlsx")
        shutil.copy(os.path.join(ROOT, "tests", "fixtures", "excel_filled.xlsx"), self.input)
        self.hist = os.path.join(self.dir, "history")

    def tearDown(self):
        shutil.rmtree(self.dir)

    def run_cli(self, *args) -> str:
        with contextlib.redirect_stdout(io.StringIO()) as out:
            main([self.input, *args])
        return out.getvalue()

    def files(self):
        return sorted(os.listdir(self.hist)) if os.path.isdir(self.hist) else []

    def test_weekly_runs(self):
        out = self.run_cli("--date", "2026-10-27")
        self.assertIn("no previous week yet", out)
        self.assertIn("Snapshot saved", out)
        self.run_cli("--date", "2026-11-03")
        out = self.run_cli("--date", "2026-11-04")
        self.assertIn("Snapshot replaced", out)
        self.assertEqual(self.files(), ["2026-W44.json", "2026-W45.json"])
        out = self.run_cli("--date", "2026-11-10", "--check")
        self.assertIn("previous week 2026-W45 (2026-11-04)", out)
        self.assertIn("trend CW 44 .. CW 46 (3 week(s))", out)
        self.assertIn("orange (was green)", out)  # PRJ-102 slips as Due comes closer
        self.assertIn("--check: nothing written", out)
        self.assertEqual(self.files(), ["2026-W44.json", "2026-W45.json"])

    def test_no_save(self):
        out = self.run_cli("--no-save")
        self.assertIn("--no-save: snapshot not stored", out)
        self.assertEqual(self.files(), [])

    def test_check_writes_nothing(self):
        self.run_cli("--check")
        self.assertFalse(os.path.exists(self.hist))


if __name__ == "__main__":
    unittest.main()
