"""Tests for M3: progress, expected %, traffic lights, version aggregation (R-CA-1 … R-CA-5)."""

from __future__ import annotations

import datetime as dt
import os
import sys
import unittest

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, ROOT)

from statusreport.calc import (  # noqa: E402
    compute, compute_task, compute_version, derive_start, expected_progress, light_for, task_progress, version_order,
)
from statusreport.loader import load  # noqa: E402
from statusreport.model import Settings, Task  # noqa: E402

D = dt.date
DONE = ["done", "closed"]
ON = D(2026, 11, 10)


def task(**kw) -> Task:
    base = dict(id="T", title="t", version="1.0")
    base.update(kw)
    return Task(**base)


class ProgressTest(unittest.TestCase):
    """R-CA-1: first rule that applies."""

    CASES = [
        # (fields, expected progress, why)
        (dict(progress_manual=40, state="done", estimate_h=10, spent_h=9), 40, "manual wins over everything"),
        (dict(progress_manual=0), 0, "manual 0 is a value"),
        (dict(state="Done", estimate_h=10, spent_h=2), 100, "done state, case-insensitive"),
        (dict(state=" closed "), 100, "done state, trimmed"),
        (dict(state="in progress", estimate_h=40, spent_h=10), 25, "spent / estimate"),
        (dict(estimate_h=10, spent_h=0), 0, "nothing spent yet"),
        (dict(estimate_h=10, spent_h=10), 99, "hours alone stop at 99 %"),
        (dict(estimate_h=10, spent_h=25), 99, "overrun stops at 99 %"),
        (dict(estimate_h=0, spent_h=5), None, "estimate 0 cannot be used"),
        (dict(estimate_h=10), None, "spent missing"),
        (dict(spent_h=10), None, "estimate missing"),
        (dict(), None, "nothing known"),
    ]

    def test_cases(self):
        for fields, expected, why in self.CASES:
            got = task_progress(task(**fields), DONE)
            if expected is None:
                self.assertIsNone(got, why)
            else:
                self.assertAlmostEqual(got, expected, msg=why)


class ExpectedTest(unittest.TestCase):
    """R-CA-2: elapsed share of Start..Due, calendar days, limited to 0-100."""

    def test_cases(self):
        s, d = D(2026, 11, 1), D(2026, 11, 21)  # 20 days
        cases = [
            (s, d, D(2026, 11, 11), 50), (s, d, s, 0), (s, d, d, 100),
            (s, d, D(2026, 10, 1), 0), (s, d, D(2026, 12, 24), 100),
            (s, s, D(2026, 10, 31), 0), (s, s, s, 100),
        ]
        for start, due, on, expected in cases:
            self.assertAlmostEqual(expected_progress(start, due, on), expected, msg=(start, due, on))
        self.assertIsNone(expected_progress(None, d, ON))
        self.assertIsNone(expected_progress(s, None, ON))
        self.assertIsNone(expected_progress(d, s, ON), "Start after Due is no plan")


class DeriveStartTest(unittest.TestCase):
    """R-CA-6: empty Start = Due minus the estimate in working days (Mon-Fri)."""

    FRI = D(2026, 11, 13)

    def test_cases(self):
        cases = [
            (self.FRI, 24, D(2026, 11, 11), "3 days ending Friday start Wednesday"),
            (self.FRI, 8, self.FRI, "1 day = the due day itself"),
            (self.FRI, 4.5, self.FRI, "part of a day rounds up to 1 day"),
            (self.FRI, 9, D(2026, 11, 12), "1.1 days round up to 2"),
            (self.FRI, 40, D(2026, 11, 9), "a week ending Friday starts Monday"),
            (self.FRI, 56, D(2026, 11, 5), "7 days skip the weekend"),
            (D(2026, 11, 15), 8, self.FRI, "Due on Sunday counts back from Friday"),
            (D(2026, 11, 16), 16, self.FRI, "2 days ending Monday start Friday"),
        ]
        for due, est, start, why in cases:
            self.assertEqual(derive_start(due, est, 8), start, why)
        self.assertEqual(derive_start(self.FRI, 15, 7.5), D(2026, 11, 12), "uses hours per day")

    def test_no_derivation_without_due_or_estimate(self):
        self.assertIsNone(derive_start(None, 24, 8))
        self.assertIsNone(derive_start(self.FRI, None, 8))
        self.assertIsNone(derive_start(self.FRI, 0, 8))

    def test_compute_task_fills_only_an_empty_start(self):
        s = Settings(report_date=D(2026, 11, 12))
        t = task(due=self.FRI, estimate_h=24, spent_h=12)
        compute_task(t, s, s.report_date)
        self.assertEqual((t.start, t.start_derived), (D(2026, 11, 11), True))
        self.assertAlmostEqual(t.expected, 50)  # 1 of 2 calendar days
        manual = task(start=D(2026, 11, 2), due=self.FRI, estimate_h=24)
        compute_task(manual, s, s.report_date)
        self.assertEqual((manual.start, manual.start_derived), (D(2026, 11, 2), False))


class LightTest(unittest.TestCase):
    """R-CA-3 with the default thresholds 10 / 25 pp."""

    FUTURE = D(2026, 12, 31)
    PAST = D(2026, 11, 9)

    CASES = [
        # (progress, expected, due, override, light, why)
        (10, 90, FUTURE, "green", "green", "override wins"),
        (100, 10, FUTURE, "red", "red", "override wins even over 100 %"),
        (100, None, PAST, None, "green", "done beats overdue"),
        (99, 100, PAST, None, "red", "overdue and not done"),
        (None, None, PAST, None, "red", "overdue with unknown progress"),
        (None, 50, FUTURE, None, "grey", "progress missing"),
        (50, None, FUTURE, None, "grey", "plan missing"),
        (None, None, None, None, "grey", "nothing known"),
        (80, 50, FUTURE, None, "green", "ahead of plan"),
        (40, 50, FUTURE, None, "green", "10 pp behind is still green (limit included)"),
        (39.9, 50, FUTURE, None, "orange", "just over 10 pp"),
        (25, 50, FUTURE, None, "orange", "25 pp behind is still orange (limit included)"),
        (24.9, 50, FUTURE, None, "red", "just over 25 pp"),
        (50, 50, ON, None, "green", "due today is not overdue"),
    ]

    def test_cases(self):
        for progress, expected, due, override, light, why in self.CASES:
            self.assertEqual(light_for(progress, expected, due, ON, 10, 25, override), light, why)

    def test_float_noise_at_the_limit(self):
        # 100/3 - 23.333... = 10.000000000000004 must not flip to orange
        self.assertEqual(light_for(100 / 3 - 10, 100 / 3, None, ON, 10, 25), "green")

    def test_custom_thresholds(self):
        self.assertEqual(light_for(40, 50, None, ON, 5, 15), "orange")
        self.assertEqual(light_for(30, 50, None, ON, 5, 15), "red")


class VersionTest(unittest.TestCase):
    def settings(self, **kw) -> Settings:
        s = Settings(report_date=D(2026, 11, 11), done_states=DONE)
        for k, v in kw.items():
            setattr(s, k, v)
        return s

    def run_compute(self, tasks, **kw):
        return compute(tasks, self.settings(**kw))

    def test_weighted_by_estimate(self):
        a = task(id="A", estimate_h=30, spent_h=15)  # 50 %
        b = task(id="B", estimate_h=10, spent_h=9)  # 90 %
        (vs,), issues = self.run_compute([a, b])
        self.assertEqual(vs.weighting, "estimate")
        self.assertAlmostEqual(vs.progress, (50 * 30 + 90 * 10) / 40)
        self.assertEqual(issues, [])

    def test_plain_average_when_an_estimate_is_missing(self):
        a = task(id="A", estimate_h=30, spent_h=15)  # 50 %
        b = task(id="B", progress_manual=90)  # no estimate
        (vs,), _ = self.run_compute([a, b])
        self.assertEqual(vs.weighting, "equal")
        self.assertAlmostEqual(vs.progress, 70)

    def test_tasks_without_progress_are_left_out_with_a_warning(self):
        a = task(id="A", progress_manual=60, estimate_h=10)
        b = task(id="B", estimate_h=50, sheet="Tasks", row=3)
        (vs,), issues = self.run_compute([a, b])
        self.assertAlmostEqual(vs.progress, 60)
        self.assertEqual(vs.uncounted, 1)
        self.assertEqual(len(issues), 1)
        self.assertIn("B: no progress", issues[0].message)
        self.assertEqual((issues[0].sheet, issues[0].row), ("Tasks", 3))

    def test_version_light_compares_like_with_like(self):
        # A is on plan, B has no dates: the version light must only use A
        a = task(id="A", progress_manual=50, start=D(2026, 11, 1), due=D(2026, 11, 21))
        b = task(id="B", progress_manual=0)
        (vs,), _ = self.run_compute([a, b])
        self.assertAlmostEqual(vs.progress, 25)  # shown value: all tasks with progress
        self.assertAlmostEqual(vs.expected, 50)  # only A has a plan
        self.assertEqual(vs.light, "green")  # A alone: 50 vs 50
        self.assertEqual((vs.start, vs.due), (D(2026, 11, 1), D(2026, 11, 21)))

    def test_version_light_behind_and_overdue(self):
        late = task(id="A", progress_manual=10, start=D(2026, 11, 1), due=D(2026, 11, 21))  # 50 expected
        (vs,), _ = self.run_compute([late])
        self.assertEqual(vs.light, "red")
        overdue = task(id="A", progress_manual=90, start=D(2026, 10, 1), due=D(2026, 11, 5))
        (vs,), _ = self.run_compute([overdue])
        self.assertEqual(vs.light, "red")
        done = task(id="A", state="done", due=D(2026, 11, 5))
        (vs,), _ = self.run_compute([done])
        self.assertEqual((vs.progress, vs.light), (100, "green"))

    def test_version_light_ignores_task_overrides(self):
        a = task(id="A", progress_manual=50, start=D(2026, 11, 1), due=D(2026, 11, 21), light_override="red")
        (vs,), _ = self.run_compute([a])
        self.assertEqual(a.light, "red")
        self.assertEqual(vs.light, "green")

    def test_empty_version_is_grey(self):
        vs = compute_version("9.9", [task(id="A")], self.settings(), D(2026, 11, 11))
        self.assertIsNone(vs.progress)
        self.assertEqual(vs.light, "grey")

    def test_version_order(self):
        names = ["2.0.0", "1.10.0", "1.2.0", "Backlog"]
        self.assertEqual(version_order(names, []), ["1.2.0", "1.10.0", "2.0.0", "Backlog"])
        self.assertEqual(version_order(names, ["backlog", "2.0.0"]), ["Backlog", "2.0.0", "1.2.0", "1.10.0"])

    def test_task_order_inside_a_version_is_input_order(self):
        tasks = [task(id=i) for i in ("Z", "A", "M")]
        (vs,), _ = self.run_compute(tasks)
        self.assertEqual([t.id for t in vs.tasks], ["Z", "A", "M"])


class FixtureTest(unittest.TestCase):
    """End to end: workbook saved by Excel -> loader -> calc, report date 2026-11-10."""

    def test_excel_filled(self):
        result = load(os.path.join(ROOT, "tests", "fixtures", "excel_filled.xlsx"))
        versions, issues = compute(result.tasks, result.settings)
        self.assertEqual([v.name for v in versions], ["1.2.0", "1.10.0"])
        v12 = versions[0]
        tasks = {t.id: t for t in v12.tasks}
        self.assertAlmostEqual(tasks["PRJ-101"].progress, 10 / 24 * 100)
        self.assertAlmostEqual(tasks["PRJ-101"].expected, 8 / 25 * 100)
        self.assertEqual(tasks["PRJ-101"].light, "green")
        self.assertAlmostEqual(tasks["PRJ-102"].expected, 40 / 43 * 100)
        self.assertEqual(tasks["PRJ-102"].light, "orange")  # 80 vs 93 %
        self.assertEqual(tasks["PRJ-104"].light, "red")  # override
        self.assertEqual((tasks["PRJ-106"].progress, tasks["PRJ-106"].light), (100, "green"))
        self.assertEqual(v12.weighting, "equal")  # PRJ-106 has no estimate
        self.assertAlmostEqual(v12.progress, (10 / 24 * 100 + 80 + 100) / 3)
        self.assertEqual(v12.light, "green")
        self.assertEqual(versions[1].light, "grey")  # PRJ-103 has no dates
        self.assertEqual([i.message.split(":")[0] for i in issues], ["PRJ-104"])

    def test_import_without_start_dates(self):
        """Polarion has no start field: Start is derived from Due and the estimate (R-CA-6)."""
        result = load(os.path.join(ROOT, "tests", "fixtures", "excel_with_import.xlsx"))
        versions, _ = compute(result.tasks, result.settings)
        tasks = {t.id: t for v in versions for t in v.tasks}
        self.assertFalse(tasks["PRJ-101"].start_derived)  # Start typed in Tasks
        self.assertEqual((tasks["PRJ-201"].start, tasks["PRJ-201"].start_derived), (D(2026, 10, 22), True))  # 1w 2d
        self.assertEqual(tasks["PRJ-203"].start, D(2026, 12, 3))  # 2d ending Fri 4 Dec
        self.assertEqual(tasks["PRJ-203"].light, "green")  # not started yet by plan, 38 % done
        self.assertIsNone(tasks["PRJ-301"].start)  # no Due, no estimate
        self.assertEqual(versions[1].light, "green")


if __name__ == "__main__":
    unittest.main()
