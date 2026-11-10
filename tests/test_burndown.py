"""Tests for the burndown slide (R-OUT-10): effort in snapshots, burndown data, chart, slide."""

from __future__ import annotations

import datetime as dt
import json
import os
import re
import sys
import tempfile
import unittest

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, ROOT)

from statusreport import burndown, history, labels  # noqa: E402
from statusreport.calc import compute  # noqa: E402
from statusreport.model import Settings, Task  # noqa: E402
from statusreport.pptx import opc  # noqa: E402
from statusreport.pptx.chart import line_chart_xml  # noqa: E402
from statusreport.pptx.deck import Deck  # noqa: E402
from statusreport.pptx.lint import lint_package  # noqa: E402
from statusreport.render import render  # noqa: E402
from statusreport.report import build_report  # noqa: E402

D = dt.date
ON = D(2026, 11, 10)  # Tuesday of 2026-W46


def tasks():
    ts = [
        Task(id="A", title="a", version="1", estimate_h=100, spent_h=50, start=D(2026, 10, 1), due=D(2026, 12, 4)),
        Task(id="B", title="b", version="1", estimate_h=50, state="done", due=D(2026, 11, 20)),
        Task(id="C", title="c", version="2", estimate_h=50, due=D(2026, 12, 18)),  # not started
        Task(id="D", title="d", version="2", progress_manual=30),  # no estimate: not in the burndown
    ]
    compute(ts, Settings(report_date=ON))
    return ts


def snap_at(day, remaining=None, task_progress=None):
    s = history.Snapshot(history.week_key(day), day)
    for tid, p in (task_progress or {}).items():
        s.tasks[tid] = history.TaskSnap("1", tid, p, "green")
    s.remaining_h = remaining
    return s


class EffortTest(unittest.TestCase):
    def test_remaining_effort(self):
        total, remaining = history.remaining_effort(tasks())
        self.assertEqual(total, 200)
        self.assertAlmostEqual(remaining, 100 * 0.5 + 0 + 50)  # A half done, B done, C not started
        self.assertEqual(history.remaining_effort([Task(id="x", title="x", version="1")]), (None, None))

    def test_snapshot_stores_effort(self):
        ts = tasks()
        versions, _ = compute(ts, Settings(report_date=ON))
        snap = history.snapshot_from(versions, ON)
        self.assertEqual((snap.total_h, snap.remaining_h), (200, 100))
        data = json.loads(history.to_json(snap))
        self.assertEqual(data["effort"], {"total_h": 200, "remaining_h": 100})
        back = history.from_json(history.to_json(snap))
        self.assertEqual((back.total_h, back.remaining_h), (200, 100))

    def test_old_snapshot_without_effort(self):
        old = {"schema": 1, "week": "2026-W45", "report_date": "2026-11-03", "versions": {}, "tasks": {}}
        snap = history.from_json(json.dumps(old))
        self.assertIsNone(snap.remaining_h)


class BuildTest(unittest.TestCase):
    def current(self, ts):
        versions, _ = compute(ts, Settings(report_date=ON))
        return history.snapshot_from(versions, ON)

    def test_no_estimates_no_burndown(self):
        ts = [Task(id="x", title="x", version="1", progress_manual=10)]
        self.assertIsNone(burndown.build({}, self.current(ts), 12, ts))

    def test_stored_and_derived_weeks(self):
        ts = tasks()
        snaps = {
            "2026-W44": snap_at(D(2026, 10, 27), remaining=180.0),
            # older format: derived from progress with today's estimates: A 20 % -> 80, B 50 % -> 25, C -> 50
            "2026-W45": snap_at(D(2026, 11, 3), task_progress={"A": 20, "B": 50, "C": None}),
        }
        bd = burndown.build(snaps, self.current(ts), 12, ts)
        self.assertEqual(bd.keys[:3], ["2026-W44", "2026-W45", "2026-W46"])
        self.assertEqual(bd.actual[:3], [180.0, 155.0, 100.0])
        self.assertEqual(bd.derived_weeks, 1)
        self.assertEqual((bd.total_h, bd.remaining_h), (200, 100))

    def test_future_weeks_and_ideal_line(self):
        ts = tasks()
        snaps = {"2026-W44": snap_at(D(2026, 10, 27), remaining=180.0)}
        bd = burndown.build(snaps, self.current(ts), 12, ts)
        self.assertEqual(bd.end, D(2026, 12, 18))  # latest Due of the tasks with an estimate
        # past weeks with a snapshot (W45 has none), then the future up to the week of the latest Due
        self.assertEqual(bd.keys, ["2026-W44", "2026-W46", "2026-W47", "2026-W48", "2026-W49", "2026-W50", "2026-W51"])
        self.assertEqual(bd.actual[:2], [180.0, 100.0])
        self.assertTrue(all(v is None for v in bd.actual[2:]))  # future: no actual values
        self.assertEqual(bd.ideal[0], 180.0)  # starts at the first actual value
        self.assertEqual(bd.ideal[-1], 0.0)  # reaches 0 h in the week of the latest Due
        self.assertEqual(bd.ideal, sorted(bd.ideal, reverse=True))
        self.assertEqual(bd.labels[0], "CW 44")

    def test_future_is_capped(self):
        ts = [Task(id="A", title="a", version="1", estimate_h=10, due=D(2028, 1, 1))]
        bd = burndown.build({}, self.current(ts), 12, ts)
        self.assertEqual(len(bd.keys), 1 + burndown.MAX_FUTURE_WEEKS)

    def test_no_ideal_when_the_project_is_overdue(self):
        ts = [Task(id="A", title="a", version="1", estimate_h=10, due=D(2026, 9, 1))]
        bd = burndown.build({}, self.current(ts), 12, ts)
        self.assertEqual(bd.keys, ["2026-W46"])
        self.assertEqual(bd.ideal, [None])


class ChartTest(unittest.TestCase):
    def test_auto_axis_hours_and_dashed_series(self):
        xml = line_chart_xml(["CW 1", "CW 2"], [("Remaining", [10, 5]), ("Ideal", [10, 0])], "rId1",
                             y_max=None, y_step=None, value_format='#,##0" h"',
                             styles=[{"color": "accent1"}, {"color": "#7F7F7F", "dash": True, "marker": False}]).decode()
        self.assertNotIn("<c:max ", xml)
        self.assertNotIn("<c:majorUnit", xml)
        self.assertIn('formatCode="#,##0&quot; h&quot;"', xml)
        self.assertEqual(xml.count('<a:prstDash val="dash"/>'), 1)
        self.assertIn('<c:symbol val="none"/>', xml)

    def test_percent_chart_unchanged(self):
        xml = line_chart_xml(["CW 1"], [("1.0", [10])], "rId1").decode()
        self.assertIn('<c:max val="100"/>', xml)
        self.assertIn('<c:majorUnit val="20"/>', xml)


class SlideTest(unittest.TestCase):
    def build(self, ts):
        s = Settings(project="P", report_date=ON)
        versions, _ = compute(ts, s)
        cur = history.snapshot_from(versions, ON)
        bd = burndown.build({}, cur, 12, ts)
        r = build_report(s, versions, None, cur.week, ["CW 46"], [(v.name, [v.progress]) for v in versions], bd)
        d = render(r, Deck.default(), s.title_layout, s.content_layout, s.rows_per_slide)
        return r, opc.Package.from_bytes(d.to_bytes())

    def test_burndown_is_the_third_slide(self):
        r, pkg = self.build(tasks())
        self.assertEqual(lint_package(pkg), [])
        slides = pkg.related(pkg.main_part(), opc.RT_SLIDE)
        third = pkg.parts[slides[2]].decode()
        self.assertIn(labels.BURNDOWN_TITLE, third)
        self.assertIn("100 h of 200 h remaining", third)
        self.assertIn("(18 Dec 2026)", third)
        chart = pkg.related(slides[2], opc.RT_CHART)[0]
        self.assertIn(labels.BURNDOWN_IDEAL, pkg.parts[chart].decode())
        self.assertIn("100 h of 200 h", r.burndown_note)

    def test_no_burndown_slide_without_estimates(self):
        ts = [Task(id="x", title="x", version="1", progress_manual=10)]
        _, pkg = self.build(ts)
        texts = "".join(re.findall(r"<a:t>([^<]*)</a:t>", "".join(pkg.parts[p].decode() for p in pkg.related(pkg.main_part(), opc.RT_SLIDE))))
        self.assertNotIn(labels.BURNDOWN_TITLE, texts)


if __name__ == "__main__":
    unittest.main()
