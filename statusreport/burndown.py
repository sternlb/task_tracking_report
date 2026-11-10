"""Burndown of the remaining effort over all tasks — R-OUT-10.

Actual line: the open share of the effort of all tasks with an estimate, in %
(remaining hours ÷ total hours of that week), per week with a snapshot (the
same weeks as the progress chart). Hours are never shown to the customer; they
only appear in the console. Ideal line: straight from the first actual value
down to 0 % at the latest Due date; the axis runs on into the future up to that
week (at most 26 weeks ahead).

Snapshots written before 0.9.1 have no stored effort. For them the open share
is derived from the stored task progress and today's estimates.
"""

from __future__ import annotations

import datetime as dt
from dataclasses import dataclass
from typing import Dict, List, Optional, Sequence

from .history import Snapshot, remaining_effort, week_key, week_labels
from .model import Task

MAX_FUTURE_WEEKS = 26


@dataclass
class Burndown:
    """Values in % of the total effort (open share); hours only for the console."""

    keys: List[str]  # ISO weeks on the x-axis, past and future
    labels: List[str]
    actual: List[Optional[float]]  # open effort in % of that week's total; None for future weeks
    ideal: List[Optional[float]]  # %
    total_h: float  # this week, hours (console only, never on a slide)
    remaining_h: float  # this week, hours (console only)
    end: Optional[dt.date]  # latest Due of the tasks with an estimate
    derived_weeks: int = 0  # older snapshots without stored effort

    @property
    def remaining_pct(self) -> float:
        return self.remaining_h / self.total_h * 100 if self.total_h else 0.0


def _monday(key: str) -> dt.date:
    """Monday of an ISO week key such as 2026-W46."""
    year, week = int(key[:4]), int(key[6:])
    jan4 = dt.date(year, 1, 4)
    return jan4 - dt.timedelta(days=jan4.weekday()) + dt.timedelta(weeks=week - 1)


def _pct(remaining: Optional[float], total: Optional[float]) -> Optional[float]:
    if remaining is None or not total:
        return None
    return max(0.0, min(100.0, remaining / total * 100))


def _from_snapshot(snap: Snapshot, estimates: Dict[str, float], current_total: float) -> Optional[float]:
    """Open effort of a past week in % of that week's total effort."""
    if snap.remaining_h is not None:
        return _pct(snap.remaining_h, snap.total_h or current_total)
    known = [(estimates[tid], t.progress) for tid, t in snap.tasks.items() if tid in estimates]
    if not known:
        return None
    total = sum(e for e, _ in known)
    remaining = sum(e * (1 - (p or 0.0) / 100.0) for e, p in known)
    return _pct(remaining, total)


def build(snapshots: Dict[str, Snapshot], current: Snapshot, weeks: int, tasks: Sequence[Task]) -> Optional[Burndown]:
    """The burndown for the report, or None if no task has an estimate."""
    total, remaining = remaining_effort(tasks)
    if total is None:
        return None
    estimates = {t.id: float(t.estimate_h) for t in tasks if t.estimate_h and t.estimate_h > 0}
    dues = [t.due for t in tasks if t.id in estimates and t.due]
    end = max(dues) if dues else None

    merged = {k: s for k, s in snapshots.items() if k < current.week}
    merged[current.week] = current
    past = sorted(merged)[-max(1, weeks):]
    actual: List[Optional[float]] = []
    derived = 0
    for k in past:
        snap = merged[k]
        if k == current.week:
            actual.append(_pct(remaining, total))
            continue
        if snap.remaining_h is None:
            derived += 1
        actual.append(_from_snapshot(snap, estimates, total))

    future: List[str] = []
    if end is not None:
        day = current.report_date + dt.timedelta(weeks=1)
        while week_key(day) <= week_key(end) and len(future) < MAX_FUTURE_WEEKS:
            future.append(week_key(day))
            day += dt.timedelta(weeks=1)
    keys = past + future
    actual += [None] * len(future)

    ideal: List[Optional[float]] = [None] * len(keys)
    first = next((i for i, v in enumerate(actual) if v is not None), None)
    if first is not None and end is not None:
        start_day, end_day = _monday(keys[first]), _monday(week_key(end))
        span = (end_day - start_day).days
        if span > 0:  # no ideal line if the latest Due lies before the first week shown
            for i in range(first, len(keys)):
                day = _monday(keys[i])
                ideal[i] = max(0.0, actual[first] * (1 - (day - start_day).days / span))
                if day >= end_day:
                    break
    return Burndown(keys, week_labels(keys), actual, ideal, total, remaining, end, derived)


__all__ = ["Burndown", "build", "MAX_FUTURE_WEEKS"]
