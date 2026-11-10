"""Progress, expected progress and traffic lights — R-CA-1 … R-CA-5, design 6.

All values are percent as float (0–100); rounding happens only for display.
Thresholds compare percentage points and include the limit: with the orange
threshold at 10, a gap of exactly 10 is still green.
"""

from __future__ import annotations

import datetime as dt
from typing import Iterable, List, Optional, Sequence, Tuple

from . import parse
from .model import Issue, Settings, Task, VersionStatus

GREEN, ORANGE, RED, GREY = "green", "orange", "red", "grey"
MAX_FROM_HOURS = 99.0  # hours alone never make a task 100 % (R-CA-1.3)
_EPS = 1e-6  # absorbs float noise at the threshold limits


# ---- single task ----------------------------------------------------------------------------------


def task_progress(task: Task, done_states: Iterable[str]) -> Optional[float]:
    """R-CA-1: manual % -> done state -> spent / estimate (max 99 %) -> no value."""
    if task.progress_manual is not None:
        return task.progress_manual
    if task.state and task.state.strip().lower() in {s.lower() for s in done_states}:
        return 100.0
    if task.estimate_h and task.estimate_h > 0 and task.spent_h is not None:
        return min(task.spent_h / task.estimate_h * 100.0, MAX_FROM_HOURS)
    return None


def expected_progress(start: Optional[dt.date], due: Optional[dt.date], on: dt.date) -> Optional[float]:
    """R-CA-2: elapsed share of the planned time on ``on``, limited to 0–100 (calendar days)."""
    if start is None or due is None or start > due:
        return None
    if due == start:
        return 0.0 if on < start else 100.0
    share = (on - start).days / (due - start).days * 100.0
    return max(0.0, min(100.0, share))


def light_for(
    progress: Optional[float],
    expected: Optional[float],
    due: Optional[dt.date],
    on: dt.date,
    orange_pp: float,
    red_pp: float,
    override: Optional[str] = None,
) -> str:
    """R-CA-3, first rule that applies.

    1. override  2. 100 % -> green  3. past Due and not 100 % -> red (also when the
    progress is unknown: an overdue task that is not reported done is not on time)
    4. plan or progress missing -> grey  5. gap = expected - progress against the thresholds.
    """
    if override:
        return override
    if progress is not None and progress >= 100.0 - _EPS:
        return GREEN
    if due is not None and on > due:
        return RED
    if progress is None or expected is None:
        return GREY
    gap = expected - progress
    if gap <= orange_pp + _EPS:
        return GREEN
    if gap <= red_pp + _EPS:
        return ORANGE
    return RED


def compute_task(task: Task, settings: Settings, on: dt.date) -> None:
    task.progress = task_progress(task, settings.done_states)
    task.expected = expected_progress(task.start, task.due, on)
    task.light = light_for(
        task.progress, task.expected, task.due, on, settings.orange_pp, settings.red_pp, task.light_override
    )


# ---- version --------------------------------------------------------------------------------------


def _weighted(pairs: Sequence[Tuple[float, float]]) -> Optional[float]:
    """Weighted mean of (value, weight) pairs."""
    total = sum(w for _, w in pairs)
    if not pairs or total <= 0:
        return None
    return sum(v * w for v, w in pairs) / total


def compute_version(name: str, tasks: List[Task], settings: Settings, on: dt.date) -> VersionStatus:
    """R-CA-4 / R-CA-5. Expects the tasks to be computed already (:func:`compute_task`)."""
    vs = VersionStatus(name, tasks)
    starts = [t.start for t in tasks if t.start]
    dues = [t.due for t in tasks if t.due]
    vs.start = min(starts) if starts else None
    vs.due = max(dues) if dues else None

    counted = [t for t in tasks if t.progress is not None]
    vs.uncounted = len(tasks) - len(counted)
    by_estimate = all(t.estimate_h is not None and t.estimate_h > 0 for t in counted)
    vs.weighting = "estimate" if by_estimate else "equal"

    def weight(t: Task) -> float:
        return float(t.estimate_h) if by_estimate else 1.0

    vs.progress = _weighted([(t.progress, weight(t)) for t in counted])
    # Compare like with like: progress and expected over the tasks that have both.
    planned = [t for t in counted if t.expected is not None]
    vs.expected = _weighted([(t.expected, weight(t)) for t in planned])
    plan_progress = _weighted([(t.progress, weight(t)) for t in planned])
    vs.light = light_for(
        vs.progress if vs.progress is not None and vs.progress >= 100.0 - _EPS else plan_progress,
        vs.expected, vs.due, on, settings.orange_pp, settings.red_pp,
    )
    return vs


def version_order(names: Iterable[str], preferred: Sequence[str]) -> List[str]:
    """Order of the version slides: the Settings order first, the rest naturally sorted."""
    names = list(dict.fromkeys(names))
    rank = {v.lower(): i for i, v in enumerate(preferred)}
    first = sorted((n for n in names if n.lower() in rank), key=lambda n: rank[n.lower()])
    rest = sorted((n for n in names if n.lower() not in rank), key=parse.natural_key)
    return first + rest


def compute(tasks: List[Task], settings: Settings) -> Tuple[List[VersionStatus], List[Issue]]:
    """Compute every task and version for ``settings.report_date``.

    Task order inside a version is the input order. Returns the versions in slide
    order and warnings for tasks that have no progress value.
    """
    on = settings.report_date or dt.date.today()
    issues: List[Issue] = []
    for task in tasks:
        compute_task(task, settings, on)
        if task.progress is None:
            issues.append(
                Issue("warning", task.sheet, task.row,
                      f"{task.id}: no progress (no Progress %, no done state, no Spent/Estimate); "
                      f"not counted in version {task.version}")
            )
    groups = {}
    for task in tasks:
        groups.setdefault(task.version, []).append(task)
    versions = [compute_version(v, groups[v], settings, on) for v in version_order(groups, settings.version_order)]
    return versions, issues


__all__ = [
    "GREEN", "ORANGE", "RED", "GREY",
    "task_progress", "expected_progress", "light_for", "compute_task", "compute_version", "version_order", "compute",
]
