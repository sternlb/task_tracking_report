"""Weekly snapshots — R-HI-1 … R-HI-4, design 7.

One human-readable JSON file per ISO calendar week in ``history/`` next to the
input workbook, e.g. ``history/2026-W46.json``. A later run in the same week
replaces that week's file. The snapshots provide the previous-week values
(shown in brackets) and the data of the progress chart.
"""

from __future__ import annotations

import datetime as dt
import json
import os
import re
import tempfile
from dataclasses import dataclass, field
from typing import Dict, List, Optional, Sequence, Tuple

from .model import Issue, VersionStatus

SCHEMA = 1
FOLDER = "history"
_FILE = re.compile(r"^(\d{4})-W(\d{2})\.json$")


@dataclass
class TaskSnap:
    version: str
    title: str
    progress: Optional[float]
    light: str


@dataclass
class VersionSnap:
    progress: Optional[float]
    expected: Optional[float]
    light: str


@dataclass
class Snapshot:
    week: str  # "2026-W46"
    report_date: dt.date
    versions: Dict[str, VersionSnap] = field(default_factory=dict)
    tasks: Dict[str, TaskSnap] = field(default_factory=dict)


# ---- weeks ------------------------------------------------------------------------------------------


def week_key(day: dt.date) -> str:
    """ISO calendar week, e.g. ``2026-W46``. 31 Dec 2026 is ``2026-W53``, 3 Jan 2021 is ``2020-W53``."""
    year, week, _ = day.isocalendar()
    return f"{year:04d}-W{week:02d}"


def week_labels(keys: Sequence[str]) -> List[str]:
    """Axis labels ``CW 46``; with the year (``CW 52/25``) when the weeks span more than one year."""
    years = {k[:4] for k in keys}
    if len(years) <= 1:
        return [f"CW {k[6:]}" for k in keys]
    return [f"CW {k[6:]}/{k[2:4]}" for k in keys]


def history_dir(input_path: str) -> str:
    return os.path.join(os.path.dirname(os.path.abspath(input_path)), FOLDER)


# ---- building and (de)serialising -------------------------------------------------------------------------


def _round(value: Optional[float]) -> Optional[float]:
    return None if value is None else round(value, 2)


def snapshot_from(versions: Sequence[VersionStatus], report_date: dt.date) -> Snapshot:
    snap = Snapshot(week_key(report_date), report_date)
    for vs in versions:
        snap.versions[vs.name] = VersionSnap(_round(vs.progress), _round(vs.expected), vs.light)
        for t in vs.tasks:
            snap.tasks[t.id] = TaskSnap(t.version, t.title, _round(t.progress), t.light)
    return snap


def to_json(snap: Snapshot, generated_at: Optional[dt.datetime] = None) -> str:
    stamp = (generated_at or dt.datetime.now()).replace(microsecond=0).isoformat()
    data = {
        "schema": SCHEMA,
        "week": snap.week,
        "report_date": snap.report_date.isoformat(),
        "generated_at": stamp,
        "versions": {
            name: {"progress": v.progress, "expected": v.expected, "light": v.light} for name, v in snap.versions.items()
        },
        "tasks": {
            tid: {"version": t.version, "title": t.title, "progress": t.progress, "light": t.light}
            for tid, t in snap.tasks.items()
        },
    }
    return json.dumps(data, indent=2, ensure_ascii=False) + "\n"


def from_json(text: str) -> Snapshot:
    data = json.loads(text)
    if not isinstance(data, dict):
        raise ValueError("not a snapshot object")
    if data.get("schema") != SCHEMA:
        raise ValueError(f"unknown schema {data.get('schema')!r}")
    snap = Snapshot(str(data["week"]), dt.date.fromisoformat(data["report_date"]))
    for name, v in (data.get("versions") or {}).items():
        snap.versions[name] = VersionSnap(v.get("progress"), v.get("expected"), v.get("light", "grey"))
    for tid, t in (data.get("tasks") or {}).items():
        snap.tasks[tid] = TaskSnap(t.get("version", ""), t.get("title", ""), t.get("progress"), t.get("light", "grey"))
    return snap


# ---- files -------------------------------------------------------------------------------------------------


def save(folder: str, snap: Snapshot, generated_at: Optional[dt.datetime] = None) -> str:
    """Write the snapshot for its week, replacing an existing one. Atomic: temp file + rename."""
    os.makedirs(folder, exist_ok=True)
    path = os.path.join(folder, f"{snap.week}.json")
    fd, tmp = tempfile.mkstemp(prefix=f".{snap.week}.", suffix=".tmp", dir=folder)
    try:
        with os.fdopen(fd, "w", encoding="utf-8", newline="\n") as fh:
            fh.write(to_json(snap, generated_at))
        os.replace(tmp, path)
    except BaseException:
        if os.path.exists(tmp):
            os.remove(tmp)
        raise
    return path


def load_all(folder: str, issues: Optional[List[Issue]] = None) -> Dict[str, Snapshot]:
    """All readable snapshots by week key. Unreadable files become warnings and are skipped."""
    out: Dict[str, Snapshot] = {}
    if not os.path.isdir(folder):
        return out
    for name in sorted(os.listdir(folder)):
        m = _FILE.match(name)
        if not m:
            continue
        try:
            with open(os.path.join(folder, name), encoding="utf-8") as fh:
                snap = from_json(fh.read())
            if snap.week != name[:-5]:
                raise ValueError(f"file contains week {snap.week}")
        except (OSError, ValueError, KeyError, TypeError, AttributeError) as exc:
            if issues is not None:
                issues.append(Issue("warning", FOLDER, 0, f"{name} skipped: {exc}"))
            continue
        out[snap.week] = snap
    return out


# ---- queries ---------------------------------------------------------------------------------------------------


def previous(snapshots: Dict[str, Snapshot], week: str) -> Optional[Snapshot]:
    """R-HI-3: the latest snapshot before ``week`` (skipped weeks are fine)."""
    earlier = [k for k in snapshots if k < week]
    return snapshots[max(earlier)] if earlier else None


def trend(
    snapshots: Dict[str, Snapshot], current: Snapshot, weeks: int, version_names: Sequence[str]
) -> Tuple[List[str], List[str], List[Tuple[str, List[Optional[float]]]]]:
    """Data for the progress chart (R-OUT-2).

    The last ``weeks`` weeks that have a snapshot, up to and including the current
    week (the in-memory ``current`` replaces a stored snapshot of the same week, so
    ``--no-save`` runs show the right values). One series per current version;
    weeks without a value for a version give ``None`` (a gap in the chart).

    Returns (week keys, axis labels, series).
    """
    merged = {k: s for k, s in snapshots.items() if k < current.week}
    merged[current.week] = current
    keys = sorted(merged)[-max(1, weeks):]
    series = []
    for name in version_names:
        values = []
        for k in keys:
            v = merged[k].versions.get(name)
            values.append(None if v is None else v.progress)
        series.append((name, values))
    return keys, week_labels(keys), series


__all__ = [
    "Snapshot", "TaskSnap", "VersionSnap", "FOLDER",
    "week_key", "week_labels", "history_dir", "snapshot_from", "to_json", "from_json",
    "save", "load_all", "previous", "trend",
]
