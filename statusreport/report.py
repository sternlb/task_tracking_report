"""Presentation-neutral report model — design 8.2, T-5.1.

Turns the computed versions, the previous snapshot and the trend data into
the content of every slide: texts, values and which previous-week value goes
into the brackets. Rendering to PowerPoint is done by :mod:`statusreport.render`.
"""

from __future__ import annotations

import datetime as dt
from dataclasses import dataclass, field
from typing import List, Optional, Sequence, Tuple

from . import labels
from .history import Snapshot
from .model import Settings, VersionStatus

MAX_TITLE = 62  # characters; longer task titles are shortened with "…"
MAX_COMMENT = 45


@dataclass
class Row:
    id: str
    title: str
    progress: Optional[float]
    prev_text: str  # "(50 %)", "(new)" or "(–)"
    light: str
    prev_light: Optional[str]  # only set when it differs from ``light`` (R-OUT-6)
    comment: str = ""
    is_total: bool = False


@dataclass
class TablePage:
    version: str
    title: str
    total: Row
    rows: List[Row]


@dataclass
class Report:
    title: str
    subtitle: str
    week: str  # "2026-W46"
    report_date: dt.date
    trend_labels: List[str]
    trend_series: List[Tuple[str, List[Optional[float]]]]
    trend_note: str  # empty unless the trend has only one week
    pages: List[TablePage] = field(default_factory=list)
    orange_pp: float = 10.0
    red_pp: float = 25.0


def shorten(text: str, limit: int) -> str:
    text = " ".join(text.split())
    return text if len(text) <= limit else text[: limit - 1].rstrip() + "…"


def prev_text(prev: Optional[Snapshot], found: bool, value: Optional[float]) -> str:
    """R-OUT-4: ``(50 %)``; ``(new)`` if the item was not in last week's report; ``(–)`` without history."""
    if prev is None:
        return labels.PREV_NONE
    if not found:
        return labels.PREV_NEW
    return labels.PREV_NONE if value is None else f"({value:.0f} %)"


def _changed(light: str, previous: Optional[str]) -> Optional[str]:
    return previous if previous and previous != light else None


def _version_rows(vs: VersionStatus, prev: Optional[Snapshot]) -> Tuple[Row, List[Row]]:
    pv = prev.versions.get(vs.name) if prev else None
    total = Row(
        id=labels.TOTAL_ROW, title="", progress=vs.progress,
        prev_text=prev_text(prev, pv is not None, pv.progress if pv else None),
        light=vs.light, prev_light=_changed(vs.light, pv.light if pv else None), is_total=True,
    )
    rows = []
    for t in vs.tasks:
        pt = prev.tasks.get(t.id) if prev else None
        rows.append(
            Row(
                id=t.id, title=shorten(t.title, MAX_TITLE), progress=t.progress,
                prev_text=prev_text(prev, pt is not None, pt.progress if pt else None),
                light=t.light, prev_light=_changed(t.light, pt.light if pt else None),
                comment=shorten(t.comment, MAX_COMMENT),
            )
        )
    return total, rows


def build_report(
    settings: Settings,
    versions: Sequence[VersionStatus],
    prev: Optional[Snapshot],
    week: str,
    trend_labels: Sequence[str],
    trend_series: Sequence[Tuple[str, List[Optional[float]]]],
) -> Report:
    day = settings.report_date or dt.date.today()
    iso = day.isocalendar()
    sub = [labels.SUBTITLE]
    if settings.customer:
        sub.append(settings.customer)
    sub.append(labels.CALENDAR_WEEK.format(week=iso[1], year=iso[0]))
    sub.append(labels.format_date(day))
    report = Report(
        title=settings.project or labels.DECK_TITLE_FALLBACK,
        subtitle=" · ".join(sub),
        week=week,
        report_date=day,
        trend_labels=list(trend_labels),
        trend_series=[(n, list(v)) for n, v in trend_series],
        trend_note=labels.TREND_FIRST_WEEK if len(trend_labels) <= 1 else "",
        orange_pp=settings.orange_pp,
        red_pp=settings.red_pp,
    )
    per_page = max(1, settings.rows_per_slide)
    for vs in versions:
        total, rows = _version_rows(vs, prev)
        chunks = [rows[i:i + per_page] for i in range(0, len(rows), per_page)] or [[]]
        for n, chunk in enumerate(chunks, 1):
            if len(chunks) == 1:
                title = labels.VERSION_TITLE.format(name=vs.name)
            else:
                title = labels.VERSION_TITLE_CONTINUED.format(name=vs.name, page=n, pages=len(chunks))
            report.pages.append(TablePage(vs.name, title, total, chunk))  # total row repeated per page
    return report


__all__ = ["Row", "TablePage", "Report", "build_report", "prev_text", "shorten"]
