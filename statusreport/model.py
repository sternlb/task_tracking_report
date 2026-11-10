"""Data model shared by loader, calculation, history and report (design 4)."""

from __future__ import annotations

import datetime as dt
from dataclasses import dataclass, field
from typing import List, Optional


@dataclass
class Settings:
    project: str = ""
    customer: str = ""
    report_date: Optional[dt.date] = None  # None = today, resolved by the loader
    orange_pp: float = 10.0
    red_pp: float = 25.0
    hours_per_day: float = 8.0
    done_states: List[str] = field(default_factory=lambda: ["done", "closed", "resolved", "verified"])
    rows_per_slide: int = 12
    trend_weeks: int = 12
    version_order: List[str] = field(default_factory=list)
    template: str = ""
    title_layout: str = "Title Slide"
    content_layout: str = "Title Only"


@dataclass
class Task:
    id: str
    title: str
    version: str
    owner: str = ""
    start: Optional[dt.date] = None
    due: Optional[dt.date] = None
    estimate_h: Optional[float] = None
    spent_h: Optional[float] = None
    progress_manual: Optional[float] = None
    state: str = ""
    light_override: Optional[str] = None
    comment: str = ""
    source: str = "Tasks"  # "Tasks", "Import" or "Import+Tasks"
    sheet: str = ""
    row: int = 0  # Excel row number, for messages
    # computed in M3 (calc)
    progress: Optional[float] = None
    expected: Optional[float] = None
    light: str = "grey"


@dataclass
class Issue:
    level: str  # "error" (row skipped) or "warning"
    sheet: str
    row: int  # Excel row number, 0 = not row-specific
    message: str

    def __str__(self) -> str:
        where = f"{self.sheet} row {self.row}" if self.row else self.sheet
        return f"{self.level:<7} {where}: {self.message}"


@dataclass
class LoadResult:
    settings: Settings
    tasks: List[Task]
    issues: List[Issue]
    import_source: str = ""  # "", "Import sheet" or the path of a --polarion file

    @property
    def errors(self) -> List[Issue]:
        return [i for i in self.issues if i.level == "error"]
