"""Default values for settings and the Polarion mapping (R-IN-3, R-IN-5)."""

from __future__ import annotations

from typing import Dict, List, NamedTuple

from .model import Settings


class SettingDef(NamedTuple):
    attr: str  # attribute in model.Settings
    label: str  # text in column A of the Settings sheet
    kind: str  # text | date | number | int | list
    help: str


SETTINGS: List[SettingDef] = [
    SettingDef("project", "Project", "text", "Project name. Used on the title slide and in the file name."),
    SettingDef("customer", "Customer", "text", "Customer name for the title slide."),
    SettingDef("report_date", "Report date", "date", "Empty = today. Sets the calendar week and the expected progress."),
    SettingDef("orange_pp", "Orange threshold (pp)", "number", "More than this many percentage points behind schedule -> orange."),
    SettingDef("red_pp", "Red threshold (pp)", "number", "More than this many percentage points behind schedule -> red."),
    SettingDef("hours_per_day", "Hours per day", "number", "For durations such as 2d. One week (1w) = 5 days."),
    SettingDef("done_states", "Done states", "list", "Comma-separated states that count as 100 %."),
    SettingDef("rows_per_slide", "Rows per slide", "int", "Longer version tables continue on the next slide."),
    SettingDef("trend_weeks", "Trend weeks", "int", "Number of calendar weeks in the progress chart."),
    SettingDef("version_order", "Version order", "list", "Comma-separated order of the version slides. Empty = 1.2 before 1.10."),
    SettingDef("template", "Template file", "text", "PowerPoint template (.pptx), relative to this file. Empty = neutral built-in design."),
    SettingDef("title_layout", "Title layout", "text", "Name of the template layout for the title slide."),
    SettingDef("content_layout", "Content layout", "text", "Name of the template layout for chart and table slides."),
    SettingDef("footer_text", "Footer text", "text", "Footer on the chart and table slides, e.g. 'Confidential' (only if the template has a footer field)."),
]


def default_value(attr: str):
    return getattr(Settings(), attr)


# Typical Polarion column names; adjust in the Mapping sheet. Empty = not imported.
POLARION_COLUMNS: Dict[str, str] = {
    "id": "ID",
    "title": "Title",
    "version": "Planned In",
    "owner": "Assignee",
    "start": "",
    "due": "Due Date",
    "estimate_h": "Initial Estimate",
    "spent_h": "Time Spent",
    "progress": "",
    "state": "Status",
    "light_override": "",
    "comment": "",
}

MAPPING_NOTES: Dict[str, str] = {
    "id": "Required. Used to match Import rows with Tasks rows.",
    "version": "Often 'Planned In', 'Target Release' or a custom field.",
    "start": "Polarion has no standard start field. Empty = Start is derived from Due minus Estimate; "
    "or enter a custom field here, or fill Start in Tasks.",
    "spent_h": "An empty cell in the export counts as 0 h (nothing booked yet).",
    "progress": "Usually empty: progress comes from hours or the state.",
    "light_override": "Usually maintained in Tasks, not in Polarion.",
    "comment": "Usually maintained in Tasks, not in Polarion.",
}
