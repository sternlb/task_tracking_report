"""All user-visible texts (constraint C-4). English only for now.

Input side (Excel template): sheet names, column headers, setting names and
help texts. The slide texts follow in M5.
"""

from __future__ import annotations

from typing import List, NamedTuple

SHEET_README = "Read me"
SHEET_TASKS = "Tasks"
SHEET_IMPORT = "Import"
SHEET_MAPPING = "Mapping"
SHEET_SETTINGS = "Settings"


class Field(NamedTuple):
    key: str  # attribute name in model.Task
    header: str  # column header in the Tasks sheet and field name in the Mapping sheet
    help: str


# Order = column order of the Tasks sheet (R-IN-2)
FIELDS: List[Field] = [
    Field("id", "ID", "Unique key, e.g. the Polarion work item ID. Start it with # to ignore the row."),
    Field("title", "Title", "Short task name shown on the slide."),
    Field("version", "Version", "Software version / release. Keep the column formatted as text (1.10 is not 1.1)."),
    Field("owner", "Owner", "Responsible person. Not shown on customer slides."),
    Field("start", "Start", "Planned start date. Empty = Due minus Estimate in working days."),
    Field("due", "Due", "Planned finish date. Needed for the traffic light. Not shown on customer slides."),
    Field("estimate_h", "Estimate (h)", "Estimated effort: hours, or a duration such as 2d 4h."),
    Field("spent_h", "Spent (h)", "Effort used so far: hours, or a duration such as 1d 2h."),
    Field("progress", "Progress (%)", "Completion 0-100. Wins over the calculation from hours."),
    Field("state", "State", "Work item state. Done states (see Settings) count as 100 %."),
    Field("light_override", "Light override", "Green, Orange or Red: sets the traffic light by hand."),
    Field("comment", "Comment", "Short remark shown on the slide."),
]
FIELD_BY_KEY = {f.key: f for f in FIELDS}
REQUIRED = ("id", "title", "version")

MAPPING_HEADERS = ("Field", "Polarion column", "Notes")
SETTINGS_HEADERS = ("Setting", "Value", "Description")
LIGHT_CHOICES = ("Green", "Orange", "Red")

# ---- slides (M5) ---------------------------------------------------------------------------------

MONTHS = ("Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec")
DECK_TITLE_FALLBACK = "Software status"
SUBTITLE = "Status report"  # joined with customer, calendar week and date by " · "
CALENDAR_WEEK = "CW {week:02d} / {year}"
TREND_TITLE = "Progress over time"
TREND_FIRST_WEEK = "The trend builds up from next week on."
BURNDOWN_TITLE = "Burndown: remaining effort"
BURNDOWN_ACTUAL = "Open effort (%)"
BURNDOWN_IDEAL = "Ideal"
BURNDOWN_NOTE = "Open share of the total effort of all tasks: {pct:.0f} %"
BURNDOWN_NOTE_END = " · ideal line down to 0 % at the latest due date ({end})"
VERSION_TITLE = "Version {name}"
VERSION_TITLE_CONTINUED = "Version {name} ({page}/{pages})"
TABLE_HEADERS = ("ID", "Task", "Progress (prev. week)", "Status", "Comment")
TOTAL_ROW = "Version total"
PREV_NEW = "(new)"
PREV_NONE = "(–)"
NO_VALUE = "–"
LIGHT_DOT = "●"
LEGEND = (
    ("green", "on track"),
    ("orange", "behind by more than {orange:g} pp"),
    ("red", "behind by more than {red:g} pp or overdue"),
    ("grey", "not assessable"),
)
LEGEND_PREVIOUS = "(…) = previous week"
OUTPUT_NAME = "{project} Status {week}.pptx"
OUTPUT_NAME_NO_PROJECT = "Status {week}.pptx"


def format_date(day) -> str:
    """10 Nov 2026 — fixed English month names, independent of the PC's locale."""
    return f"{day.day} {MONTHS[day.month - 1]} {day.year}"


README = [
    "SW Status Report - input workbook",
    "",
    "Weekly routine",
    "1. Update the task list: type it into the 'Tasks' sheet, or paste a Polarion export into 'Import' (see below).",
    "2. Save this file and drag it onto make_report.cmd (or run: python statusreport.pyz <this file>).",
    "3. The deck appears next to this file. Each run also stores this week's values in the 'history' folder,",
    "   which gives the previous-week values in brackets and the progress chart. Keep that folder.",
    "",
    "Tasks sheet",
    "- One row per work package. ID, Title and Version are required.",
    "- Due is needed for the traffic light. An empty Start is derived from Due minus Estimate (working days).",
    "- Owner and Due are never shown to the customer.",
    "- Progress: a value in 'Progress (%)' wins; otherwise a done state counts as 100 %;",
    "  otherwise Spent / Estimate is used (at most 99 % until the task is done).",
    "- Rows whose ID starts with # are ignored (like the example row).",
    "",
    "Polarion export",
    "- Export the work items from Polarion to Excel and paste the whole table (with its header row) into 'Import',",
    "  or pass the export file on the command line: --polarion export.xlsx",
    "- 'Mapping' says which Polarion column feeds which field. Adjust the column names to your export.",
    "- A filled cell in 'Tasks' wins over the import for the same ID, so comments and overrides go there.",
    "",
    "Settings",
    "- Project and customer names, report date, traffic-light thresholds and more. See the descriptions there.",
]
