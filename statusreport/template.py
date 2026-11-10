"""The empty Excel input template created by ``--init`` (R-IN-1 … R-IN-5, T-2.3)."""

from __future__ import annotations

import datetime as dt
from typing import List

from . import defaults, labels
from .xlsx_write import DropDown, Sheet, Styled, col_letter, workbook_bytes

MAX_ROWS = 2000  # rows covered by drop-downs and column formats

_WIDTHS = {
    "id": 14, "title": 40, "version": 12, "owner": 14, "start": 12, "due": 12,
    "estimate_h": 13, "spent_h": 11, "progress": 13, "state": 14, "light_override": 15, "comment": 40,
}
_EXAMPLE = {
    "id": "#EXAMPLE-1",
    "title": "Login dialog (example row - ignored because the ID starts with #)",
    "version": "1.2.0",
    "owner": "Anna",
    "start": dt.date(2026, 11, 2),
    "due": dt.date(2026, 11, 27),
    "estimate_h": "3d",
    "spent_h": 10,
    "progress": None,
    "state": "in progress",
    "light_override": None,
    "comment": "waiting for API",
}
_SETTING_VALUES = {
    "project": "My Project",
    "customer": "Customer",
}


def _readme() -> Sheet:
    rows: List[list] = []
    for i, line in enumerate(labels.README):
        if i == 0:
            rows.append([Styled(line, "title")])
        elif line and not line.startswith(("-", " ", "1", "2", "3")):
            rows.append([Styled(line, "bold")])
        else:
            rows.append([line])
    return Sheet(labels.SHEET_README, rows, widths={0: 120})


def _tasks() -> Sheet:
    keys = [f.key for f in labels.FIELDS]
    header = [Styled(f.header, "header") for f in labels.FIELDS]
    example = []
    for k in keys:
        v = _EXAMPLE[k]
        if isinstance(v, dt.date):
            example.append(Styled(v, "example_date"))
        else:
            example.append(Styled(v, "example") if v not in (None, "") else None)
    col_styles = {}
    for i, k in enumerate(keys):
        if k in ("id", "version"):
            col_styles[i] = "text"
        elif k in ("start", "due"):
            col_styles[i] = "date"
    light_col = col_letter(keys.index("light_override"))
    return Sheet(
        labels.SHEET_TASKS,
        [header, example],
        widths={i: _WIDTHS[k] for i, k in enumerate(keys)},
        col_styles=col_styles,
        freeze_header=True,
        dropdowns=[DropDown(f"{light_col}2:{light_col}{MAX_ROWS}", labels.LIGHT_CHOICES)],
    )


def _mapping() -> Sheet:
    rows: List[list] = [[Styled(h, "header") for h in labels.MAPPING_HEADERS]]
    for f in labels.FIELDS:
        note = defaults.MAPPING_NOTES.get(f.key, "")
        rows.append([f.header, defaults.POLARION_COLUMNS.get(f.key, ""), Styled(note, "note") if note else None])
    return Sheet(labels.SHEET_MAPPING, rows, widths={0: 16, 1: 24, 2: 80}, freeze_header=True)


def _settings() -> Sheet:
    rows: List[list] = [[Styled(h, "header") for h in labels.SETTINGS_HEADERS]]
    for s in defaults.SETTINGS:
        value = _SETTING_VALUES.get(s.attr, defaults.default_value(s.attr))
        if s.kind == "list":
            value = ", ".join(value)
        elif s.kind == "date":
            value = Styled(None, "date")
        elif s.kind in ("number", "int") and isinstance(value, float) and value.is_integer():
            value = int(value)
        rows.append([s.label, value, Styled(s.help, "note")])
    return Sheet(labels.SHEET_SETTINGS, rows, widths={0: 24, 1: 34, 2: 80}, freeze_header=True)


def template_bytes() -> bytes:
    sheets = [_readme(), _tasks(), Sheet(labels.SHEET_IMPORT, []), _mapping(), _settings()]
    return workbook_bytes(sheets, active=0)
