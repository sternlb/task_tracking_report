"""Excel input -> list of tasks and issues — design 5.2, R-IN-*, R-OP-3, R-OP-4.

Structural problems (a required sheet or column is missing) raise
:class:`InputError`. Problems in single rows become :class:`Issue` entries:
``error`` skips the row, ``warning`` keeps it.

The input files are only read, never written (R-OP-4).
"""

from __future__ import annotations

import datetime as dt
import os
from dataclasses import dataclass
from typing import Dict, List, Optional, Tuple

from . import defaults, labels, parse
from .model import Issue, LoadResult, Settings, Task
from .xlsx_read import Workbook

HEADER_SEARCH_ROWS = 20  # a Polarion export may have title rows above its header


class InputError(Exception):
    """The input cannot be used at all (missing sheet, missing column, ...)."""


@dataclass
class Raw:
    value: object
    sheet: str
    row: int


Record = Dict[str, Raw]


def _norm(header) -> str:
    return " ".join(parse.text(header).split()).lower()


def _row_is_empty(row) -> bool:
    return all(parse.is_empty(v) for v in row)


def _cell(row: list, col: Optional[int]):
    if col is None or col >= len(row):
        return None
    return row[col]


# ---- settings ------------------------------------------------------------------------------


def _read_settings(wb: Workbook, issues: List[Issue]) -> Settings:
    settings = Settings()
    sheet = wb.find_sheet(labels.SHEET_SETTINGS)
    if sheet is None:
        issues.append(Issue("warning", labels.SHEET_SETTINGS, 0, "sheet missing, using default settings"))
        return settings
    by_label = {s.label.lower(): s for s in defaults.SETTINGS}
    for r, row in enumerate(wb.rows(sheet)[1:], start=2):
        key = parse.text(_cell(row, 0))
        if not key:
            continue
        sdef = by_label.get(key.lower())
        if sdef is None:
            issues.append(Issue("warning", sheet, r, f'unknown setting "{key}" ignored'))
            continue
        raw = _cell(row, 1)
        try:
            if sdef.kind == "text":
                value = parse.text(raw)
            elif sdef.kind == "date":
                value = parse.date(raw)
            elif sdef.kind == "list":
                value = parse.word_list(raw)
                if sdef.attr == "done_states":
                    value = [v.lower() for v in value]
            else:
                value = parse.number(raw)
                if value is None:
                    continue  # empty -> keep default
                if sdef.kind == "int":
                    if not value.is_integer() or value < 1:
                        raise ValueError(f"{parse.text(raw)} is not a whole number >= 1")
                    value = int(value)
                elif value < 0:
                    raise ValueError(f"{value:g} must not be negative")
        except ValueError as exc:
            issues.append(Issue("warning", sheet, r, f"{sdef.label}: {exc}; default used"))
            continue
        if sdef.kind == "text" and value == "" and sdef.attr in ("title_layout", "content_layout"):
            continue  # an empty layout name keeps the default
        setattr(settings, sdef.attr, value)
    if settings.hours_per_day <= 0:
        issues.append(Issue("warning", sheet, 0, "Hours per day must be > 0; 8 used"))
        settings.hours_per_day = 8.0
    if settings.orange_pp > settings.red_pp:
        issues.append(
            Issue("warning", sheet, 0, "Orange threshold is above the red threshold; defaults 10 / 25 used")
        )
        settings.orange_pp, settings.red_pp = Settings().orange_pp, Settings().red_pp
    return settings


# ---- mapping and import ----------------------------------------------------------------------


def _read_mapping(wb: Workbook) -> Dict[str, str]:
    """Field key -> normalised Polarion column header. Unmapped fields are left out."""
    sheet = wb.find_sheet(labels.SHEET_MAPPING)
    if sheet is None:
        raise InputError(f'sheet "{labels.SHEET_MAPPING}" is missing; it is needed to read the Polarion import')
    by_header = {f.header.lower(): f.key for f in labels.FIELDS}
    mapping: Dict[str, str] = {}
    for row in wb.rows(sheet)[1:]:
        key = by_header.get(_norm(_cell(row, 0)))
        column = _norm(_cell(row, 1))
        if key and column:
            mapping[key] = column
    if "id" not in mapping:
        raise InputError(f'{labels.SHEET_MAPPING}: no Polarion column is mapped to "ID"')
    return mapping


def _find_import_sheet(wb: Workbook, id_header: str) -> Optional[str]:
    for name in wb.sheet_names:
        rows = wb.rows(name)[:HEADER_SEARCH_ROWS]
        if any(id_header in (_norm(v) for v in row) for row in rows):
            return name
    return None


def _read_import(wb: Workbook, sheet: str, source: str, mapping: Dict[str, str], issues: List[Issue]):
    rows = wb.rows(sheet)
    header_idx = next(
        (i for i, row in enumerate(rows[:HEADER_SEARCH_ROWS]) if mapping["id"] in (_norm(v) for v in row)),
        None,
    )
    if header_idx is None:
        raise InputError(
            f'{source}: column "{mapping["id"]}" (mapped to ID) not found in the first {HEADER_SEARCH_ROWS} rows'
        )
    header = [_norm(v) for v in rows[header_idx]]
    cols: Dict[str, int] = {}
    missing = []
    for key, column in mapping.items():
        if column in header:
            cols[key] = header.index(column)
        else:
            missing.append(f'"{column}" ({labels.FIELD_BY_KEY[key].header})')
    if missing:
        found = ", ".join(f'"{h}"' for h in header if h)
        raise InputError(
            f"{source}: mapped column(s) {', '.join(missing)} not found. "
            f"Fix the names in the Mapping sheet or clear them. Columns in the export: {found}"
        )
    records = []
    for i, row in enumerate(rows[header_idx + 1:], start=header_idx + 2):
        if _row_is_empty(row):
            continue
        records.append((i, {k: Raw(_cell(row, c), source, i) for k, c in cols.items()}))
    return records


# ---- tasks sheet -------------------------------------------------------------------------------


def _read_tasks_sheet(wb: Workbook):
    sheet = wb.find_sheet(labels.SHEET_TASKS)
    if sheet is None:
        raise InputError(f'sheet "{labels.SHEET_TASKS}" is missing - is this the input workbook? Create one with --init')
    rows = wb.rows(sheet)
    if not rows:
        return sheet, []
    by_header = {f.header.lower(): f.key for f in labels.FIELDS}
    cols: Dict[str, int] = {}
    for c, h in enumerate(rows[0]):
        key = by_header.get(_norm(h))
        if key and key not in cols:
            cols[key] = c
    if "id" not in cols:
        raise InputError(f'{sheet}: column "ID" not found in the header row (row 1)')
    records = []
    for i, row in enumerate(rows[1:], start=2):
        if _row_is_empty(row):
            continue
        records.append((i, {k: Raw(_cell(row, c), sheet, i) for k, c in cols.items()}))
    return sheet, records


# ---- merge and build -----------------------------------------------------------------------------


def _add_records(records, target: Dict[str, Record], sources: Dict[str, str], label: str, issues: List[Issue], overlay: bool):
    seen: Dict[str, int] = {}
    for row, rec in records:
        raw_id = rec.get("id")
        tid = parse.text(raw_id.value) if raw_id else ""
        sheet = raw_id.sheet if raw_id else label
        if not tid:
            issues.append(Issue("error", sheet, row, "ID missing, row skipped"))
            continue
        if tid.startswith("#"):
            continue  # commented-out row (R-IN-9)
        if tid in seen:
            issues.append(Issue("error", sheet, row, f"duplicate ID {tid} (first in row {seen[tid]}), row skipped"))
            continue
        seen[tid] = row
        if overlay and tid in target:
            for key, raw in rec.items():
                if not parse.is_empty(raw.value):
                    target[tid][key] = raw  # R-IN-7: a filled Tasks cell wins
            sources[tid] = "Import+Tasks"
        else:
            target[tid] = dict(rec)
            sources[tid] = label


def _build_task(tid: str, rec: Record, source: str, settings: Settings, issues: List[Issue]) -> Optional[Task]:
    first = rec["id"]
    errors: List[Issue] = []

    def get(key: str, parser, *args):
        raw = rec.get(key)
        if raw is None:
            return None
        try:
            return parser(raw.value, *args)
        except ValueError as exc:
            errors.append(Issue("error", raw.sheet, raw.row, f"{tid}: {labels.FIELD_BY_KEY[key].header}: {exc}"))
            return None

    task = Task(
        id=tid,
        title=get("title", parse.text) or "",
        version=get("version", parse.text) or "",
        owner=get("owner", parse.text) or "",
        start=get("start", parse.date),
        due=get("due", parse.date),
        estimate_h=get("estimate_h", parse.hours, settings.hours_per_day),
        spent_h=get("spent_h", parse.hours, settings.hours_per_day),
        progress_manual=get("progress", parse.percent),
        state=get("state", parse.text) or "",
        light_override=get("light_override", parse.light),
        comment=get("comment", parse.text) or "",
        source=source,
        sheet=first.sheet,
        row=first.row,
    )
    for key in ("title", "version"):
        if not getattr(task, key):
            raw = rec.get(key, first)
            errors.append(Issue("error", raw.sheet, raw.row, f"{tid}: {labels.FIELD_BY_KEY[key].header} missing"))
    if errors:
        issues.extend(errors)
        issues.append(Issue("error", first.sheet, first.row, f"{tid}: row skipped"))
        return None
    if task.start and task.due and task.start > task.due:
        issues.append(Issue("warning", first.sheet, first.row, f"{tid}: Start is after Due"))
    return task


def load(path: str, polarion: Optional[str] = None, today: Optional[dt.date] = None) -> LoadResult:
    """Read the input workbook (and an optional separate Polarion export)."""
    wb = Workbook.open(path)
    issues: List[Issue] = []
    settings = _read_settings(wb, issues)
    if settings.report_date is None:
        settings.report_date = today or dt.date.today()

    tasks_sheet, task_records = _read_tasks_sheet(wb)

    import_records = []
    import_source = ""
    if polarion:
        mapping = _read_mapping(wb)
        pwb = Workbook.open(polarion)
        sheet = _find_import_sheet(pwb, mapping["id"])
        name = os.path.basename(polarion)
        if sheet is None:
            raise InputError(f'{name}: no sheet with a column "{mapping["id"]}" (mapped to ID)')
        import_records = _read_import(pwb, sheet, name, mapping, issues)
        import_source = polarion
    else:
        isheet = wb.find_sheet(labels.SHEET_IMPORT)
        if isheet is not None and any(not _row_is_empty(r) for r in wb.rows(isheet)):
            mapping = _read_mapping(wb)
            import_records = _read_import(wb, isheet, isheet, mapping, issues)
            import_source = "Import sheet"

    merged: Dict[str, Record] = {}
    sources: Dict[str, str] = {}
    _add_records(import_records, merged, sources, "Import", issues, overlay=False)
    _add_records(task_records, merged, sources, tasks_sheet, issues, overlay=True)

    tasks = []
    for tid, rec in merged.items():
        task = _build_task(tid, rec, "Tasks" if sources[tid] == tasks_sheet else sources[tid], settings, issues)
        if task:
            tasks.append(task)

    if settings.version_order:
        known = {v.lower() for v in settings.version_order}
        for version in sorted({t.version for t in tasks}, key=parse.natural_key):
            if version.lower() not in known:
                issues.append(
                    Issue("warning", labels.SHEET_SETTINGS, 0, f'version "{version}" is not in "Version order"; shown last')
                )
    if not tasks and not [i for i in issues if i.level == "error"]:
        issues.append(Issue("warning", tasks_sheet, 0, "no tasks found"))
    issues.sort(key=lambda i: (i.sheet, i.row))  # stable: keeps the order within a row
    return LoadResult(settings, tasks, issues, import_source)


__all__ = ["InputError", "load"]
