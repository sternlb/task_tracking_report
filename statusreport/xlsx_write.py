"""Minimal .xlsx writer (standard library only) — T-2.2.

Used for the input template (``--init``) and for the workbook embedded in a chart.
Supports a fixed set of cell styles, column widths, a frozen header row and
drop-down lists (data validation). Text is written as inline strings.
"""

from __future__ import annotations

import datetime as dt
from dataclasses import dataclass, field
from typing import Dict, List, Optional, Sequence, Tuple, Union

from .pptx import opc
from .xmlutil import NS_R, NS_S, doc, esc

Value = Union[str, int, float, bool, dt.date, None]

_SML = "application/vnd.openxmlformats-officedocument.spreadsheetml."
CT_WORKBOOK = _SML + "sheet.main+xml"
CT_WORKSHEET = _SML + "worksheet+xml"
CT_STYLES = _SML + "styles+xml"
RT_WORKSHEET = opc.RT + "/worksheet"
RT_STYLES = opc.RT + "/styles"

HEADER_FILL = "FF1F4E79"

# style name -> index into cellXfs below
STYLES = {
    "": 0,
    "header": 1,
    "date": 2,
    "note": 3,
    "bold": 4,
    "example": 5,
    "example_date": 6,
    "text": 7,
    "title": 8,
}
_DATE_FMT = 164  # custom yyyy-mm-dd
_STYLES_XML = (
    f'<styleSheet xmlns="{NS_S}">'
    f'<numFmts count="1"><numFmt numFmtId="{_DATE_FMT}" formatCode="yyyy-mm-dd"/></numFmts>'
    '<fonts count="5">'
    '<font><sz val="11"/><name val="Calibri"/><family val="2"/></font>'
    '<font><b/><sz val="11"/><color rgb="FFFFFFFF"/><name val="Calibri"/><family val="2"/></font>'
    '<font><i/><sz val="10"/><color rgb="FF7F7F7F"/><name val="Calibri"/><family val="2"/></font>'
    '<font><b/><sz val="11"/><name val="Calibri"/><family val="2"/></font>'
    '<font><b/><sz val="14"/><color rgb="FF1F4E79"/><name val="Calibri"/><family val="2"/></font>'
    "</fonts>"
    '<fills count="3"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill>'
    f'<fill><patternFill patternType="solid"><fgColor rgb="{HEADER_FILL}"/><bgColor indexed="64"/></patternFill></fill></fills>'
    '<borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>'
    '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>'
    '<cellXfs count="9">'
    '<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>'
    '<xf numFmtId="0" fontId="1" fillId="2" borderId="0" xfId="0" applyFont="1" applyFill="1" applyAlignment="1">'
    '<alignment vertical="center" wrapText="1"/></xf>'
    f'<xf numFmtId="{_DATE_FMT}" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>'
    '<xf numFmtId="0" fontId="2" fillId="0" borderId="0" xfId="0" applyFont="1" applyAlignment="1">'
    '<alignment vertical="top" wrapText="1"/></xf>'
    '<xf numFmtId="0" fontId="3" fillId="0" borderId="0" xfId="0" applyFont="1"/>'
    '<xf numFmtId="0" fontId="2" fillId="0" borderId="0" xfId="0" applyFont="1"/>'
    f'<xf numFmtId="{_DATE_FMT}" fontId="2" fillId="0" borderId="0" xfId="0" applyNumberFormat="1" applyFont="1"/>'
    '<xf numFmtId="49" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>'
    '<xf numFmtId="0" fontId="4" fillId="0" borderId="0" xfId="0" applyFont="1"/>'
    "</cellXfs>"
    '<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>'
    "</styleSheet>"
)


@dataclass
class Styled:
    """A cell value with a named style from :data:`STYLES`."""

    value: Value
    style: str


@dataclass
class DropDown:
    """List validation for a cell range, e.g. ``DropDown("K2:K1000", ["Green", "Orange", "Red"])``."""

    ref: str
    choices: Sequence[str]


@dataclass
class Sheet:
    name: str
    rows: Sequence[Sequence[Union[Value, Styled]]]
    widths: Dict[int, float] = field(default_factory=dict)  # 0-based column -> width in characters
    col_styles: Dict[int, str] = field(default_factory=dict)  # default style of empty cells per column
    freeze_header: bool = False
    dropdowns: List[DropDown] = field(default_factory=list)


def col_letter(index: int) -> str:
    """0 -> A, 25 -> Z, 26 -> AA."""
    out = ""
    index += 1
    while index:
        index, rem = divmod(index - 1, 26)
        out = chr(65 + rem) + out
    return out


def date_to_serial(value: dt.date) -> float:
    if isinstance(value, dt.datetime):
        delta = value - dt.datetime(1899, 12, 30)
        return delta.days + delta.seconds / 86400
    return float((value - dt.date(1899, 12, 30)).days)


def _cell(ref: str, value, style: str = "") -> str:
    if isinstance(value, Styled):
        value, style = value.value, value.style
    if isinstance(value, dt.date) and style in ("", "text"):
        style = "date"
    s = STYLES[style]
    s_attr = f' s="{s}"' if s else ""
    if value is None or value == "":
        return f'<c r="{ref}"{s_attr}/>' if s else ""
    if isinstance(value, bool):
        return f'<c r="{ref}"{s_attr} t="b"><v>{int(value)}</v></c>'
    if isinstance(value, dt.date):
        return f'<c r="{ref}"{s_attr}><v>{date_to_serial(value):g}</v></c>'
    if isinstance(value, (int, float)):
        return f'<c r="{ref}"{s_attr}><v>{value!r}</v></c>'
    text = esc(value)
    space = ' xml:space="preserve"' if text != text.strip() or "\n" in text else ""
    return f'<c r="{ref}"{s_attr} t="inlineStr"><is><t{space}>{text}</t></is></c>'


def _sheet_xml(sheet: Sheet) -> bytes:
    parts = []
    if sheet.freeze_header:
        parts.append(
            '<sheetViews><sheetView workbookViewId="0">'
            '<pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/>'
            '<selection pane="bottomLeft" activeCell="A2" sqref="A2"/></sheetView></sheetViews>'
        )
    if sheet.widths or sheet.col_styles:
        cols = sorted(set(sheet.widths) | set(sheet.col_styles))
        entries = []
        for c in cols:
            width = sheet.widths.get(c, 10.0)
            style = STYLES[sheet.col_styles.get(c, "")]
            s_attr = f' style="{style}"' if style else ""
            entries.append(f'<col min="{c + 1}" max="{c + 1}" width="{width:g}" customWidth="1"{s_attr}/>')
        parts.append(f"<cols>{''.join(entries)}</cols>")
    rows = []
    for r, row in enumerate(sheet.rows, 1):
        cells = "".join(_cell(f"{col_letter(c)}{r}", v) for c, v in enumerate(row))
        if cells:
            rows.append(f'<row r="{r}">{cells}</row>')
    parts.append(f"<sheetData>{''.join(rows)}</sheetData>")
    if sheet.dropdowns:
        dvs = "".join(
            '<dataValidation type="list" allowBlank="1" showInputMessage="1" showErrorMessage="1" '
            f'sqref="{d.ref}"><formula1>"{esc(",".join(d.choices))}"</formula1></dataValidation>'
            for d in sheet.dropdowns
        )
        parts.append(f'<dataValidations count="{len(sheet.dropdowns)}">{dvs}</dataValidations>')
    parts.append('<pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/>')
    return doc(f'<worksheet xmlns="{NS_S}" xmlns:r="{NS_R}">{"".join(parts)}</worksheet>')


def workbook_bytes(sheets: List[Union[Sheet, Tuple[str, Sequence[Sequence[Value]]]]], active: int = 0) -> bytes:
    pkg = opc.Package()
    wb = "xl/workbook.xml"
    pkg.add_rel("", opc.RT_OFFICE_DOCUMENT, wb)
    entries = []
    for i, sheet in enumerate(sheets, 1):
        if not isinstance(sheet, Sheet):
            sheet = Sheet(sheet[0], sheet[1])
        part = f"xl/worksheets/sheet{i}.xml"
        pkg.add_part(part, _sheet_xml(sheet), CT_WORKSHEET)
        rid = pkg.add_rel(wb, RT_WORKSHEET, part)
        entries.append(f'<sheet name="{esc(sheet.name)}" sheetId="{i}" r:id="{rid}"/>')
    pkg.add_part("xl/styles.xml", doc(_STYLES_XML), CT_STYLES)
    pkg.add_rel(wb, RT_STYLES, "xl/styles.xml")
    pkg.add_part(
        wb,
        doc(
            f'<workbook xmlns="{NS_S}" xmlns:r="{NS_R}">'
            f'<bookViews><workbookView activeTab="{active}"/></bookViews>'
            f'<sheets>{"".join(entries)}</sheets></workbook>'
        ),
        CT_WORKBOOK,
    )
    return pkg.to_bytes()


def save_workbook(path: str, sheets: List[Sheet], active: int = 0) -> None:
    data = workbook_bytes(sheets, active)
    with open(path, "wb") as fh:
        fh.write(data)


__all__ = ["Sheet", "Styled", "DropDown", "STYLES", "col_letter", "date_to_serial", "workbook_bytes", "save_workbook"]
