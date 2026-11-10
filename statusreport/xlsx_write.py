"""Minimal .xlsx writer (standard library only).

M1 scope: plain values for the workbook embedded in a chart. Styles,
column widths and drop-downs for the input template follow in M2 (T-2.2).
"""

from __future__ import annotations

from typing import List, Optional, Sequence, Tuple, Union

from .pptx import opc
from .xmlutil import NS_R, NS_S, doc, esc

Value = Union[str, int, float, None]

_SML = "application/vnd.openxmlformats-officedocument.spreadsheetml."
CT_WORKBOOK = _SML + "sheet.main+xml"
CT_WORKSHEET = _SML + "worksheet+xml"
CT_STYLES = _SML + "styles+xml"
RT_WORKSHEET = opc.RT + "/worksheet"
RT_STYLES = opc.RT + "/styles"

_STYLES = (
    f'<styleSheet xmlns="{NS_S}">'
    '<fonts count="1"><font><sz val="11"/><name val="Calibri"/><family val="2"/></font></fonts>'
    '<fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills>'
    '<borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>'
    '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>'
    '<cellXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/></cellXfs>'
    '<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>'
    "</styleSheet>"
)


def col_letter(index: int) -> str:
    """0 -> A, 25 -> Z, 26 -> AA."""
    out = ""
    index += 1
    while index:
        index, rem = divmod(index - 1, 26)
        out = chr(65 + rem) + out
    return out


def _cell(ref: str, value: Value) -> str:
    if value is None or value == "":
        return ""
    if isinstance(value, bool):
        return f'<c r="{ref}" t="b"><v>{int(value)}</v></c>'
    if isinstance(value, (int, float)):
        return f'<c r="{ref}"><v>{value!r}</v></c>'
    text = esc(value)
    space = ' xml:space="preserve"' if text != text.strip() else ""
    return f'<c r="{ref}" t="inlineStr"><is><t{space}>{text}</t></is></c>'


def _sheet_xml(rows: Sequence[Sequence[Value]]) -> bytes:
    out = []
    for r, row in enumerate(rows, 1):
        cells = "".join(_cell(f"{col_letter(c)}{r}", v) for c, v in enumerate(row))
        out.append(f'<row r="{r}">{cells}</row>')
    return doc(f'<worksheet xmlns="{NS_S}"><sheetData>{"".join(out)}</sheetData></worksheet>')


def workbook_bytes(sheets: List[Tuple[str, Sequence[Sequence[Optional[Value]]]]]) -> bytes:
    pkg = opc.Package()
    wb = "xl/workbook.xml"
    pkg.add_rel("", opc.RT_OFFICE_DOCUMENT, wb)
    entries = []
    for i, (name, rows) in enumerate(sheets, 1):
        part = f"xl/worksheets/sheet{i}.xml"
        pkg.add_part(part, _sheet_xml(rows), CT_WORKSHEET)
        rid = pkg.add_rel(wb, RT_WORKSHEET, part)
        entries.append(f'<sheet name="{esc(name)}" sheetId="{i}" r:id="{rid}"/>')
    pkg.add_part("xl/styles.xml", doc(_STYLES), CT_STYLES)
    pkg.add_rel(wb, RT_STYLES, "xl/styles.xml")
    pkg.add_part(
        wb,
        doc(f'<workbook xmlns="{NS_S}" xmlns:r="{NS_R}"><sheets>{"".join(entries)}</sheets></workbook>'),
        CT_WORKBOOK,
    )
    return pkg.to_bytes()
