"""Minimal .xlsx reader (standard library only) — design 5.1, T-2.1.

Reads cell values of every sheet into a list of rows. Values are Python types:

- ``str`` for text (shared strings, inline strings, formula string results)
- ``float`` for numbers, :class:`Percent` for numbers with a percent format
- ``datetime.date`` / ``datetime.datetime`` for numbers with a date format
- ``bool`` for booleans, ``None`` for empty cells and error values (#N/A, …)

Empty cells keep their column position; trailing empty cells and rows are trimmed.
"""

from __future__ import annotations

import datetime as dt
import re
import zipfile
from typing import Dict, List, Optional
from xml.etree import ElementTree as ET

from .pptx import opc

# Excel built-in number formats that show dates or times
_DATE_FORMAT_IDS = set(range(14, 23)) | {27, 30, 36, 45, 46, 47, 50, 57}
_PERCENT_FORMAT_IDS = {9, 10}
_REF = re.compile(r"([A-Z]+)(\d+)")
_OLE_MAGIC = b"\xd0\xcf\x11\xe0"


class XlsxError(Exception):
    """The file cannot be read as an .xlsx workbook."""


class Percent(float):
    """A number shown with a percent format in Excel (0.65 displayed as 65 %)."""


def _local(tag: str) -> str:
    return tag.rsplit("}", 1)[-1]


def _children(el: ET.Element, name: str):
    return [c for c in el if _local(c.tag) == name]


def _child(el: ET.Element, name: str) -> Optional[ET.Element]:
    for c in el:
        if _local(c.tag) == name:
            return c
    return None


def _text_of(el: ET.Element) -> str:
    """Concatenated <t> text of a string item, without phonetic runs (<rPh>)."""
    parts = []
    for c in el:
        name = _local(c.tag)
        if name == "t":
            parts.append(c.text or "")
        elif name == "r":
            t = _child(c, "t")
            if t is not None:
                parts.append(t.text or "")
    return "".join(parts)


def col_index(letters: str) -> int:
    """A -> 0, Z -> 25, AA -> 26."""
    n = 0
    for ch in letters:
        n = n * 26 + (ord(ch) - 64)
    return n - 1


def is_date_format(code: str) -> bool:
    """True if a custom number format shows a date or time."""
    code = re.sub(r'"[^"]*"|\[[^\]]*\]|\\.|_.|\*.', "", code)  # drop literals, colours, escapes
    code = code.split(";")[0].lower()
    return bool(re.search(r"[dmyhs]", code)) and "general" not in code


def serial_to_date(value: float, date1904: bool = False):
    base = dt.datetime(1904, 1, 1) if date1904 else dt.datetime(1899, 12, 30)
    stamp = base + dt.timedelta(days=value)
    stamp = stamp.replace(microsecond=0)
    if stamp.hour == stamp.minute == stamp.second == 0:
        return stamp.date()
    return stamp


class Workbook:
    def __init__(self, data: bytes, name: str = "workbook") -> None:
        if data[:4] == _OLE_MAGIC:
            raise XlsxError(f"{name} is an old .xls file - please open it in Excel and save it as .xlsx")
        try:
            self.pkg = opc.Package.from_bytes(data)
        except (zipfile.BadZipFile, ValueError) as exc:
            raise XlsxError(f"{name} is not an Excel .xlsx file ({exc})") from None
        try:
            self._book = self.pkg.main_part()
            root = ET.fromstring(self.pkg.parts[self.pkg.find(self._book)])
        except (ValueError, KeyError, TypeError, ET.ParseError) as exc:
            raise XlsxError(f"{name} has no readable workbook part ({exc})") from None
        pr = _child(root, "workbookPr")
        self.date1904 = pr is not None and pr.get("date1904", "0").lower() in ("1", "true")
        rels = {r.rid: r for r in self.pkg.rels(self._book)}
        self._sheets: Dict[str, str] = {}
        sheets = _child(root, "sheets")
        for sh in (_children(sheets, "sheet") if sheets is not None else []):
            rid = next((v for k, v in sh.attrib.items() if _local(k) == "id"), "")
            rel = rels.get(rid)
            if rel is not None:
                self._sheets[sh.get("name", "")] = opc.resolve(self._book, rel.target)
        self._strings = self._read_shared_strings(rels)
        self._formats = self._read_styles(rels)
        self._cache: Dict[str, List[list]] = {}

    @classmethod
    def open(cls, path: str) -> "Workbook":
        try:
            with open(path, "rb") as fh:
                data = fh.read()
        except OSError as exc:
            raise XlsxError(f"cannot open {path}: {exc.strerror}") from None
        return cls(data, path)

    # ---- workbook-level parts -------------------------------------------------------

    def _part_by_type(self, rels, suffix: str) -> Optional[bytes]:
        for r in rels.values():
            if r.type.endswith(suffix) and not r.external:
                name = self.pkg.find(opc.resolve(self._book, r.target))
                if name:
                    return self.pkg.parts[name]
        return None

    def _read_shared_strings(self, rels) -> List[str]:
        data = self._part_by_type(rels, "/sharedStrings")
        if not data:
            return []
        return [_text_of(si) for si in ET.fromstring(data) if _local(si.tag) == "si"]

    def _read_styles(self, rels) -> List[str]:
        """Per cell style index: 'date', 'percent' or ''."""
        data = self._part_by_type(rels, "/styles")
        if not data:
            return []
        root = ET.fromstring(data)
        custom: Dict[int, str] = {}
        fmts = _child(root, "numFmts")
        for f in (_children(fmts, "numFmt") if fmts is not None else []):
            custom[int(f.get("numFmtId", "0"))] = f.get("formatCode", "")
        kinds = []
        xfs = _child(root, "cellXfs")
        for xf in (_children(xfs, "xf") if xfs is not None else []):
            fid = int(xf.get("numFmtId", "0"))
            code = custom.get(fid)
            if fid in _DATE_FORMAT_IDS or (code is not None and is_date_format(code)):
                kinds.append("date")
            elif fid in _PERCENT_FORMAT_IDS or (code is not None and "%" in re.sub(r'"[^"]*"', "", code)):
                kinds.append("percent")
            else:
                kinds.append("")
        return kinds

    # ---- sheets --------------------------------------------------------------------------

    @property
    def sheet_names(self) -> List[str]:
        return list(self._sheets)

    def find_sheet(self, name: str) -> Optional[str]:
        """Actual sheet name, matched case-insensitively and ignoring surrounding spaces."""
        want = name.strip().lower()
        return next((n for n in self._sheets if n.strip().lower() == want), None)

    def rows(self, sheet: str) -> List[list]:
        if sheet in self._cache:
            return self._cache[sheet]
        part = self.pkg.find(self._sheets[sheet])
        if part is None:
            raise XlsxError(f'sheet "{sheet}" is missing in the file')
        root = ET.fromstring(self.pkg.parts[part])
        data = _child(root, "sheetData")
        out: List[list] = []
        next_row = 1
        for row in (_children(data, "row") if data is not None else []):
            r = int(row.get("r", next_row))
            next_row = r + 1
            values: list = []
            next_col = 0
            for c in _children(row, "c"):
                m = _REF.match(c.get("r", ""))
                col = col_index(m.group(1)) if m else next_col
                next_col = col + 1
                value = self._value(c)
                if value is None:
                    continue
                values.extend([None] * (col + 1 - len(values)))
                values[col] = value
            if values:
                out.extend([[] for _ in range(r - 1 - len(out))])
                out.append(values)
        self._cache[sheet] = out
        return out

    def _value(self, c: ET.Element):
        t = c.get("t", "n")
        v = _child(c, "v")
        raw = v.text if v is not None else None
        if t == "inlineStr":
            is_ = _child(c, "is")
            text = _text_of(is_) if is_ is not None else ""
            return text if text != "" else None
        if raw is None or raw == "":
            return None
        if t == "s":
            idx = int(raw)
            text = self._strings[idx] if idx < len(self._strings) else ""
            return text if text != "" else None
        if t in ("str", "e"):
            return None if t == "e" else raw
        if t == "b":
            return raw.strip() in ("1", "true")
        if t == "d":
            try:
                return dt.datetime.fromisoformat(raw.rstrip("Z")).date()
            except ValueError:
                return raw
        try:
            number = float(raw)
        except ValueError:
            return raw
        style = int(c.get("s", "0"))
        kind = self._formats[style] if style < len(self._formats) else ""
        if kind == "date":
            try:
                return serial_to_date(number, self.date1904)
            except OverflowError:
                return number
        if kind == "percent":
            return Percent(number)
        return number
