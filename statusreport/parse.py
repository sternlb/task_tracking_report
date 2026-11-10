"""Parsing of single cell values — design 5.3, R-IN-6, R-IN-8.

Every parser takes a raw cell value (see :mod:`statusreport.xlsx_read`) and
returns the parsed value or ``None`` for an empty cell. Invalid input raises
``ValueError`` with a message meant for the user.
"""

from __future__ import annotations

import datetime as dt
import re
from typing import List, Optional

from .xlsx_read import Percent, serial_to_date

_DATE_PATTERNS = (
    (re.compile(r"^(\d{4})-(\d{1,2})-(\d{1,2})"), ("y", "m", "d")),
    (re.compile(r"^(\d{1,2})\.(\d{1,2})\.(\d{4})"), ("d", "m", "y")),
    (re.compile(r"^(\d{1,2})/(\d{1,2})/(\d{4})"), ("d", "m", "y")),
)
_DURATION_PART = re.compile(r"(\d+(?:[.,]\d+)?)\s*([wdhm])", re.I)
_DURATION_FULL = re.compile(r"^\s*(?:\d+(?:[.,]\d+)?\s*[wdhm]\s*)+$", re.I)
_NUMBER = re.compile(r"^\s*[-+]?\d+(?:[.,]\d+)?\s*$")
_SERIAL_MIN, _SERIAL_MAX = 32874, 73051  # 1990-01-01 .. 2100-01-01

LIGHTS = ("green", "orange", "red")
_LIGHT_ALIASES = {
    "green": "green", "g": "green",
    "orange": "orange", "o": "orange", "amber": "orange", "yellow": "orange", "y": "orange",
    "red": "red", "r": "red",
}


def is_empty(value) -> bool:
    return value is None or (isinstance(value, str) and value.strip() == "")


def text(value) -> str:
    """Cell value as trimmed text. Whole numbers lose their ``.0`` (``123.0`` -> ``"123"``)."""
    if is_empty(value):
        return ""
    if isinstance(value, bool):
        return "TRUE" if value else "FALSE"
    if isinstance(value, float):
        return str(int(value)) if value.is_integer() else repr(value)
    if isinstance(value, dt.datetime):
        return value.isoformat(sep=" ")
    if isinstance(value, dt.date):
        return value.isoformat()
    return str(value).strip()


def number(value) -> Optional[float]:
    if is_empty(value):
        return None
    if isinstance(value, bool):
        raise ValueError(f"expected a number, got {text(value)}")
    if isinstance(value, (int, float)):
        return float(value)
    s = str(value).strip()
    if _NUMBER.match(s):
        return float(s.replace(",", "."))
    raise ValueError(f'"{s}" is not a number')


def date(value) -> Optional[dt.date]:
    """Excel date, ``YYYY-MM-DD``, ``DD.MM.YYYY`` or ``DD/MM/YYYY``; a time part is ignored."""
    if is_empty(value):
        return None
    if isinstance(value, dt.datetime):
        return value.date()
    if isinstance(value, dt.date):
        return value
    if isinstance(value, (int, float)) and not isinstance(value, bool):
        # A date serial in a cell without date format (e.g. pasted values).
        # Only plausible project dates (1990-2100) are accepted.
        if _SERIAL_MIN <= value <= _SERIAL_MAX:
            d = serial_to_date(float(value))
            return d.date() if isinstance(d, dt.datetime) else d
        raise ValueError(f"{text(value)} is not a date (use YYYY-MM-DD or DD.MM.YYYY)")
    s = str(value).strip()
    for pattern, order in _DATE_PATTERNS:
        m = pattern.match(s)
        if m:
            parts = dict(zip(order, (int(g) for g in m.groups())))
            try:
                return dt.date(parts["y"], parts["m"], parts["d"])
            except ValueError:
                break
    raise ValueError(f'"{s}" is not a date (use YYYY-MM-DD or DD.MM.YYYY)')


def hours(value, hours_per_day: float = 8.0) -> Optional[float]:
    """Effort in hours: a number, ``12 h``, or Polarion style ``1w 2d 4h 30m`` (1 w = 5 d)."""
    if is_empty(value):
        return None
    if isinstance(value, (int, float)) and not isinstance(value, bool):
        result = float(value)
    else:
        s = str(value).strip()
        if _NUMBER.match(s):
            result = float(s.replace(",", "."))
        elif _DURATION_FULL.match(s):
            factor = {"w": 5 * hours_per_day, "d": hours_per_day, "h": 1.0, "m": 1 / 60}
            result = sum(float(n.replace(",", ".")) * factor[u.lower()] for n, u in _DURATION_PART.findall(s))
        else:
            raise ValueError(f'"{s}" is not a duration (use hours, or e.g. "1d 4h")')
    if result < 0:
        raise ValueError(f"negative effort {result:g}")
    return result


def percent(value) -> Optional[float]:
    """Completion in % (0–100). ``65``, ``"65 %"`` and a 65 %-formatted cell all give 65."""
    if is_empty(value):
        return None
    if isinstance(value, Percent):
        result = float(value) * 100
    elif isinstance(value, (int, float)) and not isinstance(value, bool):
        result = float(value)
    else:
        s = str(value).strip().rstrip("%").strip()
        if not _NUMBER.match(s):
            raise ValueError(f'"{value}" is not a percentage')
        result = float(s.replace(",", "."))
    result = round(result, 6)
    if not 0 <= result <= 100:
        raise ValueError(f"{result:g} % is outside 0-100")
    return result


def light(value) -> Optional[str]:
    """``green``/``orange``/``red`` (also ``g``/``o``/``r``, ``amber``, ``yellow``)."""
    if is_empty(value):
        return None
    key = str(value).strip().lower()
    if key in _LIGHT_ALIASES:
        return _LIGHT_ALIASES[key]
    raise ValueError(f'"{value}" is not a traffic light (use Green, Orange or Red)')


def word_list(value) -> List[str]:
    """Comma- or semicolon-separated list, trimmed, empty entries dropped."""
    return [p.strip() for p in re.split(r"[,;\n]", text(value)) if p.strip()]


def natural_key(s: str):
    """Sort key so that ``1.2`` < ``1.10`` and ``v9`` < ``v10``."""
    return [(0, int(p), "") if p.isdigit() else (1, 0, p.lower()) for p in re.split(r"(\d+)", s) if p != ""]
