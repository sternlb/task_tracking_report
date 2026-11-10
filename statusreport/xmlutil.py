"""Small helpers for writing Office Open XML as text."""

from __future__ import annotations

from xml.sax.saxutils import escape as _escape

XML_DECL = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n'

NS_A = "http://schemas.openxmlformats.org/drawingml/2006/main"
NS_R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
NS_P = "http://schemas.openxmlformats.org/presentationml/2006/main"
NS_C = "http://schemas.openxmlformats.org/drawingml/2006/chart"
NS_S = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"

# Namespace declarations for the root element of PresentationML parts.
PML_NS = f'xmlns:a="{NS_A}" xmlns:r="{NS_R}" xmlns:p="{NS_P}"'


def esc(text: object) -> str:
    """Escape text for element content and double-quoted attributes."""
    return _escape(str(text), {'"': "&quot;"})


def doc(body: str) -> bytes:
    """Prefix the XML declaration and encode as UTF-8."""
    return (XML_DECL + body).encode("utf-8")
