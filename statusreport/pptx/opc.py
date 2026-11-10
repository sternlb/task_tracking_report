"""Open Packaging Conventions: the zip container behind .pptx and .xlsx.

A package is a set of parts (zip entries) plus relationships (``_rels/*.rels``)
and a content-type table (``[Content_Types].xml``). Part names are stored
without a leading slash, e.g. ``ppt/presentation.xml``. OPC part names are
case-insensitive, so lookups compare lower-cased names.
"""

from __future__ import annotations

import io
import posixpath
import zipfile
from dataclasses import dataclass
from typing import Dict, List, Optional
from urllib.parse import unquote
from xml.etree import ElementTree as ET

from ..xmlutil import doc, esc

NS_CT = "http://schemas.openxmlformats.org/package/2006/content-types"
NS_RELS = "http://schemas.openxmlformats.org/package/2006/relationships"

RT = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
RT_OFFICE_DOCUMENT = RT + "/officeDocument"
RT_EXTENDED_PROPS = RT + "/extended-properties"
RT_CORE_PROPS = "http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties"
RT_SLIDE = RT + "/slide"
RT_SLIDE_LAYOUT = RT + "/slideLayout"
RT_SLIDE_MASTER = RT + "/slideMaster"
RT_NOTES_SLIDE = RT + "/notesSlide"
RT_THEME = RT + "/theme"
RT_PRES_PROPS = RT + "/presProps"
RT_VIEW_PROPS = RT + "/viewProps"
RT_TABLE_STYLES = RT + "/tableStyles"
RT_CHART = RT + "/chart"
RT_PACKAGE = RT + "/package"

CT_RELS = "application/vnd.openxmlformats-package.relationships+xml"
CT_XML = "application/xml"
CT_XLSX = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
_PML = "application/vnd.openxmlformats-officedocument.presentationml."
CT_PRESENTATION = _PML + "presentation.main+xml"
CT_SLIDE = _PML + "slide+xml"
CT_SLIDE_LAYOUT = _PML + "slideLayout+xml"
CT_SLIDE_MASTER = _PML + "slideMaster+xml"
CT_PRES_PROPS = _PML + "presProps+xml"
CT_VIEW_PROPS = _PML + "viewProps+xml"
CT_TABLE_STYLES = _PML + "tableStyles+xml"
CT_THEME = "application/vnd.openxmlformats-officedocument.theme+xml"
CT_CHART = "application/vnd.openxmlformats-officedocument.drawingml.chart+xml"
CT_CORE_PROPS = "application/vnd.openxmlformats-package.core-properties+xml"
CT_EXTENDED_PROPS = "application/vnd.openxmlformats-officedocument.extended-properties+xml"

CONTENT_TYPES = "[Content_Types].xml"

# Fixed zip timestamp so that identical input gives identical output (NF-2).
_ZIP_TIME = (1980, 1, 1, 0, 0, 0)


@dataclass
class Rel:
    rid: str
    type: str
    target: str  # as written in the .rels file (relative URI or external URL)
    external: bool = False


def rels_name(source: str) -> str:
    """Name of the .rels part that holds the relationships of ``source``.

    ``source`` is ``""`` for the package itself.
    """
    if not source:
        return "_rels/.rels"
    folder, base = posixpath.split(source)
    return posixpath.join(folder, "_rels", base + ".rels")


def source_of_rels(rels_part: str) -> str:
    """Inverse of :func:`rels_name`."""
    folder, base = posixpath.split(rels_part)
    parent = posixpath.dirname(folder)
    name = base[: -len(".rels")]
    return posixpath.join(parent, name) if name else ""


def resolve(source: str, target: str) -> str:
    """Turn a relationship target into a part name."""
    target = unquote(target.split("#", 1)[0])
    if target.startswith("/"):
        return posixpath.normpath(target[1:])
    return posixpath.normpath(posixpath.join(posixpath.dirname(source), target))


def relative(source: str, part: str) -> str:
    """Relationship target that points from ``source`` to ``part``."""
    return posixpath.relpath(part, posixpath.dirname(source) or ".")


class Package:
    def __init__(self) -> None:
        self.parts: Dict[str, bytes] = {}
        self.defaults: Dict[str, str] = {"rels": CT_RELS, "xml": CT_XML}
        self.overrides: Dict[str, str] = {}  # lower-cased part name -> content type

    # ---- loading and saving -------------------------------------------------

    @classmethod
    def from_bytes(cls, data: bytes) -> "Package":
        pkg = cls()
        with zipfile.ZipFile(io.BytesIO(data)) as zf:
            for info in zf.infolist():
                if info.is_dir():
                    continue
                pkg.parts[info.filename] = zf.read(info)
        ct = pkg.parts.pop(CONTENT_TYPES, None)
        if ct is None:
            raise ValueError("not an Office file: [Content_Types].xml missing")
        root = ET.fromstring(ct)
        for el in root:
            tag = el.tag.split("}")[-1]
            if tag == "Default":
                pkg.defaults[el.get("Extension", "").lower()] = el.get("ContentType", "")
            elif tag == "Override":
                name = el.get("PartName", "").lstrip("/")
                pkg.overrides[name.lower()] = el.get("ContentType", "")
        return pkg

    @classmethod
    def open(cls, path: str) -> "Package":
        with open(path, "rb") as fh:
            return cls.from_bytes(fh.read())

    def to_bytes(self) -> bytes:
        buf = io.BytesIO()
        with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zf:
            def put(name: str, data: bytes) -> None:
                zf.writestr(zipfile.ZipInfo(name, _ZIP_TIME), data, zipfile.ZIP_DEFLATED)

            put(CONTENT_TYPES, self._content_types_xml())
            order = sorted(self.parts, key=lambda n: (n != "_rels/.rels", n))
            for name in order:
                put(name, self.parts[name])
        return buf.getvalue()

    def save(self, path: str) -> None:
        data = self.to_bytes()
        with open(path, "wb") as fh:
            fh.write(data)

    def _content_types_xml(self) -> bytes:
        lines = [f'<Types xmlns="{NS_CT}">']
        for ext in sorted(self.defaults):
            lines.append(f'<Default Extension="{esc(ext)}" ContentType="{esc(self.defaults[ext])}"/>')
        for name in sorted(self.parts):
            ct = self.overrides.get(name.lower())
            if ct:
                lines.append(f'<Override PartName="/{esc(name)}" ContentType="{esc(ct)}"/>')
        lines.append("</Types>")
        return doc("".join(lines))

    # ---- parts ----------------------------------------------------------------

    def find(self, name: str) -> Optional[str]:
        """Actual stored name of a part (case-insensitive), or None."""
        if name in self.parts:
            return name
        low = name.lower()
        for existing in self.parts:
            if existing.lower() == low:
                return existing
        return None

    def content_type(self, name: str) -> Optional[str]:
        ct = self.overrides.get(name.lower())
        if ct:
            return ct
        ext = posixpath.splitext(name)[1][1:].lower()
        return self.defaults.get(ext)

    def add_part(self, name: str, data: bytes, content_type: Optional[str] = None) -> None:
        self.parts[name] = data
        if content_type:
            ext = posixpath.splitext(name)[1][1:].lower()
            if self.defaults.get(ext) != content_type:
                self.overrides[name.lower()] = content_type

    def remove_part(self, name: str) -> None:
        actual = self.find(name)
        if actual is None:
            return
        del self.parts[actual]
        self.overrides.pop(actual.lower(), None)
        rels = self.find(rels_name(actual))
        if rels:
            del self.parts[rels]

    def next_name(self, pattern: str) -> str:
        """First free part name for a pattern like ``ppt/slides/slide{}.xml``."""
        i = 1
        while self.find(pattern.format(i)) is not None:
            i += 1
        return pattern.format(i)

    # ---- relationships --------------------------------------------------------

    def rels(self, source: str) -> List[Rel]:
        name = self.find(rels_name(source))
        if name is None:
            return []
        out = []
        for el in ET.fromstring(self.parts[name]):
            out.append(
                Rel(
                    rid=el.get("Id", ""),
                    type=el.get("Type", ""),
                    target=el.get("Target", ""),
                    external=el.get("TargetMode") == "External",
                )
            )
        return out

    def set_rels(self, source: str, rels: List[Rel]) -> None:
        name = self.find(rels_name(source)) or rels_name(source)
        if not rels:
            self.parts.pop(name, None)
            return
        lines = [f'<Relationships xmlns="{NS_RELS}">']
        for r in rels:
            mode = ' TargetMode="External"' if r.external else ""
            lines.append(
                f'<Relationship Id="{esc(r.rid)}" Type="{esc(r.type)}" Target="{esc(r.target)}"{mode}/>'
            )
        lines.append("</Relationships>")
        self.parts[name] = doc("".join(lines))

    def add_rel(self, source: str, rel_type: str, part: str) -> str:
        """Add an internal relationship from ``source`` to ``part``; return its rId."""
        rels = self.rels(source)
        used = {r.rid for r in rels}
        n = len(rels) + 1
        while f"rId{n}" in used:
            n += 1
        rid = f"rId{n}"
        rels.append(Rel(rid, rel_type, relative(source, part)))
        self.set_rels(source, rels)
        return rid

    def related(self, source: str, rel_type: str) -> List[str]:
        """Part names that ``source`` points to with the given relationship type."""
        return [resolve(source, r.target) for r in self.rels(source) if r.type == rel_type and not r.external]

    def main_part(self) -> str:
        parts = self.related("", RT_OFFICE_DOCUMENT)
        if not parts:
            raise ValueError("package has no main document")
        return parts[0]

    # ---- garbage collection -----------------------------------------------------

    def reachable(self) -> set:
        """Lower-cased names of all parts reachable from the package root."""
        seen: set = set()
        stack = [""]
        while stack:
            src = stack.pop()
            for r in self.rels(src):
                if r.external:
                    continue
                target = resolve(src, r.target)
                actual = self.find(target)
                if actual is None or actual.lower() in seen:
                    continue
                seen.add(actual.lower())
                stack.append(actual)
        return seen

    def drop_unreachable(self) -> List[str]:
        """Remove parts no relationship chain leads to. Returns the removed names."""
        keep = self.reachable()
        removed = []
        for name in list(self.parts):
            if name.endswith(".rels"):
                continue
            if name.lower() not in keep:
                self.remove_part(name)
                removed.append(name)
        # .rels of removed parts are gone; drop orphaned .rels as well
        for name in list(self.parts):
            if name.endswith(".rels"):
                src = source_of_rels(name)
                if src and self.find(src) is None:
                    del self.parts[name]
        return removed
