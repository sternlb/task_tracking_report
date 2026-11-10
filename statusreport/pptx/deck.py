"""Deck assembly on top of a base .pptx (design 8.1, decision D-2).

Opening a base removes all of its slides; layouts, masters and the theme stay.
Slides are then added on layouts chosen by name, so the same code serves the
built-in neutral base and a customer template.
"""

from __future__ import annotations

import pkgutil
import re
from dataclasses import dataclass
from typing import Dict, List, Optional, Sequence
from xml.etree import ElementTree as ET

from ..xmlutil import NS_A, NS_P, NS_R, PML_NS, doc
from . import opc
from .chart import Series, chart_workbook, line_chart_xml
from .lint import lint_package
from .shapes import Box, Para, Row, chart_frame, inch, placeholder_sp, table_frame, textbox_sp

_P = "{%s}" % NS_P
_A = "{%s}" % NS_A

_GRP = (
    '<p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>'
    '<p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/>'
    '<a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>'
)
_SECTION_EXT = "{521415D9-36F7-43E2-AB2F-B90AF26B5E84}"
_TITLE_TYPES = ("title", "ctrTitle")
MARGIN = inch(0.4)


@dataclass
class Placeholder:
    type: str  # "title", "ctrTitle", "subTitle", "body", "obj", ...
    idx: str
    attrs: Dict[str, str]  # attributes of <p:ph> as written in the layout
    box: Optional[Box]


@dataclass
class Layout:
    name: str
    part: str
    placeholders: List[Placeholder]

    def find(self, *types: str) -> Optional[Placeholder]:
        for ph in self.placeholders:
            if ph.type in types:
                return ph
        return None


def _placeholders(root: ET.Element) -> List[Placeholder]:
    out = []
    for sp in root.iter(_P + "sp"):
        ph = sp.find(f"{_P}nvSpPr/{_P}nvPr/{_P}ph")
        if ph is None:
            continue
        box = None
        off = sp.find(f"{_P}spPr/{_A}xfrm/{_A}off")
        ext = sp.find(f"{_P}spPr/{_A}xfrm/{_A}ext")
        if off is not None and ext is not None:
            box = Box(int(off.get("x")), int(off.get("y")), int(ext.get("cx")), int(ext.get("cy")))
        out.append(Placeholder(ph.get("type", "obj"), ph.get("idx", "0"), dict(ph.attrib), box))
    return out


def _master_kind(ph_type: str) -> str:
    if ph_type in _TITLE_TYPES:
        return "title"
    if ph_type in ("body", "obj", "subTitle", "tbl", "chart", "pic", "dgm", "media", "clipArt"):
        return "body"
    return ph_type


class Slide:
    def __init__(self, deck: "Deck", layout: Layout, part: str) -> None:
        self.deck = deck
        self.layout = layout
        self.part = part
        self._shapes: List[str] = []
        self._next_id = 2

    def _id(self) -> int:
        sid = self._next_id
        self._next_id += 1
        return sid

    # ---- placeholders ---------------------------------------------------------

    def _fill(self, ph: Optional[Placeholder], name: str, paras: Sequence[Para]) -> bool:
        if ph is None:
            return False
        attrs = {k: v for k, v in ph.attrs.items() if k in ("type", "idx", "orient", "sz")}
        self._shapes.append(placeholder_sp(self._id(), name, attrs, paras))
        return True

    def set_title(self, paras: Sequence[Para]) -> bool:
        return self._fill(self.layout.find(*_TITLE_TYPES), "Title", paras)

    def set_subtitle(self, paras: Sequence[Para]) -> bool:
        return self._fill(self.layout.find("subTitle"), "Subtitle", paras)

    # ---- free shapes ------------------------------------------------------------

    def content_area(self) -> Box:
        return self.deck.content_area(self.layout)

    def add_textbox(self, box: Box, paras: Sequence[Para], anchor: str = "t") -> None:
        sid = self._id()
        self._shapes.append(textbox_sp(sid, f"TextBox {sid}", box, paras, anchor))

    def add_table(self, box: Box, col_widths: Sequence[int], rows: Sequence[Row]) -> None:
        sid = self._id()
        self._shapes.append(table_frame(sid, f"Table {sid}", box, col_widths, rows))

    def add_line_chart(self, box: Box, categories: Sequence[str], series: Sequence[Series], **kw) -> None:
        pkg = self.deck.pkg
        chart = pkg.next_name("ppt/charts/chart{}.xml")
        book = pkg.next_name("ppt/embeddings/Microsoft_Excel_Worksheet{}.xlsx")
        pkg.add_part(book, chart_workbook(categories, series), opc.CT_XLSX)
        book_rid = pkg.add_rel(chart, opc.RT_PACKAGE, book)
        pkg.add_part(chart, line_chart_xml(categories, series, book_rid, **kw), opc.CT_CHART)
        rid = pkg.add_rel(self.part, opc.RT_CHART, chart)
        sid = self._id()
        self._shapes.append(chart_frame(sid, f"Chart {sid}", box, rid))

    def xml(self) -> bytes:
        return doc(
            f"<p:sld {PML_NS}><p:cSld><p:spTree>{_GRP}{''.join(self._shapes)}</p:spTree></p:cSld>"
            "<p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sld>"
        )


class Deck:
    def __init__(self, data: bytes) -> None:
        self.pkg = opc.Package.from_bytes(data)
        self.pkg.defaults.setdefault("xlsx", opc.CT_XLSX)
        self.pres = self.pkg.main_part()
        if self.pkg.content_type(self.pres) != opc.CT_PRESENTATION:
            raise ValueError("not a PowerPoint presentation (.pptx)")
        size = re.search(rb'<(?:\w+:)?sldSz\b[^>]*\bcx="(\d+)"[^>]*\bcy="(\d+)"', self.pkg.parts[self.pres])
        self.width, self.height = (int(size.group(1)), int(size.group(2))) if size else (12192000, 6858000)
        self._strip_slides()
        self.layouts = self._read_layouts()
        self.slides: List[Slide] = []
        self.warnings: List[str] = []

    @classmethod
    def default(cls) -> "Deck":
        data = pkgutil.get_data(__package__, "base.pptx")
        if data is None:
            raise FileNotFoundError("built-in base.pptx not found")
        return cls(data)

    @classmethod
    def from_file(cls, path: str) -> "Deck":
        with open(path, "rb") as fh:
            return cls(fh.read())

    # ---- base preparation -------------------------------------------------------

    def _strip_slides(self) -> None:
        pkg, pres = self.pkg, self.pres
        pkg.set_rels(pres, [r for r in pkg.rels(pres) if r.type != opc.RT_SLIDE])
        xml = pkg.parts[pres].decode("utf-8")
        for tag in ("sldIdLst", "custShowLst"):
            xml = re.sub(rf"<(\w+:)?{tag}\b[^>]*/>|<(\w+:)?{tag}\b.*?</(\w+:)?{tag}>", "", xml, flags=re.S)
        xml = re.sub(
            rf'<(\w+:)?ext uri="{re.escape(_SECTION_EXT)}">.*?</\1ext>', "", xml, flags=re.S
        )
        pkg.parts[pres] = xml.encode("utf-8")
        pkg.drop_unreachable()

    def _read_layouts(self) -> Dict[str, Layout]:
        pkg = self.pkg
        layouts: Dict[str, Layout] = {}
        for master in pkg.related(self.pres, opc.RT_SLIDE_MASTER):
            master_ph = _placeholders(ET.fromstring(pkg.parts[pkg.find(master)]))
            for part in pkg.related(master, opc.RT_SLIDE_LAYOUT):
                actual = pkg.find(part)
                if actual is None:
                    continue
                root = ET.fromstring(pkg.parts[actual])
                csld = root.find(_P + "cSld")
                name = csld.get("name", "") if csld is not None else ""
                phs = _placeholders(root)
                for ph in phs:  # inherit position from the master
                    if ph.box is None:
                        kind = _master_kind(ph.type)
                        src = next((m for m in master_ph if _master_kind(m.type) == kind and m.box), None)
                        ph.box = src.box if src else None
                layouts.setdefault(name, Layout(name, actual, phs))
        if not layouts:
            raise ValueError("template has no slide layouts")
        return layouts

    # ---- layouts and areas --------------------------------------------------------

    def layout(self, name: str) -> Layout:
        if name in self.layouts:
            return self.layouts[name]
        low = name.lower()
        for key, lay in self.layouts.items():
            if key.lower() == low:
                return lay
        fallback = next((l for l in self.layouts.values() if l.find(*_TITLE_TYPES)), next(iter(self.layouts.values())))
        self.warnings.append(f'layout "{name}" not found in template, using "{fallback.name}"')
        return fallback

    def content_area(self, layout: Layout) -> Box:
        body = next(
            (p for p in layout.placeholders if p.idx == "1" and p.type in ("obj", "body") and p.box), None
        )
        if body:
            return body.box
        title = layout.find(*_TITLE_TYPES)
        if title and title.box:
            top = title.box.y + title.box.h + inch(0.2)
            return Box(title.box.x, top, title.box.w, self.height - top - MARGIN)
        return Box(MARGIN, MARGIN, self.width - 2 * MARGIN, self.height - 2 * MARGIN)

    # ---- slides ---------------------------------------------------------------------

    def add_slide(self, layout_name: str) -> Slide:
        layout = self.layout(layout_name)
        part = self.pkg.next_name("ppt/slides/slide{}.xml")
        self.pkg.add_part(part, b"", opc.CT_SLIDE)  # reserve the name
        self.pkg.add_rel(part, opc.RT_SLIDE_LAYOUT, layout.part)
        slide = Slide(self, layout, part)
        self.slides.append(slide)
        return slide

    # ---- output ---------------------------------------------------------------------

    def _write_slide_list(self) -> None:
        pkg, pres = self.pkg, self.pres
        entries = []
        for i, slide in enumerate(self.slides):
            pkg.parts[slide.part] = slide.xml()
            rid = pkg.add_rel(pres, opc.RT_SLIDE, slide.part)
            entries.append((256 + i, rid))
        xml = pkg.parts[pres].decode("utf-8")
        root_tag = re.search(r"<(\w+:)?presentation\b[^>]*>", xml)
        p = root_tag.group(1) or ""
        r = re.search(r'xmlns:(\w+)="%s"' % re.escape(NS_R), root_tag.group(0))
        if r is None:
            xml = xml.replace(root_tag.group(0), root_tag.group(0)[:-1] + f' xmlns:r="{NS_R}">', 1)
            r_prefix = "r"
        else:
            r_prefix = r.group(1)
        if entries:
            ids = "".join(f'<{p}sldId id="{sid}" {r_prefix}:id="{rid}"/>' for sid, rid in entries)
            block = f"<{p}sldIdLst>{ids}</{p}sldIdLst>"
            ends = [m.end() for m in re.finditer(rf"</{p}(sldMasterIdLst|notesMasterIdLst|handoutMasterIdLst)>", xml)]
            ends += [m.end() for m in re.finditer(rf"<{p}(notesMasterIdLst|handoutMasterIdLst)\b[^>]*/>", xml)]
            if not ends:
                raise ValueError("presentation.xml has no slide master list")
            pos = max(ends)
            xml = xml[:pos] + block + xml[pos:]
        pkg.parts[pres] = xml.encode("utf-8")

        app = pkg.find("docProps/app.xml")
        if app:
            pkg.parts[app] = doc(
                '<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" '
                'xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">'
                f"<Application>sw-status-report</Application><Slides>{len(self.slides)}</Slides></Properties>"
            )

    def to_bytes(self, check: bool = True) -> bytes:
        if getattr(self, "_written", False):
            raise RuntimeError("deck already written; build a new Deck for another output")
        self._written = True
        self._write_slide_list()
        if check:
            problems = lint_package(self.pkg)
            if problems:
                raise ValueError("generated deck failed the package check:\n  " + "\n  ".join(problems))
        return self.pkg.to_bytes()

    def save(self, path: str, check: bool = True) -> None:
        data = self.to_bytes(check)
        with open(path, "wb") as fh:
            fh.write(data)
