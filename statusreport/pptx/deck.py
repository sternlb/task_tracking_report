"""Deck assembly on top of a base .pptx or .potx (design 8.1, decision D-2, R-OUT-8).

Opening a base removes all of its slides; layouts, masters and the theme stay.
Slides are then added on layouts chosen by name, with a language-independent
fallback by layout type, so the same code serves the built-in neutral base and
any customer template.
"""

from __future__ import annotations

import pkgutil
import re
from dataclasses import dataclass
from typing import Dict, List, Optional, Sequence, Tuple
from xml.etree import ElementTree as ET

from ..xmlutil import NS_A, NS_P, NS_R, PML_NS, doc, esc
from . import opc
from .chart import Series, chart_workbook, line_chart_xml
from .lint import lint_package
from .shapes import Box, Para, Row, Run, chart_frame, inch, placeholder_sp, table_frame, textbox_sp

_P = "{%s}" % NS_P
_A = "{%s}" % NS_A

CT_TEMPLATE = "application/vnd.openxmlformats-officedocument.presentationml.template.main+xml"

_GRP = (
    '<p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>'
    '<p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/>'
    '<a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>'
)
_SECTION_EXT = "{521415D9-36F7-43E2-AB2F-B90AF26B5E84}"
_SLIDENUM_FIELD_ID = "{B6F15528-21DE-4FAA-801E-634DDDAF4B2B}"
_TITLE_TYPES = ("title", "ctrTitle")
_FOOTER_TYPES = ("dt", "ftr", "sldNum")
MARGIN = inch(0.4)
FOOTER_GAP = inch(0.1)

# Fallback when no layout has the configured name (names depend on the template's
# language: "Title Only" is "Nur Titel" in a German Office). Layout types are fixed.
_ROLE_TYPES = {
    "title": ("title",),
    "content": ("titleOnly", "obj", "tx"),
}


@dataclass
class Placeholder:
    type: str  # "title", "ctrTitle", "subTitle", "body", "obj", "sldNum", ...
    idx: str
    attrs: Dict[str, str]  # attributes of <p:ph> as written in the layout
    box: Optional[Box]
    text: str = ""  # text in the layout (e.g. a fixed footer)


@dataclass
class Layout:
    name: str
    part: str
    placeholders: List[Placeholder]
    type: str = "cust"  # sldLayout/@type: title, titleOnly, obj, ... (language independent)

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
        ph_type = ph.get("type", "obj")
        text = ""
        if ph_type != "sldNum":  # the slide number is a field, not text
            text = "".join(t.text or "" for t in sp.iter(_A + "t")).strip()
        out.append(Placeholder(ph_type, ph.get("idx", "0"), dict(ph.attrib), box, text))
    return out


def _master_kind(ph_type: str) -> str:
    if ph_type in _TITLE_TYPES:
        return "title"
    if ph_type in ("body", "obj", "subTitle", "tbl", "chart", "pic", "dgm", "media", "clipArt"):
        return "body"
    return ph_type


def _ph_attrs(ph: Placeholder) -> Dict[str, str]:
    return {k: v for k, v in ph.attrs.items() if k in ("type", "idx", "orient", "sz")}


class Slide:
    def __init__(self, deck: "Deck", layout: Layout, part: str, number: int) -> None:
        self.deck = deck
        self.layout = layout
        self.part = part
        self.number = number
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
        self._shapes.append(placeholder_sp(self._id(), name, _ph_attrs(ph), paras))
        return True

    def set_title(self, paras: Sequence[Para]) -> bool:
        return self._fill(self.layout.find(*_TITLE_TYPES), "Title", paras)

    def set_subtitle(self, paras: Sequence[Para]) -> bool:
        return self._fill(self.layout.find("subTitle"), "Subtitle", paras)

    def add_footers(self, footer_text: str = "") -> None:
        """Slide number and footer, if the layout has placeholders for them (as PowerPoint does).

        The footer gets ``footer_text``, or else the layout's own footer text; an empty
        footer is left out. The date placeholder is not used.
        """
        for ph in self.layout.placeholders:
            if ph.type == "sldNum":
                attrs = "".join(f' {k}="{esc(v)}"' for k, v in _ph_attrs(ph).items())
                sid = self._id()
                self._shapes.append(
                    f'<p:sp><p:nvSpPr><p:cNvPr id="{sid}" name="Slide Number {sid}"/>'
                    f'<p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr><p:nvPr><p:ph{attrs}/></p:nvPr></p:nvSpPr>'
                    "<p:spPr/><p:txBody><a:bodyPr/><a:lstStyle/><a:p>"
                    f'<a:fld id="{_SLIDENUM_FIELD_ID}" type="slidenum"><a:rPr lang="en-US"/><a:t>{self.number}</a:t></a:fld>'
                    '<a:endParaRPr lang="en-US"/></a:p></p:txBody></p:sp>'
                )
            elif ph.type == "ftr":
                text = footer_text or ph.text
                if text:
                    self._fill(ph, "Footer", [Para([Run(text)])])

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
        ct = self.pkg.content_type(self.pres)
        if ct == CT_TEMPLATE:  # a .potx: the output is a normal presentation
            self.pkg.overrides[self.pres.lower()] = opc.CT_PRESENTATION
        elif ct != opc.CT_PRESENTATION:
            raise ValueError("not a PowerPoint presentation or template (.pptx / .potx)")
        size = re.search(rb'<(?:\w+:)?sldSz\b[^>]*\bcx="(\d+)"[^>]*\bcy="(\d+)"', self.pkg.parts[self.pres])
        self.width, self.height = (int(size.group(1)), int(size.group(2))) if size else (12192000, 6858000)
        self._strip_slides()
        self.layouts = self._read_layouts()
        self.slides: List[Slide] = []
        self.warnings: List[str] = []
        self._chosen: Dict[Tuple[str, str], Layout] = {}

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

    def _read_layouts(self) -> List[Layout]:
        pkg = self.pkg
        layouts: List[Layout] = []
        for master in pkg.related(self.pres, opc.RT_SLIDE_MASTER):
            master_ph = _placeholders(ET.fromstring(pkg.parts[pkg.find(master)]))
            for part in pkg.related(master, opc.RT_SLIDE_LAYOUT):
                actual = pkg.find(part)
                if actual is None:
                    continue
                root = ET.fromstring(pkg.parts[actual])
                csld = root.find(_P + "cSld")
                name = " ".join((csld.get("name", "") if csld is not None else "").split())
                phs = _placeholders(root)
                for ph in phs:  # inherit position from the master
                    if ph.box is None:
                        kind = _master_kind(ph.type)
                        src = next((m for m in master_ph if _master_kind(m.type) == kind and m.box), None)
                        ph.box = src.box if src else None
                layouts.append(Layout(name, actual, phs, root.get("type", "cust")))
        if not layouts:
            raise ValueError("template has no slide layouts")
        return layouts

    # ---- layouts and areas --------------------------------------------------------

    @property
    def layout_names(self) -> List[str]:
        return [l.name for l in self.layouts]

    def layout(self, name: str, role: str = "content") -> Layout:
        """Layout by name (case-insensitive); otherwise by layout type for the role.

        ``role`` is ``"title"`` (title slide) or ``"content"`` (chart and table slides).
        A fallback adds one warning per name and role.
        """
        key = (name, role)
        if key in self._chosen:
            return self._chosen[key]
        want = " ".join(name.split()).lower()
        found = next((l for l in self.layouts if l.name.lower() == want), None)
        if found is None:
            found = self._fallback(role)
            names = [f'"{n}"' for n in self.layout_names]
            more = f" and {len(names) - 10} more" if len(names) > 10 else ""
            self.warnings.append(
                f'layout "{name}" not found, using "{found.name}" (set "Title layout" / "Content layout" '
                f"in Settings to choose). Layouts: {', '.join(names[:10])}{more}"
            )
        self._chosen[key] = found
        return found

    def _fallback(self, role: str) -> Layout:
        for layout_type in _ROLE_TYPES.get(role, ()):
            hit = next((l for l in self.layouts if l.type == layout_type), None)
            if hit:
                return hit
        if role == "title":
            hit = next((l for l in self.layouts if l.find("ctrTitle")), None)
        else:  # a layout with a title that is not a title slide, preferring few placeholders
            candidates = [l for l in self.layouts if l.find("title") and not l.find("ctrTitle")]
            hit = min(candidates, key=lambda l: len([p for p in l.placeholders if p.type not in _FOOTER_TYPES]), default=None)
        return hit or next((l for l in self.layouts if l.find(*_TITLE_TYPES)), self.layouts[0])

    def content_area(self, layout: Layout) -> Box:
        """Where tables and charts go: the layout's body placeholder, else below the title.

        The area never reaches into the footer zone (date, footer, slide number).
        """
        body = next(
            (p for p in layout.placeholders if p.idx == "1" and p.type in ("obj", "body") and p.box), None
        )
        if body:
            area = Box(body.box.x, body.box.y, body.box.w, body.box.h)
        else:
            title = layout.find(*_TITLE_TYPES)
            if title and title.box:
                top = title.box.y + title.box.h + inch(0.2)
                area = Box(title.box.x, top, title.box.w, self.height - top - MARGIN)
            else:
                area = Box(MARGIN, MARGIN, self.width - 2 * MARGIN, self.height - 2 * MARGIN)
        footer_tops = [
            p.box.y for p in layout.placeholders
            if p.type in _FOOTER_TYPES and p.box and p.box.y > area.y + area.h // 2
        ]
        if footer_tops:
            bottom = min(footer_tops) - FOOTER_GAP
            area.h = max(inch(1.0), min(area.h, bottom - area.y))
        return area

    # ---- slides ---------------------------------------------------------------------

    def add_slide(self, layout_name: str, role: str = "content", footer_text: str = "") -> Slide:
        """New slide on the layout; content slides get slide number and footer placeholders."""
        layout = self.layout(layout_name, role)
        part = self.pkg.next_name("ppt/slides/slide{}.xml")
        self.pkg.add_part(part, b"", opc.CT_SLIDE)  # reserve the name
        self.pkg.add_rel(part, opc.RT_SLIDE_LAYOUT, layout.part)
        slide = Slide(self, layout, part, len(self.slides) + 1)
        if role != "title":  # by convention no slide number on the title slide
            slide.add_footers(footer_text)
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
