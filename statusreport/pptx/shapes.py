"""DrawingML / PresentationML shape XML: text, placeholders, tables, chart frames.

All functions return XML text fragments. Positions and sizes are EMU
(914400 per inch, 12700 per point).
"""

from __future__ import annotations

from dataclasses import dataclass, field
from typing import Dict, List, Optional, Sequence

from ..xmlutil import NS_C, esc

EMU_PER_INCH = 914400
EMU_PER_PT = 12700


def inch(value: float) -> int:
    return int(round(value * EMU_PER_INCH))


@dataclass
class Box:
    x: int
    y: int
    w: int
    h: int

    def xfrm(self, tag: str = "a:xfrm") -> str:
        return f'<{tag}><a:off x="{self.x}" y="{self.y}"/><a:ext cx="{self.w}" cy="{self.h}"/></{tag}>'


@dataclass
class Run:
    text: str
    size: Optional[float] = None  # points
    bold: bool = False
    color: Optional[str] = None  # see color_xml()


@dataclass
class Para:
    runs: List[Run] = field(default_factory=list)
    align: Optional[str] = None  # "l", "ctr", "r"


@dataclass
class Cell:
    paras: List[Para]
    fill: Optional[str] = None
    align: Optional[str] = None


@dataclass
class Row:
    cells: List[Cell]
    height: int = inch(0.32)


def color_xml(spec: str) -> str:
    """``#RRGGBB`` -> RGB, ``accent1`` -> theme colour, ``accent1/20`` -> 20 % tint of it."""
    if spec.startswith("#"):
        return f'<a:srgbClr val="{spec[1:].upper()}"/>'
    if "/" in spec:
        name, pct = spec.split("/", 1)
        keep = int(pct) * 1000
        return (
            f'<a:schemeClr val="{name}"><a:lumMod val="{keep}"/>'
            f'<a:lumOff val="{100000 - keep}"/></a:schemeClr>'
        )
    return f'<a:schemeClr val="{spec}"/>'


def solid(spec: str) -> str:
    return f"<a:solidFill>{color_xml(spec)}</a:solidFill>"


def run_xml(run: Run) -> str:
    attrs = ' lang="en-US"'
    if run.size:
        attrs += f' sz="{int(round(run.size * 100))}"'
    if run.bold:
        attrs += ' b="1"'
    fill = solid(run.color) if run.color else ""
    return f"<a:r><a:rPr{attrs} dirty=\"0\">{fill}</a:rPr><a:t>{esc(run.text)}</a:t></a:r>"


def para_xml(para: Para) -> str:
    ppr = f'<a:pPr algn="{para.align}"/>' if para.align else ""
    runs = "".join(run_xml(r) for r in para.runs if r.text)
    size = next((r.size for r in para.runs if r.size), None)
    end = f'<a:endParaRPr lang="en-US" sz="{int(size * 100)}" dirty="0"/>' if size else '<a:endParaRPr lang="en-US" dirty="0"/>'
    return f"<a:p>{ppr}{runs}{end}</a:p>"


def text_body(paras: Sequence[Para], tag: str = "p:txBody", body_pr: str = "<a:bodyPr/>") -> str:
    inner = "".join(para_xml(p) for p in paras) or '<a:p><a:endParaRPr lang="en-US"/></a:p>'
    return f"<{tag}>{body_pr}<a:lstStyle/>{inner}</{tag}>"


def paras_of(text: str, **run_kw) -> List[Para]:
    """One paragraph per line of ``text``."""
    return [Para([Run(line, **run_kw)]) for line in text.split("\n")]


def placeholder_sp(sid: int, name: str, ph: Dict[str, str], paras: Sequence[Para]) -> str:
    attrs = "".join(f' {k}="{esc(v)}"' for k, v in ph.items())
    return (
        f'<p:sp><p:nvSpPr><p:cNvPr id="{sid}" name="{esc(name)}"/>'
        f'<p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr><p:nvPr><p:ph{attrs}/></p:nvPr></p:nvSpPr>'
        f"<p:spPr/>{text_body(paras)}</p:sp>"
    )


def textbox_sp(sid: int, name: str, box: Box, paras: Sequence[Para], anchor: str = "t") -> str:
    body_pr = f'<a:bodyPr wrap="square" lIns="0" tIns="0" rIns="0" bIns="0" rtlCol="0" anchor="{anchor}"><a:noAutofit/></a:bodyPr>'
    return (
        f'<p:sp><p:nvSpPr><p:cNvPr id="{sid}" name="{esc(name)}"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr>'
        f'<p:spPr>{box.xfrm()}<a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/></p:spPr>'
        f"{text_body(paras, body_pr=body_pr)}</p:sp>"
    )


_BORDER = '<a:ln{side} w="6350" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:srgbClr val="D9D9D9"/></a:solidFill></a:ln{side}>'
_NO_BORDER = '<a:ln{side} w="0"><a:noFill/></a:ln{side}>'


def _cell_xml(cell: Cell) -> str:
    paras = [Para(p.runs, p.align or cell.align) for p in cell.paras]
    borders = (
        _NO_BORDER.format(side="L") + _NO_BORDER.format(side="R")
        + _BORDER.format(side="T") + _BORDER.format(side="B")
    )
    fill = solid(cell.fill) if cell.fill else "<a:noFill/>"
    return (
        f"<a:tc>{text_body(paras, tag='a:txBody')}"
        f'<a:tcPr marL="68580" marR="68580" marT="27432" marB="27432" anchor="ctr">{borders}{fill}</a:tcPr></a:tc>'
    )


def table_frame(sid: int, name: str, box: Box, col_widths: Sequence[int], rows: Sequence[Row]) -> str:
    grid = "".join(f'<a:gridCol w="{w}"/>' for w in col_widths)
    trs = "".join(f'<a:tr h="{r.height}">' + "".join(_cell_xml(c) for c in r.cells) + "</a:tr>" for r in rows)
    height = sum(r.height for r in rows)
    frame = Box(box.x, box.y, sum(col_widths), height)
    return (
        f'<p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id="{sid}" name="{esc(name)}"/>'
        '<p:cNvGraphicFramePr><a:graphicFrameLocks noGrp="1"/></p:cNvGraphicFramePr><p:nvPr/></p:nvGraphicFramePr>'
        f"{frame.xfrm('p:xfrm')}"
        '<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/table">'
        f'<a:tbl><a:tblPr firstRow="1" bandRow="1"/><a:tblGrid>{grid}</a:tblGrid>{trs}</a:tbl>'
        "</a:graphicData></a:graphic></p:graphicFrame>"
    )


def chart_frame(sid: int, name: str, box: Box, rid: str) -> str:
    return (
        f'<p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id="{sid}" name="{esc(name)}"/>'
        '<p:cNvGraphicFramePr><a:graphicFrameLocks noGrp="1"/></p:cNvGraphicFramePr><p:nvPr/></p:nvGraphicFramePr>'
        f"{box.xfrm('p:xfrm')}"
        f'<a:graphic><a:graphicData uri="{NS_C}">'
        f'<c:chart xmlns:c="{NS_C}" r:id="{rid}"/>'
        "</a:graphicData></a:graphic></p:graphicFrame>"
    )
