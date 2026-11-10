"""Native PowerPoint line chart with cached values and an embedded workbook (D-5)."""

from __future__ import annotations

from typing import List, Optional, Sequence, Tuple

from ..xlsx_write import col_letter, workbook_bytes
from ..xmlutil import NS_A, NS_C, NS_R, doc, esc
from .shapes import color_xml

Series = Tuple[str, Sequence[Optional[float]]]

SHEET = "Sheet1"


def _str_cache(values: Sequence[str]) -> str:
    pts = "".join(f'<c:pt idx="{i}"><c:v>{esc(v)}</c:v></c:pt>' for i, v in enumerate(values))
    return f'<c:strCache><c:ptCount val="{len(values)}"/>{pts}</c:strCache>'


def _num_cache(values: Sequence[Optional[float]]) -> str:
    pts = "".join(
        f'<c:pt idx="{i}"><c:v>{v:g}</c:v></c:pt>' for i, v in enumerate(values) if v is not None
    )
    return f'<c:numCache><c:formatCode>0</c:formatCode><c:ptCount val="{len(values)}"/>{pts}</c:numCache>'


def _text_props(size_pt: float, color: str = "tx1") -> str:
    return (
        '<c:txPr><a:bodyPr/><a:lstStyle/><a:p><a:pPr>'
        f'<a:defRPr sz="{int(size_pt * 100)}"><a:solidFill>{color_xml(color)}</a:solidFill></a:defRPr>'
        '</a:pPr><a:endParaRPr lang="en-US"/></a:p></c:txPr>'
    )


def line_chart_xml(
    categories: Sequence[str],
    series: Sequence[Series],
    workbook_rid: str,
    y_max: Optional[float] = 100,
    y_step: Optional[float] = 20,
    font_pt: float = 12,
    value_format: str = '0" %"',
    styles: Optional[Sequence[dict]] = None,
) -> bytes:
    """Line chart. ``y_max``/``y_step`` None = PowerPoint chooses. ``styles`` per series:
    ``{"color": "accent1" | "#RRGGBB", "dash": bool, "marker": bool}``."""
    n = len(categories)
    cat_ref = f"{SHEET}!$A$2:$A${n + 1}"
    sers = []
    for i, (name, values) in enumerate(series):
        col = col_letter(i + 1)
        style = (styles[i] if styles and i < len(styles) else None) or {}
        color = color_xml(style.get("color", f"accent{i % 6 + 1}"))
        dash = '<a:prstDash val="dash"/>' if style.get("dash") else ""
        if style.get("marker", True):
            marker = (
                '<c:marker><c:symbol val="circle"/><c:size val="7"/>'
                f'<c:spPr><a:solidFill>{color}</a:solidFill><a:ln w="9525"><a:solidFill>{color}</a:solidFill></a:ln></c:spPr></c:marker>'
            )
        else:
            marker = '<c:marker><c:symbol val="none"/></c:marker>'
        sers.append(
            f'<c:ser><c:idx val="{i}"/><c:order val="{i}"/>'
            f"<c:tx><c:strRef><c:f>{SHEET}!${col}$1</c:f>{_str_cache([name])}</c:strRef></c:tx>"
            f'<c:spPr><a:ln w="28575" cap="rnd"><a:solidFill>{color}</a:solidFill>{dash}<a:round/></a:ln></c:spPr>'
            f"{marker}"
            f"<c:cat><c:strRef><c:f>{cat_ref}</c:f>{_str_cache(categories)}</c:strRef></c:cat>"
            f"<c:val><c:numRef><c:f>{SHEET}!${col}$2:${col}${n + 1}</c:f>{_num_cache(values)}</c:numRef></c:val>"
            '<c:smooth val="0"/></c:ser>'
        )
    y_max_xml = f'<c:max val="{y_max:g}"/>' if y_max is not None else ""
    y_step_xml = f'<c:majorUnit val="{y_step:g}"/>' if y_step is not None else ""
    grid = '<c:spPr><a:ln w="9525"><a:solidFill><a:srgbClr val="D9D9D9"/></a:solidFill></a:ln></c:spPr>'
    axis_line = '<c:spPr><a:ln w="9525"><a:solidFill><a:srgbClr val="BFBFBF"/></a:solidFill></a:ln></c:spPr>'
    body = (
        f'<c:chartSpace xmlns:c="{NS_C}" xmlns:a="{NS_A}" xmlns:r="{NS_R}">'
        '<c:date1904 val="0"/><c:lang val="en-US"/><c:roundedCorners val="0"/>'
        '<c:chart><c:autoTitleDeleted val="1"/><c:plotArea><c:layout/>'
        '<c:lineChart><c:grouping val="standard"/><c:varyColors val="0"/>'
        + "".join(sers)
        + '<c:marker val="1"/><c:axId val="500000001"/><c:axId val="500000002"/></c:lineChart>'
        '<c:catAx><c:axId val="500000001"/><c:scaling><c:orientation val="minMax"/></c:scaling>'
        '<c:delete val="0"/><c:axPos val="b"/><c:numFmt formatCode="General" sourceLinked="1"/>'
        '<c:majorTickMark val="out"/><c:minorTickMark val="none"/><c:tickLblPos val="nextTo"/>'
        f'{axis_line}<c:crossAx val="500000002"/><c:crosses val="autoZero"/><c:auto val="1"/>'
        '<c:lblAlgn val="ctr"/><c:lblOffset val="100"/><c:noMultiLvlLbl val="0"/></c:catAx>'
        '<c:valAx><c:axId val="500000002"/><c:scaling><c:orientation val="minMax"/>'
        f'{y_max_xml}<c:min val="0"/></c:scaling><c:delete val="0"/><c:axPos val="l"/>'
        f"<c:majorGridlines>{grid}</c:majorGridlines>"
        f'<c:numFmt formatCode="{esc(value_format)}" sourceLinked="0"/>'
        '<c:majorTickMark val="none"/><c:minorTickMark val="none"/><c:tickLblPos val="nextTo"/>'
        '<c:spPr><a:ln><a:noFill/></a:ln></c:spPr>'
        '<c:crossAx val="500000001"/><c:crosses val="autoZero"/><c:crossBetween val="between"/>'
        f"{y_step_xml}</c:valAx>"
        '<c:spPr><a:noFill/><a:ln><a:noFill/></a:ln></c:spPr>'
        "</c:plotArea>"
        '<c:legend><c:legendPos val="b"/><c:overlay val="0"/></c:legend>'
        '<c:plotVisOnly val="1"/><c:dispBlanksAs val="gap"/></c:chart>'
        '<c:spPr><a:noFill/><a:ln><a:noFill/></a:ln></c:spPr>'
        f"{_text_props(font_pt)}"
        f'<c:externalData r:id="{workbook_rid}"><c:autoUpdate val="0"/></c:externalData>'
        "</c:chartSpace>"
    )
    return doc(body)


def chart_workbook(categories: Sequence[str], series: Sequence[Series]) -> bytes:
    """The data behind the chart, so that "Edit Data" in PowerPoint works."""
    rows: List[list] = [[""] + [name for name, _ in series]]
    for i, cat in enumerate(categories):
        rows.append([cat] + [values[i] for _, values in series])
    return workbook_bytes([(SHEET, rows)])
