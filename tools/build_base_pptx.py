"""Generate the built-in neutral template statusreport/pptx/base.pptx (dev only).

Design D-2: every deck is built by opening a base .pptx and adding slides.
Without a customer template this neutral base is used. Run this script only
when the base design changes, then commit the resulting base.pptx.

    python tools/build_base_pptx.py
"""

from __future__ import annotations

import os
import sys

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, ROOT)

from statusreport.pptx import opc  # noqa: E402
from statusreport.xmlutil import NS_A, PML_NS, doc  # noqa: E402

OUT = os.path.join(ROOT, "statusreport", "pptx", "base.pptx")

SLIDE_W, SLIDE_H = 12192000, 6858000  # 16:9, 13.333" x 7.5"
MARGIN = 609600  # 0.667"
TITLE_Y, TITLE_H = 304800, 868680
BAR_Y = TITLE_Y + TITLE_H + 76200
BODY_Y = BAR_Y + 228600
BODY_H = SLIDE_H - BODY_Y - 457200
CONTENT_W = SLIDE_W - 2 * MARGIN

THEME_NAME = "SW Status Report"
COLORS = {
    "dk2": "1F2A44",
    "lt2": "E7E9EE",
    "accent1": "1F4E79",
    "accent2": "2E86AB",
    "accent3": "7A8B99",
    "accent4": "8E6C8A",
    "accent5": "4F7942",
    "accent6": "B5651D",
    "hlink": "0563C1",
    "folHlink": "954F72",
}
FONT = "Calibri"

GRP = (
    '<p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>'
    '<p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/>'
    '<a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>'
)


def xfrm(x: int, y: int, w: int, h: int) -> str:
    return f'<a:xfrm><a:off x="{x}" y="{y}"/><a:ext cx="{w}" cy="{h}"/></a:xfrm>'


def placeholder(sid: int, name: str, ph: str, box=None, body_pr="<a:bodyPr/>", lst="<a:lstStyle/>", text="") -> str:
    sp_pr = f'<p:spPr>{xfrm(*box)}<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr>' if box else "<p:spPr/>"
    run = f'<a:r><a:rPr lang="en-US"/><a:t>{text}</a:t></a:r>' if text else ""
    return (
        f'<p:sp><p:nvSpPr><p:cNvPr id="{sid}" name="{name}"/>'
        f'<p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr><p:nvPr>{ph}</p:nvPr></p:nvSpPr>'
        f"{sp_pr}<p:txBody>{body_pr}{lst}<a:p>{run}<a:endParaRPr lang=\"en-US\"/></a:p></p:txBody></p:sp>"
    )


def accent_bar(sid: int, x: int, y: int, w: int, h: int) -> str:
    return (
        f'<p:sp><p:nvSpPr><p:cNvPr id="{sid}" name="Accent Bar"/><p:cNvSpPr/><p:nvPr userDrawn="1"/></p:nvSpPr>'
        f'<p:spPr>{xfrm(x, y, w, h)}<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>'
        '<a:solidFill><a:schemeClr val="accent1"/></a:solidFill><a:ln><a:noFill/></a:ln></p:spPr>'
        '<p:txBody><a:bodyPr rtlCol="0" anchor="ctr"/><a:lstStyle/><a:p><a:endParaRPr lang="en-US"/></a:p></p:txBody></p:sp>'
    )


def def_rpr(size: int, color: str, font: str, bold: bool = False) -> str:
    b = ' b="1"' if bold else ""
    return (
        f'<a:defRPr sz="{size}"{b} kern="1200"><a:solidFill><a:schemeClr val="{color}"/></a:solidFill>'
        f'<a:latin typeface="+{font}-lt"/><a:ea typeface="+{font}-ea"/><a:cs typeface="+{font}-cs"/></a:defRPr>'
    )


def theme_xml() -> bytes:
    # order inside clrScheme: dk1 lt1 dk2 lt2 accent1..6 hlink folHlink
    dk2 = f'<a:dk2><a:srgbClr val="{COLORS["dk2"]}"/></a:dk2>'
    lt2 = f'<a:lt2><a:srgbClr val="{COLORS["lt2"]}"/></a:lt2>'
    accents = "".join(f'<a:accent{i}><a:srgbClr val="{COLORS[f"accent{i}"]}"/></a:accent{i}>' for i in range(1, 7))
    clr = (
        '<a:dk1><a:sysClr val="windowText" lastClr="000000"/></a:dk1>'
        '<a:lt1><a:sysClr val="window" lastClr="FFFFFF"/></a:lt1>'
        + dk2 + lt2 + accents
        + f'<a:hlink><a:srgbClr val="{COLORS["hlink"]}"/></a:hlink>'
        + f'<a:folHlink><a:srgbClr val="{COLORS["folHlink"]}"/></a:folHlink>'
    )
    font = f'<a:latin typeface="{FONT}"/><a:ea typeface=""/><a:cs typeface=""/>'
    solid = '<a:solidFill><a:schemeClr val="phClr"/></a:solidFill>'
    tint = '<a:solidFill><a:schemeClr val="phClr"><a:tint val="50000"/></a:schemeClr></a:solidFill>'
    shade = '<a:solidFill><a:schemeClr val="phClr"><a:shade val="80000"/></a:schemeClr></a:solidFill>'
    lines = "".join(
        f'<a:ln w="{w}" cap="flat" cmpd="sng" algn="ctr">{solid}<a:prstDash val="solid"/><a:miter lim="800000"/></a:ln>'
        for w in (6350, 12700, 19050)
    )
    effects = "<a:effectStyle><a:effectLst/></a:effectStyle>" * 3
    return doc(
        f'<a:theme xmlns:a="{NS_A}" name="{THEME_NAME}"><a:themeElements>'
        f'<a:clrScheme name="{THEME_NAME}">{clr}</a:clrScheme>'
        f'<a:fontScheme name="{THEME_NAME}"><a:majorFont>{font}</a:majorFont><a:minorFont>{font}</a:minorFont></a:fontScheme>'
        f'<a:fmtScheme name="{THEME_NAME}">'
        f"<a:fillStyleLst>{solid}{tint}{shade}</a:fillStyleLst>"
        f"<a:lnStyleLst>{lines}</a:lnStyleLst>"
        f"<a:effectStyleLst>{effects}</a:effectStyleLst>"
        f"<a:bgFillStyleLst>{solid}{tint}{shade}</a:bgFillStyleLst>"
        "</a:fmtScheme></a:themeElements><a:objectDefaults/><a:extraClrSchemeLst/></a:theme>"
    )


def master_xml(layout_rids) -> bytes:
    title = placeholder(
        2, "Title Placeholder 1", '<p:ph type="title"/>',
        (MARGIN, TITLE_Y, CONTENT_W, TITLE_H),
        '<a:bodyPr vert="horz" lIns="0" tIns="0" rIns="0" bIns="0" anchor="b"><a:normAutofit/></a:bodyPr>',
        text="Click to edit Master title style",
    )
    body = placeholder(
        3, "Text Placeholder 2", '<p:ph type="body" idx="1"/>',
        (MARGIN, BODY_Y, CONTENT_W, BODY_H),
        '<a:bodyPr vert="horz" lIns="0" tIns="0" rIns="0" bIns="0"><a:normAutofit/></a:bodyPr>',
        text="Click to edit Master text styles",
    )
    bar = accent_bar(4, MARGIN, BAR_Y, 1371600, 45720)
    layouts = "".join(
        f'<p:sldLayoutId id="{2147483649 + i}" r:id="{rid}"/>' for i, rid in enumerate(layout_rids)
    )
    title_style = (
        '<p:titleStyle><a:lvl1pPr algn="l" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1">'
        '<a:lnSpc><a:spcPct val="90000"/></a:lnSpc><a:spcBef><a:spcPct val="0"/></a:spcBef><a:buNone/>'
        + def_rpr(3000, "tx2", "mj", bold=True)
        + "</a:lvl1pPr></p:titleStyle>"
    )
    body_style = (
        '<p:bodyStyle><a:lvl1pPr marL="228600" indent="-228600" algn="l" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1">'
        '<a:spcBef><a:spcPts val="1000"/></a:spcBef><a:buFont typeface="Arial"/><a:buChar char="&#8226;"/>'
        + def_rpr(2000, "tx1", "mn")
        + "</a:lvl1pPr></p:bodyStyle>"
    )
    other_style = (
        '<p:otherStyle><a:defPPr><a:defRPr lang="en-US"/></a:defPPr>'
        '<a:lvl1pPr marL="0" algn="l" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1">'
        + def_rpr(1800, "tx1", "mn")
        + "</a:lvl1pPr></p:otherStyle>"
    )
    return doc(
        f"<p:sldMaster {PML_NS}><p:cSld>"
        '<p:bg><p:bgRef idx="1001"><a:schemeClr val="bg1"/></p:bgRef></p:bg>'
        f"<p:spTree>{GRP}{title}{body}{bar}</p:spTree></p:cSld>"
        '<p:clrMap bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2" accent1="accent1" accent2="accent2" '
        'accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6" hlink="hlink" folHlink="folHlink"/>'
        f"<p:sldLayoutIdLst>{layouts}</p:sldLayoutIdLst>"
        f"<p:txStyles>{title_style}{body_style}{other_style}</p:txStyles></p:sldMaster>"
    )


def layout_xml(name: str, kind: str, shapes: str, show_master: bool = True) -> bytes:
    show = "" if show_master else ' showMasterSp="0"'
    return doc(
        f'<p:sldLayout {PML_NS} type="{kind}" preserve="1"{show}>'
        f'<p:cSld name="{name}"><p:spTree>{GRP}{shapes}</p:spTree></p:cSld>'
        "<p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sldLayout>"
    )


def title_slide_layout() -> bytes:
    title_y = 2286000
    title = placeholder(
        2, "Title 1", '<p:ph type="ctrTitle"/>', (MARGIN, title_y, CONTENT_W, 1219200),
        '<a:bodyPr lIns="0" tIns="0" rIns="0" bIns="0" anchor="b"><a:normAutofit/></a:bodyPr>',
        '<a:lstStyle><a:lvl1pPr algn="l"><a:defRPr sz="4000"/></a:lvl1pPr></a:lstStyle>',
    )
    bar = accent_bar(3, MARGIN, title_y + 1219200 + 152400, 1371600, 45720)
    sub = placeholder(
        4, "Subtitle 2", '<p:ph type="subTitle" idx="1"/>', (MARGIN, title_y + 1219200 + 381000, CONTENT_W, 914400),
        '<a:bodyPr lIns="0" tIns="0" rIns="0" bIns="0"><a:normAutofit/></a:bodyPr>',
        '<a:lstStyle><a:lvl1pPr marL="0" indent="0" algn="l"><a:buNone/>'
        '<a:defRPr sz="2000"><a:solidFill><a:schemeClr val="tx2"/></a:solidFill></a:defRPr>'
        "</a:lvl1pPr></a:lstStyle>",
    )
    return layout_xml("Title Slide", "title", title + bar + sub, show_master=False)


def title_only_layout() -> bytes:
    title = placeholder(2, "Title 1", '<p:ph type="title"/>')
    return layout_xml("Title Only", "titleOnly", title)


def title_content_layout() -> bytes:
    title = placeholder(2, "Title 1", '<p:ph type="title"/>')
    body = placeholder(3, "Content Placeholder 2", '<p:ph idx="1"/>')
    return layout_xml("Title and Content", "obj", title + body)


def presentation_xml() -> bytes:
    return doc(
        f'<p:presentation {PML_NS} saveSubsetFonts="1">'
        '<p:sldMasterIdLst><p:sldMasterId id="2147483648" r:id="rId1"/></p:sldMasterIdLst>'
        f'<p:sldSz cx="{SLIDE_W}" cy="{SLIDE_H}"/><p:notesSz cx="6858000" cy="9144000"/>'
        '<p:defaultTextStyle><a:defPPr><a:defRPr lang="en-US"/></a:defPPr>'
        '<a:lvl1pPr marL="0" algn="l" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1">'
        + def_rpr(1800, "tx1", "mn")
        + "</a:lvl1pPr></p:defaultTextStyle></p:presentation>"
    )


def build() -> opc.Package:
    pkg = opc.Package()
    pres = "ppt/presentation.xml"
    master = "ppt/slideMasters/slideMaster1.xml"
    theme = "ppt/theme/theme1.xml"

    pkg.add_part(
        "docProps/core.xml",
        doc(
            '<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" '
            'xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" '
            'xmlns:dcmitype="http://purl.org/dc/dcmitype/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">'
            "<dc:title>SW Status Report</dc:title><dc:creator>sw-status-report</dc:creator></cp:coreProperties>"
        ),
        opc.CT_CORE_PROPS,
    )
    pkg.add_part(
        "docProps/app.xml",
        doc(
            '<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" '
            'xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">'
            "<Application>sw-status-report</Application><Slides>0</Slides></Properties>"
        ),
        opc.CT_EXTENDED_PROPS,
    )
    pkg.add_rel("", opc.RT_OFFICE_DOCUMENT, pres)
    pkg.add_rel("", opc.RT_CORE_PROPS, "docProps/core.xml")
    pkg.add_rel("", opc.RT_EXTENDED_PROPS, "docProps/app.xml")

    pkg.add_part(pres, presentation_xml(), opc.CT_PRESENTATION)
    pkg.add_rel(pres, opc.RT_SLIDE_MASTER, master)  # rId1, referenced in presentation.xml
    for name, rt, ct, body in (
        ("ppt/presProps.xml", opc.RT_PRES_PROPS, opc.CT_PRES_PROPS, f"<p:presentationPr {PML_NS}/>"),
        (
            "ppt/viewProps.xml", opc.RT_VIEW_PROPS, opc.CT_VIEW_PROPS,
            f'<p:viewPr {PML_NS}><p:normalViewPr><p:restoredLeft sz="15620"/><p:restoredTop sz="94660"/>'
            '</p:normalViewPr><p:gridSpacing cx="76200" cy="76200"/></p:viewPr>',
        ),
        (
            "ppt/tableStyles.xml", opc.RT_TABLE_STYLES, opc.CT_TABLE_STYLES,
            f'<a:tblStyleLst xmlns:a="{NS_A}" def="{{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}}"/>',
        ),
    ):
        pkg.add_part(name, doc(body), ct)
        pkg.add_rel(pres, rt, name)
    pkg.add_part(theme, theme_xml(), opc.CT_THEME)
    pkg.add_rel(pres, opc.RT_THEME, theme)

    layouts = [title_slide_layout(), title_only_layout(), title_content_layout()]
    rids = []
    for i, data in enumerate(layouts, 1):
        name = f"ppt/slideLayouts/slideLayout{i}.xml"
        pkg.add_part(name, data, opc.CT_SLIDE_LAYOUT)
        pkg.add_rel(name, opc.RT_SLIDE_MASTER, master)
        rids.append(pkg.add_rel(master, opc.RT_SLIDE_LAYOUT, name))
    pkg.add_rel(master, opc.RT_THEME, theme)
    pkg.add_part(master, master_xml(rids), opc.CT_SLIDE_MASTER)
    return pkg


def main() -> None:
    from statusreport.pptx.lint import lint_package

    pkg = build()
    problems = lint_package(pkg)
    if problems:
        sys.exit("base.pptx has problems:\n  " + "\n  ".join(problems))
    pkg.save(OUT)
    print(f"written {OUT}")


if __name__ == "__main__":
    main()
