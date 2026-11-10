"""Structural checks for an Office package (design 8.1 step 6, tools/opc_lint.py).

These catch the mistakes that make PowerPoint show its "repair" prompt:
broken XML, missing content types, dangling relationships, duplicate IDs.
They do not validate against the full XML schema.
"""

from __future__ import annotations

import re
from typing import List
from xml.etree import ElementTree as ET

from . import opc
from ..xmlutil import NS_P, NS_R

_ID_ATTR = re.compile(r"\{%s\}(id|embed|link|pict|dm|lo|qs|cs)$" % re.escape(NS_R))

_EXPECTED_CT = {
    opc.RT_SLIDE: opc.CT_SLIDE,
    opc.RT_SLIDE_LAYOUT: opc.CT_SLIDE_LAYOUT,
    opc.RT_SLIDE_MASTER: opc.CT_SLIDE_MASTER,
    opc.RT_CHART: opc.CT_CHART,
    opc.RT_THEME: opc.CT_THEME,
}


def lint_package(pkg: opc.Package) -> List[str]:
    problems: List[str] = []
    trees = {}

    for name, data in pkg.parts.items():
        if name.endswith((".xml", ".rels")):
            try:
                trees[name] = ET.fromstring(data)
            except ET.ParseError as exc:
                problems.append(f"{name}: XML not well-formed ({exc})")

    for name in pkg.parts:
        if not name.endswith(".rels") and not pkg.content_type(name):
            problems.append(f"{name}: no content type")

    rel_ids = {}
    for name in pkg.parts:
        if not name.endswith(".rels") or name not in trees:
            continue
        source = opc.source_of_rels(name)
        if source and pkg.find(source) is None:
            problems.append(f"{name}: relationships of a part that does not exist")
            continue
        seen = set()
        for r in pkg.rels(source):
            if r.rid in seen:
                problems.append(f"{name}: duplicate Id {r.rid}")
            seen.add(r.rid)
            if r.external:
                continue
            target = opc.resolve(source, r.target)
            if pkg.find(target) is None:
                problems.append(f"{name}: {r.rid} points to missing part {target}")
                continue
            expected = _EXPECTED_CT.get(r.type)
            if expected and pkg.content_type(target) != expected:
                problems.append(f"{target}: content type {pkg.content_type(target)} does not fit {r.type}")
        rel_ids[source] = seen

    # every r:id / r:embed used inside an XML part must exist in its .rels
    for name, root in trees.items():
        if name.endswith(".rels") or name == opc.CONTENT_TYPES:
            continue
        known = rel_ids.get(name, set())
        for el in root.iter():
            for attr, value in el.attrib.items():
                if _ID_ATTR.match(attr) and value and value not in known:
                    problems.append(f"{name}: reference {value} has no relationship")

    try:
        main = pkg.main_part()
    except ValueError as exc:
        return problems + [str(exc)]
    if main in trees:
        problems += _lint_presentation(pkg, main, trees)
    return problems


def _lint_presentation(pkg: opc.Package, main: str, trees) -> List[str]:
    problems: List[str] = []
    p = "{%s}" % NS_P
    rid_attr = "{%s}id" % NS_R
    root = trees[main]
    rels = {r.rid: r for r in pkg.rels(main)}

    slide_ids = set()
    for el in root.iter(p + "sldId"):
        sid = int(el.get("id", "0"))
        if not 256 <= sid < 2147483648 or sid in slide_ids:
            problems.append(f"{main}: invalid or duplicate slide id {sid}")
        slide_ids.add(sid)
        rel = rels.get(el.get(rid_attr, ""))
        if rel is None or rel.type != opc.RT_SLIDE:
            problems.append(f"{main}: sldId {sid} does not point to a slide")

    master_ids = set()
    for el in root.iter(p + "sldMasterId"):
        master_ids.add(int(el.get("id", "0")))
        rel = rels.get(el.get(rid_attr, ""))
        if rel is None:
            continue
        master = opc.resolve(main, rel.target)
        mroot = trees.get(pkg.find(master) or "")
        if mroot is None:
            continue
        for lay in mroot.iter(p + "sldLayoutId"):
            lid = int(lay.get("id", "0"))
            if lid < 2147483648 or lid in master_ids:
                problems.append(f"{master}: invalid or duplicate layout id {lid}")
            master_ids.add(lid)

    for name in pkg.parts:
        if pkg.content_type(name) != opc.CT_SLIDE or name not in trees:
            continue
        layouts = pkg.related(name, opc.RT_SLIDE_LAYOUT)
        if len(layouts) != 1:
            problems.append(f"{name}: needs exactly one slide layout, has {len(layouts)}")
        ids = [el.get("id") for el in trees[name].iter(p + "cNvPr")]
        if len(ids) != len(set(ids)):
            problems.append(f"{name}: duplicate shape ids")
    return problems
