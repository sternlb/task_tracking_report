"""Command line (R-OP-1).

    python -m statusreport --init Tasks.xlsx
    python -m statusreport Tasks.xlsx [--polarion export.xlsx] [--date 2026-11-10]
                                      [--template customer.pptx] [--out deck.pptx] [--no-save] [--force]
    python -m statusreport Tasks.xlsx --check

A normal run writes the deck next to the workbook and stores this week's
snapshot in history/ (R-HI-1). --no-save writes only the deck, --check nothing.
"""

from __future__ import annotations

import argparse
import os
import sys
import zipfile
from collections import Counter
from typing import List, Optional

from . import __version__, history, output, parse
from .calc import compute
from .loader import InputError, load
from .model import VersionStatus
from .pptx.deck import Deck
from .render import render
from .report import build_report
from .template import template_bytes
from .xlsx_read import XlsxError

EXIT_OK, EXIT_ROW_ERRORS, EXIT_FAILED = 0, 1, 2


def _parse_args(argv: Optional[List[str]]):
    ap = argparse.ArgumentParser(
        prog="statusreport",
        description="Weekly software status deck (.pptx) from a task list (.xlsx).",
    )
    ap.add_argument("input", help="input workbook (.xlsx)")
    ap.add_argument("--init", action="store_true", help="create an empty input workbook at INPUT and stop")
    ap.add_argument("--polarion", metavar="FILE", help="separate Polarion Excel export instead of the Import sheet")
    ap.add_argument("--date", metavar="YYYY-MM-DD", help="report date (overrides the Settings sheet)")
    ap.add_argument("--check", action="store_true", help="only read and validate the input, write nothing")
    ap.add_argument("--no-save", action="store_true", help="do not store this week's snapshot in history/")
    ap.add_argument("--template", metavar="PPTX", help="customer PowerPoint template (overrides the Settings sheet)")
    ap.add_argument("--out", metavar="PPTX", help="output file (default: '<Project> Status <week>.pptx' next to the input)")
    ap.add_argument("--force", action="store_true", help="overwrite existing output files")
    ap.add_argument("--version", action="version", version=f"%(prog)s {__version__}")
    return ap.parse_args(argv)


def _init(path: str, force: bool) -> int:
    if not path.lower().endswith(".xlsx"):
        path += ".xlsx"
    if os.path.exists(path) and not force:
        print(f"{path} already exists - use --force to replace it")
        return EXIT_FAILED
    with open(path, "wb") as fh:
        fh.write(template_bytes())
    print(f"created {path}")
    return EXIT_OK


def _pct(value: Optional[float]) -> str:
    return "  -  " if value is None else f"{value:3.0f} %"


def _prev_pct(prev: Optional[history.Snapshot], found, value: Optional[float]) -> str:
    """Bracket text as on the slide (R-OUT-4): (50 %), (new), or (-) without any history."""
    if prev is None:
        text = "(-)"
    elif found is None:
        text = "(new)"
    else:
        text = "(-)" if value is None else f"({value:.0f} %)"
    return text.ljust(7)


def _print_versions(versions: List[VersionStatus], prev: Optional[history.Snapshot]) -> None:
    for vs in versions:
        note = "weighted by estimate" if vs.weighting == "estimate" else "plain average (estimate missing)"
        if vs.uncounted:
            note += f", {vs.uncounted} task(s) without progress"
        pv = prev.versions.get(vs.name) if prev else None
        due = vs.due.isoformat() if vs.due else "-"
        print(
            f"\nVersion {vs.name}: {_pct(vs.progress)} {_prev_pct(prev, pv, pv.progress if pv else None)}"
            f" expected {_pct(vs.expected)}  {vs.light.upper():<6}  due {due}  ({note})"
        )
        for t in vs.tasks:
            pt = prev.tasks.get(t.id) if prev else None
            light = t.light + (f" (was {pt.light})" if pt and pt.light != t.light else "")
            override = " (override)" if t.light_override else ""
            start = (t.start.isoformat() + ("*" if t.start_derived else " ")) if t.start else "     -     "
            due = t.due.isoformat() if t.due else "    -     "
            print(
                f"  {t.id:<14} {_pct(t.progress)} {_prev_pct(prev, pt, pt.progress if pt else None)}"
                f" expected {_pct(t.expected)}  {light:<20}{override}  {start} .. {due}  {t.title[:40]}"
            )
    if any(t.start_derived for vs in versions for t in vs.tasks):
        print("\n* Start was empty and is derived from Due minus the estimate (working days).")


def main(argv: Optional[List[str]] = None) -> int:
    for stream in (sys.stdout, sys.stderr):
        try:
            stream.reconfigure(errors="replace")  # type: ignore[attr-defined]
        except AttributeError:
            pass
    args = _parse_args(argv)
    if args.init:
        return _init(args.input, args.force)

    try:
        report_date = parse.date(args.date) if args.date else None
    except ValueError as exc:
        print(f"--date: {exc}")
        return EXIT_FAILED
    try:
        result = load(args.input, polarion=args.polarion)
    except (InputError, XlsxError) as exc:
        print(f"error: {exc}")
        return EXIT_FAILED
    if report_date:
        result.settings.report_date = report_date

    s = result.settings
    versions, calc_issues = compute(result.tasks, s)
    hist_issues: list = []
    folder = history.history_dir(args.input)
    snapshots = history.load_all(folder, hist_issues)
    current = history.snapshot_from(versions, s.report_date)
    prev = history.previous(snapshots, current.week)
    keys, axis, series = history.trend(snapshots, current, s.trend_weeks, [v.name for v in versions])

    week = s.report_date.isocalendar()
    print(f"Project:      {s.project or '-'} / {s.customer or '-'}")
    print(f"Report date:  {s.report_date.isoformat()} (CW {week[1]:02d}/{week[0]})")
    print(f"Source:       Tasks sheet" + (f" + {result.import_source}" if result.import_source else ""))
    per_version = Counter(t.version for t in result.tasks)
    print(f"Tasks:        {len(result.tasks)} in {len(per_version)} version(s)")
    print(f"Thresholds:   orange > {s.orange_pp:g} pp, red > {s.red_pp:g} pp behind")
    prev_text = f"previous week {prev.week} ({prev.report_date.isoformat()})" if prev else "no previous week yet"
    print(f"History:      {len(snapshots)} snapshot(s), {prev_text}; trend {axis[0]} .. {axis[-1]} ({len(keys)} week(s))")

    _print_versions(versions, prev)

    issues = sorted(result.issues + calc_issues + hist_issues, key=lambda i: (i.sheet, i.row))
    if issues:
        print(f"\n{len(result.errors)} error(s), {len(issues) - len(result.errors)} warning(s):")
        for issue in issues:
            print("  " + str(issue))

    if args.check:
        print("\n--check: nothing written.")
        return EXIT_ROW_ERRORS if result.errors else EXIT_OK
    if not versions:
        print("\nerror: no tasks to report - nothing written.")
        return EXIT_FAILED

    # 1. the deck (first: if it cannot be written, the week is not stored either)
    template = output.template_path(args.input, s.template, args.template)
    try:
        deck = Deck.from_file(template) if template else Deck.default()
    except (OSError, ValueError, KeyError, zipfile.BadZipFile) as exc:
        print(f"\nerror: cannot use the template {template}: {exc}")
        return EXIT_FAILED
    report = build_report(s, versions, prev, current.week, axis, series)
    render(report, deck, s.title_layout, s.content_layout, s.rows_per_slide, s.footer_text)
    target = args.out or os.path.join(os.path.dirname(os.path.abspath(args.input)), output.default_name(s.project, current.week))
    target = output.free_path(target, args.force)
    try:
        deck.save(target)
    except PermissionError:
        print(f"\nerror: cannot write {target} - is it open in PowerPoint?")
        return EXIT_FAILED
    for warning in deck.warnings:
        print(f"warning template: {warning}")
    print(f"\nDeck written: {target} ({len(deck.slides)} slides)")

    # 2. the snapshot
    if args.no_save:
        print("--no-save: snapshot not stored.")
    else:
        replaced = current.week in snapshots
        path = history.save(folder, current)
        print(f"Snapshot {'replaced' if replaced else 'saved'}: {path}")
    return EXIT_ROW_ERRORS if result.errors else EXIT_OK


if __name__ == "__main__":
    sys.exit(main())
