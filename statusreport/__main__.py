"""Command line (R-OP-1). Deck generation is added in M5; M2 provides --init and --check.

    python -m statusreport --init Tasks.xlsx
    python -m statusreport Tasks.xlsx --check [--polarion export.xlsx] [--date 2026-11-10]
"""

from __future__ import annotations

import argparse
import datetime as dt
import os
import sys
from collections import Counter
from typing import List, Optional

from . import __version__, parse
from .loader import InputError, load
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
    ap.add_argument("--check", action="store_true", help="only read and validate the input, list all problems")
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
    week = s.report_date.isocalendar()
    print(f"Project:      {s.project or '-'} / {s.customer or '-'}")
    print(f"Report date:  {s.report_date.isoformat()} (CW {week[1]:02d}/{week[0]})")
    print(f"Source:       Tasks sheet" + (f" + {result.import_source}" if result.import_source else ""))
    per_version = Counter(t.version for t in result.tasks)
    print(f"Tasks:        {len(result.tasks)} in {len(per_version)} version(s)")
    for version in sorted(per_version, key=parse.natural_key):
        print(f"  {version:<16} {per_version[version]:>4}")
    if result.issues:
        print(f"\n{len(result.errors)} error(s), {len(result.issues) - len(result.errors)} warning(s):")
        for issue in result.issues:
            print("  " + str(issue))
    if not args.check:
        print("\nDeck generation follows in milestone M5; nothing written.")
    return EXIT_ROW_ERRORS if result.errors else EXIT_OK


if __name__ == "__main__":
    sys.exit(main())
