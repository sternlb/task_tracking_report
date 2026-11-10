"""Build the delivery (T-7.2, decision D-6).

    python tools/build_pyz.py

Creates:
    dist/sw-status-report/statusreport.pyz   the whole tool in one file (stdlib zipapp)
    dist/sw-status-report/make_report.cmd    drag the input workbook onto it (Windows)
    dist/sw-status-report/Tasks.xlsx         an empty input workbook (same as --init)
    dist/sw-status-report/README.md          the user guide
    dist/sw-status-report-<version>.zip      the folder above as one zip to hand over

Runs with the standard library only; nothing is downloaded or installed.
"""

from __future__ import annotations

import os
import shutil
import sys
import tempfile
import zipapp
import zipfile

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, ROOT)

from statusreport import __version__  # noqa: E402
from statusreport.template import template_bytes  # noqa: E402

DIST = os.path.join(ROOT, "dist")
NAME = "sw-status-report"
ENTRY = "from statusreport.__main__ import main\nimport sys\nsys.exit(main())\n"


def build_pyz(target: str) -> None:
    with tempfile.TemporaryDirectory() as stage:
        shutil.copytree(
            os.path.join(ROOT, "statusreport"),
            os.path.join(stage, "statusreport"),
            ignore=shutil.ignore_patterns("__pycache__", "*.pyc"),
        )
        with open(os.path.join(stage, "__main__.py"), "w", encoding="utf-8") as fh:
            fh.write(ENTRY)
        zipapp.create_archive(stage, target, interpreter="/usr/bin/env python3", compressed=True)


def main(dist: str = DIST) -> str:
    """Build into ``dist``; returns the path of the zip."""
    folder = os.path.join(dist, NAME)
    if os.path.isdir(folder):
        shutil.rmtree(folder)
    os.makedirs(folder)
    build_pyz(os.path.join(folder, "statusreport.pyz"))
    shutil.copy(os.path.join(ROOT, "make_report.cmd"), folder)
    shutil.copy(os.path.join(ROOT, "README.md"), folder)
    with open(os.path.join(folder, "Tasks.xlsx"), "wb") as fh:
        fh.write(template_bytes())

    archive = os.path.join(dist, f"{NAME}-{__version__}.zip")
    with zipfile.ZipFile(archive, "w", zipfile.ZIP_DEFLATED) as zf:
        for name in sorted(os.listdir(folder)):
            zf.write(os.path.join(folder, name), f"{NAME}/{name}")
    for name in sorted(os.listdir(folder)):
        print(f"  {name:<20} {os.path.getsize(os.path.join(folder, name)):>8} bytes")
    print(f"built {archive}")
    return archive


if __name__ == "__main__":
    main()
