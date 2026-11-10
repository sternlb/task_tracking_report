"""Demo input with six weeks of history (dev tool, T-5.4).

Writes ``<folder>/Tasks.xlsx`` and ``<folder>/history/2026-W40.json`` … ``W45``. The
history is not invented: each past week is calculated with the real rules, with
proportionally fewer hours booked, so lights and trends behave realistically.
Then run the tool for CW 46:

    python tools/make_demo.py out/demo
    python -m statusreport out/demo/Tasks.xlsx
"""

from __future__ import annotations

import copy
import datetime as dt
import os
import shutil
import sys

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, ROOT)

from statusreport import history, labels  # noqa: E402
from statusreport.calc import compute  # noqa: E402
from statusreport.model import Settings, Task  # noqa: E402
from statusreport.xlsx_write import Sheet, Styled, workbook_bytes  # noqa: E402

D = dt.date
REPORT_DATE = D(2026, 11, 10)
NEW_THIS_WEEK = {"PRJ-215"}

# id, title, version, owner, start, due, estimate, spent, progress %, state, override, comment
TASKS = [
    ("PRJ-101", "Boot sequence and early init", "1.1.0", "Anna", D(2026, 8, 3), D(2026, 9, 25), 120, 128, None, "Done", None, ""),
    ("PRJ-102", "Persistent configuration storage", "1.1.0", "Ben", D(2026, 8, 10), D(2026, 10, 9), 80, 76, None, "Done", None, ""),
    ("PRJ-103", "Release notes and SW package 1.1.0", "1.1.0", "Chris", D(2026, 10, 12), D(2026, 10, 30), 24, 22, None, "Closed", None, "delivered 30 Oct"),
    ("PRJ-201", "Login dialog with single sign-on", "1.2.0", "Anna", D(2026, 10, 19), D(2026, 11, 27), 80, 44, None, "In Progress", None, ""),
    ("PRJ-202", "CAN bus driver for the new ECU generation", "1.2.0", "Ben", D(2026, 10, 1), D(2026, 11, 13), 120, 92, None, "In Progress", None, "HW sample arrived late"),
    ("PRJ-203", "Diagnostic trouble codes (DTC) handling", "1.2.0", "Chris", D(2026, 9, 28), D(2026, 11, 6), 64, 60, None, "In Review", None, "review open"),
    ("PRJ-204", "Bootloader update over UDS", "1.2.0", "Dana", D(2026, 10, 26), D(2026, 12, 18), 96, 30, None, "In Progress", None, ""),
    ("PRJ-205", "Logging and trace export", "1.2.0", "Eli", None, D(2026, 12, 11), "2d", None, None, "Open", None, ""),
    ("PRJ-206", "Unit tests for the CRC module", "1.2.0", "Fay", D(2026, 10, 5), D(2026, 11, 20), 40, 40, 100, "Done", None, ""),
    ("PRJ-207", "Watchdog handling with configurable timeouts", "1.2.0", "Gus", D(2026, 10, 12), D(2026, 11, 27), 48, 20, None, "In Progress", None, ""),
    ("PRJ-208", "Startup self test of RAM, ROM and peripherals according to the safety concept", "1.2.0", "Hana", D(2026, 10, 19), D(2026, 12, 4), 72, 18, None, "In Progress", None, "scope under discussion with the safety team"),
    ("PRJ-209", "Power management: sleep and wake-up", "1.2.0", "Ivo", D(2026, 10, 5), D(2026, 12, 11), 100, 64, None, "In Progress", None, ""),
    ("PRJ-210", "Error memory and freeze frames", "1.2.0", "Jana", D(2026, 10, 26), D(2026, 12, 18), 60, 8, None, "In Progress", "Orange", "waits for spec v2.1"),
    ("PRJ-211", "Coding and variant handling", "1.2.0", "Kai", D(2026, 11, 2), D(2026, 12, 11), 40, 6, None, "In Progress", None, ""),
    ("PRJ-212", "Security access (seed and key)", "1.2.0", "Lea", D(2026, 10, 12), D(2026, 11, 20), 56, 50, None, "In Progress", None, ""),
    ("PRJ-213", "Flash driver integration", "1.2.0", "Max", D(2026, 10, 5), D(2026, 11, 13), 48, 46, None, "Verified", None, ""),
    ("PRJ-214", "Integration test on HIL", "1.2.0", "Nina", D(2026, 11, 16), D(2026, 12, 18), 80, 0, None, "Open", None, ""),
    ("PRJ-215", "Hotfix: wrong unit in speed signal", "1.2.0", "Otto", D(2026, 11, 9), D(2026, 11, 13), 8, 2, None, "In Progress", None, "new this week"),
    ("PRJ-301", "Architecture concept 2.0", "2.0.0", "Anna", D(2026, 10, 12), D(2026, 12, 18), 120, 40, None, "In Progress", None, ""),
    ("PRJ-302", "Toolchain migration", "2.0.0", "Ben", D(2026, 10, 19), D(2026, 11, 27), 60, 12, None, "In Progress", None, "license order pending"),
    ("PRJ-303", "Requirements review with customer", "2.0.0", "Chris", D(2027, 1, 11), D(2027, 2, 5), 40, None, None, "Open", None, ""),
    ("PRJ-304", "Prototype of the new HMI", "2.0.0", "Dana", None, None, None, None, 15, "In Progress", None, "rough estimate"),
]


def tasks_sheet() -> Sheet:
    rows = [[Styled(f.header, "header") for f in labels.FIELDS]]
    for t in TASKS:
        rows.append(list(t))
    return Sheet(labels.SHEET_TASKS, rows, freeze_header=True,
                 widths={0: 12, 1: 52, 2: 10, 3: 10, 4: 12, 5: 12, 6: 12, 7: 10, 8: 12, 9: 12, 10: 14, 11: 40})


def settings_sheet() -> Sheet:
    rows = [
        list(labels.SETTINGS_HEADERS),
        ["Project", "Project Falcon"],
        ["Customer", "ACME Automotive"],
        ["Report date", REPORT_DATE],
        ["Version order", "1.2.0, 2.0.0, 1.1.0"],
    ]
    return Sheet(labels.SHEET_SETTINGS, rows, widths={0: 22, 1: 30})


def as_task(row) -> Task:
    tid, title, version, owner, start, due, est, spent, pct, state, override, comment = row
    est_h = 16.0 if est == "2d" else est
    return Task(tid, title, version, owner, start, due, est_h, spent, pct, state,
                override.lower() if override else None, comment)


def past_week(day: dt.date, share: float) -> history.Snapshot:
    """The project as it looked on ``day``: hours and manual % scaled by ``share``."""
    tasks = []
    for row in TASKS:
        t = as_task(row)
        if t.id in NEW_THIS_WEEK:
            continue
        if t.spent_h is not None:
            t.spent_h = round(t.spent_h * share, 1)
        if t.progress_manual is not None and t.progress_manual < 100:
            t.progress_manual = round(t.progress_manual * share)
        if t.state.lower() in ("done", "closed", "verified") and share < 0.9:
            t.state, t.progress_manual = "In Progress", None
        t.light_override = None
        tasks.append(copy.copy(t))
    settings = Settings(report_date=day)
    versions, _ = compute(tasks, settings)
    return history.snapshot_from(versions, day)


def main() -> None:
    folder = os.path.abspath(sys.argv[1] if len(sys.argv) > 1 else os.path.join(ROOT, "out", "demo"))
    if os.path.isdir(folder):
        shutil.rmtree(folder)
    os.makedirs(folder)
    with open(os.path.join(folder, "Tasks.xlsx"), "wb") as fh:
        fh.write(workbook_bytes([tasks_sheet(), settings_sheet()]))
    hist = os.path.join(folder, history.FOLDER)
    for weeks_back in range(6, 0, -1):  # W40 .. W45
        day = REPORT_DATE - dt.timedelta(weeks=weeks_back)
        history.save(hist, past_week(day, 1 - weeks_back * 0.13), dt.datetime(2026, 1, 1))
    print(f"demo written to {folder}")


if __name__ == "__main__":
    main()
