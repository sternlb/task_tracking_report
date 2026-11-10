"""M1 spike (T-1.5): title slide, trend chart and a status table with sample data.

Goal: PowerPoint opens the result without a repair prompt. The layout of the
real slides is built in M5; this only exercises every XML construct they need.

    python tools/spike.py [--template customer.pptx] [--out out/spike.pptx]
"""

from __future__ import annotations

import argparse
import os
import sys

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, ROOT)

from statusreport.pptx.deck import Deck  # noqa: E402
from statusreport.pptx.shapes import Box, Cell, Para, Row, Run, inch  # noqa: E402

LIGHT = {"green": "#2E9E44", "orange": "#F0A202", "red": "#D62828", "grey": "#A0A0A0"}
GREY_TEXT = "#7F7F7F"

WEEKS = ["CW 41", "CW 42", "CW 43", "CW 44", "CW 45", "CW 46"]
TREND = [
    ("1.1.0", [55, 68, 80, 91, 100, 100]),
    ("1.2.0", [5, 14, 26, 38, 55, 62]),
    ("2.0.0", [None, None, 0, 4, 9, 12]),
]
TASKS = [
    # id, title, owner, due, now, prev, light, prev light, comment
    ("PRJ-101", "Login dialog with SSO", "Anna", "20 Nov", 65, 50, "green", "green", ""),
    ("PRJ-102", "CAN bus driver for new ECU", "Ben", "13 Nov", 80, 70, "orange", "green", "HW sample late"),
    ("PRJ-103", "Diagnostic trouble codes", "Chris", "06 Nov", 90, 90, "red", "orange", "overdue, review open"),
    ("PRJ-104", "Bootloader update over UDS", "Dana", "04 Dec", 30, 20, "green", "green", ""),
    ("PRJ-105", "Logging & trace export", "Eli", "18 Dec", None, None, "grey", None, "no estimate yet"),
    ("PRJ-106", "Unit tests for CRC module", "Fay", "27 Nov", 100, 85, "green", "green", "done"),
]


def progress_cell(now, prev, size=14) -> Cell:
    if now is None:
        return Cell([Para([Run("–", size, color=GREY_TEXT)])], align="ctr")
    prev_txt = f" ({prev:.0f} %)" if prev is not None else " (new)"
    return Cell([Para([Run(f"{now:.0f} %", size, bold=True), Run(prev_txt, 10, color=GREY_TEXT)])], align="ctr")


def light_cell(now, prev) -> Cell:
    runs = [Run("●", 18, color=LIGHT[now])]
    if prev and prev != now:
        runs.append(Run(" (", 10, color=GREY_TEXT))
        runs.append(Run("●", 13, color=LIGHT[prev]))
        runs.append(Run(")", 10, color=GREY_TEXT))
    return Cell([Para(runs)], align="ctr")


def text_cell(text, size=11, bold=False, color=None, fill=None, align=None) -> Cell:
    return Cell([Para([Run(text, size, bold, color)])], fill=fill, align=align)


def build(template: str | None) -> Deck:
    deck = Deck.from_file(template) if template else Deck.default()

    s = deck.add_slide("Title Slide")
    s.set_title([Para([Run("Project Falcon")])])
    s.set_subtitle([Para([Run("Status report · ACME Automotive · CW 46 / 2026 · 10 Nov 2026")])])

    s = deck.add_slide("Title Only")
    s.set_title([Para([Run("Progress over time")])])
    s.add_line_chart(s.content_area(), WEEKS, TREND)

    s = deck.add_slide("Title Only")
    s.set_title([Para([Run("Version 1.2.0")])])
    area = s.content_area()
    # Owner and Due are input only (Due feeds the traffic light), not shown to the customer.
    shares = [11, 40, 17, 10, 22]
    widths = [area.w * p // 100 for p in shares]
    widths[-1] += area.w - sum(widths)
    head = Row(
        [text_cell(h, 11, True, "bg1", "accent1", "ctr" if i in (2, 3) else None)
         for i, h in enumerate(["ID", "Task", "Progress (prev. week)", "Status", "Comment"])],
        inch(0.35),
    )
    total = Row(
        [
            text_cell("Version total", 11, True),
            text_cell(""),
            progress_cell(62, 55),
            light_cell("orange", "green"),
            text_cell(""),
        ]
    )
    for c in total.cells:
        c.fill = "accent1/20"
    rows = [head, total]
    for n, (tid, title, _owner, _due, now, prev, light, prev_light, comment) in enumerate(TASKS):
        zebra = "#F2F2F2" if n % 2 else None
        cells = [
            text_cell(tid), text_cell(title),
            progress_cell(now, prev), light_cell(light, prev_light), text_cell(comment, 10, color=GREY_TEXT),
        ]
        for c in cells:
            c.fill = zebra
        rows.append(Row(cells))
    s.add_table(area, widths, rows)
    # Rows grow with their text, so the declared heights are only a minimum:
    # anchor the legend to the bottom of the content area instead.
    s.add_textbox(
        Box(area.x, area.y + area.h - inch(0.3), area.w, inch(0.3)),
        [Para([
            Run("● ", 10, color=LIGHT["green"]), Run("on track   ", 9, color=GREY_TEXT),
            Run("● ", 10, color=LIGHT["orange"]), Run("≤ 25 pp behind   ", 9, color=GREY_TEXT),
            Run("● ", 10, color=LIGHT["red"]), Run("> 25 pp behind or overdue   ", 9, color=GREY_TEXT),
            Run("● ", 10, color=LIGHT["grey"]), Run("not assessable   ·   (…) = previous week", 9, color=GREY_TEXT),
        ])],
    )
    return deck


def main() -> None:
    ap = argparse.ArgumentParser()
    ap.add_argument("--template")
    ap.add_argument("--out", default=os.path.join(ROOT, "out", "spike.pptx"))
    args = ap.parse_args()
    deck = build(args.template)
    os.makedirs(os.path.dirname(os.path.abspath(args.out)), exist_ok=True)
    deck.save(args.out)
    for w in deck.warnings:
        print("warning:", w)
    print(f"written {args.out}")


if __name__ == "__main__":
    main()
