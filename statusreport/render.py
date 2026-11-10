"""Report model -> PowerPoint deck — R-OUT-1 … R-OUT-8, design 8.2, T-5.2/T-5.3.

Colours: header and total row use theme slots (accent1, bg1), so a customer
template changes them; only the traffic lights are fixed RGB (D-4).
"""

from __future__ import annotations

from typing import List, Optional

from . import labels
from .pptx.deck import Deck, Slide
from .pptx.shapes import Box, Cell, Para, Row, Run, inch
from .report import Report, Row as ReportRow, TablePage

LIGHT_RGB = {"green": "#2E9E44", "orange": "#F0A202", "red": "#D62828", "grey": "#A0A0A0"}
GREY_TEXT = "#7F7F7F"
ZEBRA = "#F2F2F2"
HEADER_FILL, HEADER_TEXT, TOTAL_FILL = "accent1", "bg1", "accent1/20"

COLUMN_SHARES = (11, 40, 17, 10, 22)  # ID, Task, Progress, Status, Comment (percent of the width)
CENTERED = (2, 3)
HEADER_H = inch(0.36)
ROW_H = inch(0.34)
MIN_ROW_H = inch(0.28)
LEGEND_H = inch(0.3)


def _widths(total: int) -> List[int]:
    widths = [total * p // 100 for p in COLUMN_SHARES]
    widths[-1] += total - sum(widths)
    return widths


def _text(text: str, size: float = 11, bold: bool = False, color: Optional[str] = None, fill: Optional[str] = None, align: Optional[str] = None) -> Cell:
    return Cell([Para([Run(text, size, bold, color)])], fill=fill, align=align)


def _progress_cell(row: ReportRow) -> Cell:
    if row.progress is None:
        runs = [Run(labels.NO_VALUE, 14, color=GREY_TEXT)]
    else:
        runs = [Run(f"{row.progress:.0f} %", 14, bold=True)]
    runs.append(Run(" " + row.prev_text, 10, color=GREY_TEXT))
    return Cell([Para(runs)], align="ctr")


def _light_cell(row: ReportRow) -> Cell:
    runs = [Run(labels.LIGHT_DOT, 18, color=LIGHT_RGB.get(row.light, LIGHT_RGB["grey"]))]
    if row.prev_light:
        runs += [
            Run(" (", 10, color=GREY_TEXT),
            Run(labels.LIGHT_DOT, 13, color=LIGHT_RGB.get(row.prev_light, LIGHT_RGB["grey"])),
            Run(")", 10, color=GREY_TEXT),
        ]
    return Cell([Para(runs)], align="ctr")


def _table_rows(page: TablePage, row_h: int) -> List[Row]:
    header = Row(
        [_text(h, 11, True, HEADER_TEXT, HEADER_FILL, "ctr" if i in CENTERED else None) for i, h in enumerate(labels.TABLE_HEADERS)],
        HEADER_H,
    )
    total = page.total
    total_cells = [_text(total.id, 11, True), _text(""), _progress_cell(total), _light_cell(total), _text("")]
    for c in total_cells:
        c.fill = TOTAL_FILL
    rows = [header, Row(total_cells, row_h)]
    for n, r in enumerate(page.rows):
        cells = [_text(r.id), _text(r.title), _progress_cell(r), _light_cell(r), _text(r.comment, 10, color=GREY_TEXT)]
        if n % 2:
            for c in cells:
                c.fill = ZEBRA
        rows.append(Row(cells, row_h))
    return rows


def _legend(report: Report) -> List[Para]:
    runs: List[Run] = []
    for light, text in labels.LEGEND:
        runs.append(Run(labels.LIGHT_DOT + " ", 10, color=LIGHT_RGB[light]))
        runs.append(Run(text.format(orange=report.orange_pp, red=report.red_pp) + "    ", 9, color=GREY_TEXT))
    runs.append(Run(labels.LEGEND_PREVIOUS, 9, color=GREY_TEXT))
    return [Para(runs)]


def _title(slide: Slide, text: str) -> None:
    slide.set_title([Para([Run(text)])])


def _table_slide(deck: Deck, layout: str, report: Report, page: TablePage, max_rows: int) -> None:
    slide = deck.add_slide(layout)
    _title(slide, page.title)
    area = slide.content_area()
    usable = area.h - LEGEND_H - inch(0.1) - HEADER_H
    row_h = max(MIN_ROW_H, min(ROW_H, usable // max(1, max_rows + 1)))
    slide.add_table(area, _widths(area.w), _table_rows(page, row_h))
    # rows grow with their text, so the legend sits at the bottom of the area, not under the table
    slide.add_textbox(Box(area.x, area.y + area.h - LEGEND_H, area.w, LEGEND_H), _legend(report), anchor="b")


def _trend_slide(deck: Deck, layout: str, report: Report) -> None:
    slide = deck.add_slide(layout)
    _title(slide, labels.TREND_TITLE)
    area = slide.content_area()
    chart = area
    if report.trend_note:
        chart = Box(area.x, area.y, area.w, area.h - LEGEND_H)
        slide.add_textbox(
            Box(area.x, area.y + area.h - LEGEND_H, area.w, LEGEND_H),
            [Para([Run(report.trend_note, 11, color=GREY_TEXT)])], anchor="b",
        )
    slide.add_line_chart(chart, report.trend_labels, report.trend_series)


def render(report: Report, deck: Deck, title_layout: str, content_layout: str, rows_per_slide: int) -> Deck:
    """Add all slides of ``report`` to ``deck`` (R-OUT-1 order). Returns the deck."""
    s = deck.add_slide(title_layout)
    _title(s, report.title)
    if not s.set_subtitle([Para([Run(report.subtitle)])]):
        area = s.content_area()  # layout without subtitle placeholder
        s.add_textbox(Box(area.x, area.y, area.w, inch(0.6)), [Para([Run(report.subtitle, 18)])])
    if report.trend_series:
        _trend_slide(deck, content_layout, report)
    for page in report.pages:
        _table_slide(deck, content_layout, report, page, rows_per_slide)
    return deck


__all__ = ["render", "LIGHT_RGB"]
