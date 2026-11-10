"""Report model -> PowerPoint deck — R-OUT-1 … R-OUT-8, design 8.2, T-5.2/T-5.3, T-6.2.

Colours: header and total row use theme slots (accent1, bg1), so a customer
template changes them; only the traffic lights are fixed RGB (D-4).

The table is fitted to the content area of the chosen layout (M6): on a small
or narrow area the fonts shrink (down to 80 %), texts are shortened to one line
of their column, and fewer rows go on a slide than ``Rows per slide`` allows.
"""

from __future__ import annotations

from dataclasses import dataclass
from typing import List, Optional

from . import labels
from .pptx.deck import Deck, Slide
from .pptx.shapes import EMU_PER_PT, Box, Cell, Para, Row, Run, inch
from .report import Report, Row as ReportRow, TablePage, paginate, shorten

LIGHT_RGB = {"green": "#2E9E44", "orange": "#F0A202", "red": "#D62828", "grey": "#A0A0A0"}
GREY_TEXT = "#7F7F7F"
ZEBRA = "#F2F2F2"
HEADER_FILL, HEADER_TEXT, TOTAL_FILL = "accent1", "bg1", "accent1/20"

COLUMN_SHARES = (11, 40, 17, 10, 22)  # ID, Task, Progress, Status, Comment (percent of the width)
CENTERED = (2, 3)
LEGEND_H = inch(0.3)
GAP = inch(0.1)
REFERENCE_W = 10972800  # content width of the built-in base: fonts at 100 %
MIN_SCALE = 0.8
CELL_MARGIN_PT = 2 * 68580 / EMU_PER_PT  # left + right cell margin
CELL_PAD = 2 * 27432 + inch(0.06)  # top + bottom margin plus some air
CHAR_WIDTH = 0.52  # average character width in em; on the safe side for Calibri, Arial, Segoe
MIN_ROWS = 3


@dataclass
class Metrics:
    scale: float
    widths: List[int]
    text_pt: float
    small_pt: float
    value_pt: float
    header_h: int
    row_h: int
    title_chars: int
    comment_chars: int
    per_page: int


def _widths(total: int) -> List[int]:
    widths = [total * p // 100 for p in COLUMN_SHARES]
    widths[-1] += total - sum(widths)
    return widths


def _chars(width_emu: int, size_pt: float) -> int:
    """How many characters fit on one line of a column."""
    return max(8, int((width_emu / EMU_PER_PT - CELL_MARGIN_PT) / (size_pt * CHAR_WIDTH)))


def _line_h(size_pt: float, lines: int = 1) -> int:
    return int(lines * size_pt * 1.2 * EMU_PER_PT)


def metrics(area: Box, rows_per_slide: int) -> Metrics:
    scale = max(MIN_SCALE, min(1.0, area.w / REFERENCE_W))
    widths = _widths(area.w)
    text_pt, small_pt, value_pt = 11 * scale, 10 * scale, 14 * scale
    header_lines = max(
        1 if len(h) <= _chars(w, text_pt) else 2 for h, w in zip(labels.TABLE_HEADERS, widths)
    )
    header_h = _line_h(text_pt, header_lines) + CELL_PAD
    row_h = max(_line_h(value_pt) + CELL_PAD, int(inch(0.34) * scale))
    available = area.h - LEGEND_H - GAP - header_h
    fit = available // row_h - 1  # minus the total row
    return Metrics(
        scale=scale, widths=widths, text_pt=text_pt, small_pt=small_pt, value_pt=value_pt,
        header_h=header_h, row_h=row_h,
        title_chars=_chars(widths[1], text_pt), comment_chars=_chars(widths[4], small_pt),
        per_page=max(MIN_ROWS, min(rows_per_slide, fit)),
    )


def _text(text: str, size: float, bold: bool = False, color: Optional[str] = None, fill: Optional[str] = None, align: Optional[str] = None) -> Cell:
    return Cell([Para([Run(text, size, bold, color)])], fill=fill, align=align)


def _progress_cell(row: ReportRow, m: Metrics) -> Cell:
    if row.progress is None:
        runs = [Run(labels.NO_VALUE, m.value_pt, color=GREY_TEXT)]
    else:
        runs = [Run(f"{row.progress:.0f} %", m.value_pt, bold=True)]
    runs.append(Run(" " + row.prev_text, m.small_pt, color=GREY_TEXT))
    return Cell([Para(runs)], align="ctr")


def _light_cell(row: ReportRow, m: Metrics) -> Cell:
    runs = [Run(labels.LIGHT_DOT, 18 * m.scale, color=LIGHT_RGB.get(row.light, LIGHT_RGB["grey"]))]
    if row.prev_light:
        runs += [
            Run(" (", m.small_pt, color=GREY_TEXT),
            Run(labels.LIGHT_DOT, 13 * m.scale, color=LIGHT_RGB.get(row.prev_light, LIGHT_RGB["grey"])),
            Run(")", m.small_pt, color=GREY_TEXT),
        ]
    return Cell([Para(runs)], align="ctr")


def _table_rows(page: TablePage, m: Metrics) -> List[Row]:
    header = Row(
        [_text(h, m.text_pt, True, HEADER_TEXT, HEADER_FILL, "ctr" if i in CENTERED else None)
         for i, h in enumerate(labels.TABLE_HEADERS)],
        m.header_h,
    )
    total = page.total
    label = _text(total.id, m.text_pt, True)
    label.span = 2  # "Version total" across ID and Task, so it never wraps
    total_cells = [label, _progress_cell(total, m), _light_cell(total, m), _text("", m.text_pt)]
    for c in total_cells:
        c.fill = TOTAL_FILL
    rows = [header, Row(total_cells, m.row_h)]
    for n, r in enumerate(page.rows):
        cells = [
            _text(r.id, m.text_pt),
            _text(shorten(r.title, m.title_chars), m.text_pt),
            _progress_cell(r, m),
            _light_cell(r, m),
            _text(shorten(r.comment, m.comment_chars), m.small_pt, color=GREY_TEXT),
        ]
        if n % 2:
            for c in cells:
                c.fill = ZEBRA
        rows.append(Row(cells, m.row_h))
    return rows


def _legend(report: Report, scale: float) -> List[Para]:
    runs: List[Run] = []
    for light, text in labels.LEGEND:
        runs.append(Run(labels.LIGHT_DOT + " ", 10 * scale, color=LIGHT_RGB[light]))
        runs.append(Run(text.format(orange=report.orange_pp, red=report.red_pp) + "    ", 9 * scale, color=GREY_TEXT))
    runs.append(Run(labels.LEGEND_PREVIOUS, 9 * scale, color=GREY_TEXT))
    return [Para(runs)]


def _title(slide: Slide, text: str) -> None:
    slide.set_title([Para([Run(text)])])


def _table_slide(deck: Deck, layout: str, report: Report, page: TablePage, m: Metrics, footer: str) -> None:
    slide = deck.add_slide(layout, "content", footer)
    _title(slide, page.title)
    area = slide.content_area()
    slide.add_table(area, m.widths, _table_rows(page, m))
    # rows grow with their text, so the legend sits at the bottom of the area, not under the table
    slide.add_textbox(Box(area.x, area.y + area.h - LEGEND_H, area.w, LEGEND_H), _legend(report, m.scale), anchor="b")


def _trend_slide(deck: Deck, layout: str, report: Report, footer: str) -> None:
    slide = deck.add_slide(layout, "content", footer)
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


def _burndown_slide(deck: Deck, layout: str, report: Report, footer: str) -> None:
    """R-OUT-10: remaining effort of all tasks per week, with the ideal line to the latest Due."""
    bd = report.burndown
    slide = deck.add_slide(layout, "content", footer)
    _title(slide, labels.BURNDOWN_TITLE)
    area = slide.content_area()
    chart = Box(area.x, area.y, area.w, area.h - LEGEND_H)
    slide.add_line_chart(
        chart, bd.labels,
        [(labels.BURNDOWN_ACTUAL, bd.actual), (labels.BURNDOWN_IDEAL, bd.ideal)],
        y_max=None, y_step=None, value_format='#,##0" h"',
        styles=[{"color": "accent1"}, {"color": "#7F7F7F", "dash": True, "marker": False}],
    )
    slide.add_textbox(
        Box(area.x, area.y + area.h - LEGEND_H, area.w, LEGEND_H),
        [Para([Run(report.burndown_note, 10, color=GREY_TEXT)])], anchor="b",
    )


def render(
    report: Report, deck: Deck, title_layout: str, content_layout: str, rows_per_slide: int, footer_text: str = ""
) -> Deck:
    """Add all slides of ``report`` to ``deck`` (R-OUT-1 order). Returns the deck.

    Content slides carry the template's slide number and footer placeholders (R-OUT-8).
    """
    s = deck.add_slide(title_layout, "title")
    _title(s, report.title)
    if not s.set_subtitle([Para([Run(report.subtitle)])]):
        area = s.content_area()  # layout without subtitle placeholder
        s.add_textbox(Box(area.x, area.y, area.w, inch(0.6)), [Para([Run(report.subtitle, 18)])])
    if report.trend_series:
        _trend_slide(deck, content_layout, report, footer_text)
    if report.burndown is not None:
        _burndown_slide(deck, content_layout, report, footer_text)
    m = metrics(deck.content_area(deck.layout(content_layout, "content")), rows_per_slide)
    for table in report.tables:
        for page in paginate(table, m.per_page):
            _table_slide(deck, content_layout, report, page, m, footer_text)
    return deck


__all__ = ["render", "metrics", "LIGHT_RGB"]
