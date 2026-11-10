# Design — SW Status Report

Status: **draft v0.1** (2026-11-10) · Covers: [requirements.md](requirements.md)

## 1. Overview

```
 input.xlsx ─┐                        history/2026-W45.json ─┐
 (Settings,  │                        history/2026-W44.json ─┤
  Tasks,     ├─► loader ─► tasks ─► calc ─► report model ─┬──┴─► trend data
  Import,    │   (map,              (progress, expected,  │
  Mapping)   │    merge,             lights, versions)    ├─► history/2026-W46.json
 polarion ───┘    validate)                               │
 .xlsx (opt.)                                             └─► pptx builder ─► deck.pptx
                                       base.pptx or customer template ──┘
```

One pass, no state apart from the `history/` folder. The input Excel is only read.

## 2. Key decisions

| # | Decision | Reason |
|---|---|---|
| D-1 | Write and read Office files ourselves as Office Open XML (`zipfile` + `xml.etree`) | Constraint C-1: standard library only. python-pptx needs lxml (compiled), so it cannot be bundled |
| D-2 | **One code path for both modes:** the deck is always built by opening a base `.pptx`, removing its slides and adding new ones. Without a customer template the tool uses a built-in neutral `base.pptx` | The template mode (R-OUT-8) then gets tested with every run, not only when the customer template arrives. Generating masters and layouts from scratch is the most fragile part of the format, so it is done only once |
| D-3 | Content is placed in a **content area** taken from the chosen layout (body placeholder bounds, otherwise "below the title, inside the margins") | The tool must work without knowing the template |
| D-4 | Colours: table header and summary row use **theme colours** (`accent1`, `tx1`, `bg1`); traffic lights use **fixed RGB** | The deck adopts the customer's look, while green/orange/red keep their meaning |
| D-5 | Chart = native `c:lineChart` with cached values **and** an embedded workbook | Renders without recalculation and stays editable ("Edit data") in PowerPoint |
| D-6 | Distribution as a **zipapp** `statusreport.pyz` (stdlib `zipapp`) plus `make_report.cmd` | One file to copy, starts with `python statusreport.pyz` or by double-click with the `py` launcher |
| D-7 | History as one JSON file per ISO week | Human-readable, mergeable in Git, easy to fix or delete by hand |
| D-8 | Previous value and previous light shown in the same pattern: **big current value (small previous value)** | Consistent reading, and fulfils R-OUT-4 and R-OUT-6 with one design |

## 3. Module structure

```
sw-status-report/
├── statusreport/
│   ├── __main__.py        CLI (argparse), orchestration, exit codes
│   ├── labels.py          all visible texts (EN), one dict per language
│   ├── defaults.py        default settings, column names, colours
│   ├── model.py           dataclasses: Settings, Task, VersionStatus, Snapshot, Issue
│   ├── xlsx_read.py       minimal .xlsx reader → list of rows per sheet
│   ├── xlsx_write.py      minimal .xlsx writer (template, chart workbook)
│   ├── parse.py           dates, durations ("1d 4h"), percent, lights, natural sort
│   ├── loader.py          Settings, Mapping, Import, Tasks → Task list + Issues
│   ├── template.py        the empty input workbook for --init
│   ├── calc.py            progress, expected, lights, version aggregation
│   ├── history.py         load/save snapshots, previous week, trend series
│   ├── report.py          builds the presentation-neutral report model
│   ├── render.py          report model -> slides on a Deck
│   ├── output.py          output file name, overwrite suffix, template path
│   └── pptx/
│       ├── opc.py         package read/write: parts, rels, [Content_Types].xml
│       ├── deck.py        open base/template, strip slides, find layouts, add slides
│       ├── shapes.py      text boxes, placeholders, table (a:tbl), circles
│       ├── chart.py       chart part + embedded workbook
│       └── base.pptx      built-in neutral 16:9 template (asset)
├── tools/
│   ├── build_base_pptx.py generates base.pptx once (dev only)
│   ├── build_pyz.py       builds dist/statusreport.pyz
│   └── opc_lint.py        checks a pptx: all rels resolve, content types complete
├── tests/                 unittest, fixtures/ (xlsx from Excel and from Polarion)
├── make_report.cmd
└── docs/                  requirements.md, design.md, tasks.md
```

## 4. Data model

```python
@dataclass
class Task:
    id: str; title: str; version: str
    owner: str = ""; start: date | None = None; due: date | None = None
    estimate_h: float | None = None; spent_h: float | None = None
    progress_manual: float | None = None; state: str = ""
    light_override: str | None = None; comment: str = ""
    source: str = "Tasks"          # "Tasks", "Import" or "Import+Tasks"
    row: int = 0                   # for messages
    # computed by calc
    progress: float | None = None; expected: float | None = None; light: str = "grey"

@dataclass
class VersionStatus:
    name: str; progress: float | None; expected: float | None; light: str
    tasks: list[Task]

@dataclass
class Issue:
    level: str   # "error" (row skipped) | "warning"
    sheet: str; row: int; message: str
```

Lights are the strings `green`, `orange`, `red`, `grey`.

## 5. Input processing

### 5.1 xlsx reader (`xlsx_read.py`)

- Reads `xl/workbook.xml` (sheet names, `date1904`), `xl/_rels/workbook.xml.rels`,
  `xl/sharedStrings.xml`, `xl/styles.xml`, `xl/worksheets/sheetN.xml`.
- Cell types: shared string, inline string, number, boolean, formula with cached value
  (`<v>`), error → empty plus warning.
- Date detection via the cell style's number format: built-in IDs 14–22 and 45–47, or a
  custom format containing `d`, `m` or `y` outside quotes and brackets.
- Returns `list[list[value]]` per sheet, with gaps filled from the cell reference
  (`C7`), so empty cells keep their column position.
- Old `.xls` files are rejected with the hint "save as .xlsx".

### 5.2 Loader (`loader.py`)

1. Read `Settings` (key/value table, column A = key, column B = value) → `Settings`,
   with defaults for empty values.
2. Import source: the `--polarion` file (its first sheet) or else the `Import` sheet. If
   the sheet is empty, no import is used.
3. Apply `Mapping`: field → header text (case-insensitive, trimmed). A mapped header
   that is missing in the import is a structural error.
4. Read `Tasks` by its fixed headers.
   An empty mapped Time Spent becomes `0.0` already here (R-IN-11), so the merge below
   lets a filled `Tasks` cell win as usual.
5. Merge by ID (R-IN-7): start with the import rows, then let non-empty `Tasks` cells
   win; `Tasks` IDs that are not in the import are added.
6. Validate and parse each row (`parse.py`) → `Task` or `Issue`. All errors of a row are
   reported together, then the row is skipped. Issues are sorted by sheet and row.

Template (`template.py`): sheets `Read me` (short user guide, opens first), `Tasks`
(styled header, frozen first row, ID and Version formatted as text, Start/Due as dates,
drop-down for `Light override`, example row `#EXAMPLE-1`), `Import` (empty), `Mapping`
(typical Polarion names: ID, Title, Planned In, Assignee, Due Date, Initial Estimate,
Time Spent, Status; Start is empty because Polarion has no standard start field) and
`Settings` (value plus a description per setting).

### 5.3 Parsing rules (`parse.py`)

| Value | Accepted |
|---|---|
| Date | Excel date serial (1900 and 1904 system), `YYYY-MM-DD`, `DD.MM.YYYY`, `DD/MM/YYYY`, optional time part ignored |
| Duration | Number = hours. Text: regex `(\d+(?:[.,]\d+)?)\s*([wdhm])` repeated, e.g. `1w 2d 4h 30m`; d = hours per day, w = 5 d |
| Percent | `65`, `65 %`, `0.65` formatted as a percent in Excel (cell format contains `%`) → 65 |
| Light | `green/orange/red` case-insensitive, also `g/o/r`, `amber`, `yellow` → orange |
| Done state | case-insensitive comparison against the `Done states` setting |

## 6. Calculation (`calc.py`)

Directly implements R-CA-1 to R-CA-5. Details:

- Empty Start (R-CA-6): `derive_start(due, estimate_h, hours_per_day)` counts
  `ceil(estimate_h / hours_per_day)` working days back, with Due as the last one. It sets
  `Task.start` and `Task.start_derived = True` before anything else is computed, so the
  expected %, the version start and the history all use the derived date.
- Expected % = `(report_date − start) / (due − start) × 100`, limited to 0–100. If
  `due == start`: 0 before, 100 from that day. (The derived start uses working days, the
  elapsed share uses calendar days, as R-CA-2 says.)
- The thresholds compare percentage points and include the limit (gap = 10 → green
  with the orange threshold at 10).
- Version aggregation with weights `w = estimate_h` (or 1 for all, see R-CA-4). Version
  expected % uses the same weights on the tasks' expected %.
- Version light: gap = weighted expected − weighted progress over the tasks that have
  both values. A version at 100 % is green; overdue = after the latest Due.
- Comparisons add `1e-6` to the thresholds, so float noise (`10.000000000000004`) does
  not flip a light at the limit.
- All values are kept as `float`; rounding to whole % only happens when displayed.
- `compute(tasks, settings)` fills `Task.progress/expected/light` and returns the
  `VersionStatus` list in slide order (Settings `Version order` first, then natural sort)
  plus warnings for tasks without progress. Inside a version the input order is kept.
- `--check` prints the computed values per version and task, for checking by hand.

## 7. History (`history.py`)

File `history/2026-W46.json`:

```json
{
  "schema": 1,
  "week": "2026-W46",
  "report_date": "2026-11-10",
  "generated_at": "2026-11-10T09:12:44",
  "versions": { "1.2.0": { "progress": 62.5, "light": "orange" } },
  "tasks": { "PRJ-123": { "version": "1.2.0", "title": "Login", "progress": 65.0, "light": "green" } }
}
```

- Write as temp file, then rename (atomic, nothing half-written).
- Previous week = largest week key `< current week` (R-HI-3).
- Trend series: the last `N` weeks *that have a snapshot*, up to and including the
  current one. The in-memory current snapshot replaces a stored one of the same week
  (so `--no-save` shows the right values), and snapshots after the report week are
  ignored. Versions = the versions in the current report. Missing values → empty point,
  the chart shows a gap.
- Axis labels `CW 46`; with the two-digit year (`CW 53/26`, `CW 01/27`) when the weeks
  span a year boundary.
- Values are stored rounded to 2 decimals. Version snapshots also store `expected`
  (not shown yet, kept for later use).
- Unreadable snapshot (bad JSON, unknown schema, week in the file ≠ file name) →
  warning, the week is skipped. Other files in `history/` are ignored.
- `generated_at` makes snapshot files differ between runs; NF-2 (determinism) applies
  to the deck, not to the history.

## 8. PowerPoint generation

### 8.1 Deck assembly (`pptx/deck.py`)

1. Open the base (`base.pptx` or `--template`/Settings `Template file`).
2. Remove all existing slides: entries in `presentation.xml/p:sldIdLst`, their rels, their
   parts and dependent parts (notes slides, charts and embeddings used only by them),
   and their entries in `[Content_Types].xml`.
3. Find the layouts by name (`p:cSld/@name`, case-insensitive). The settings
   `Title layout` (default `Title Slide`) and `Content layout` (default `Title Only`)
   can be changed. Fallback by `sldLayout/@type` (M6): title slide → `title`; content →
   `titleOnly`, `obj`, `tx`; then a layout with a `ctrTitle` (title) or the title layout
   with the fewest placeholders (content). One warning per name and role lists the
   available layout names. A `.potx` (template main content type) is accepted, and its
   main part is switched to the presentation content type.
4. Add the slides: each slide references its layout, and title placeholders are filled
   via `p:ph type="title"/"ctrTitle"/"subTitle"`, so position and font come from the
   template.
5. Content area = bounds of the layout's body/object placeholder (`idx=1`), otherwise
   the slide area below the title placeholder minus 0.4" margins. In both cases it ends
   0.1" above the highest footer placeholder (`dt`, `ftr`, `sldNum`) in the lower half
   of the area. The table and chart are placed inside it.
5a. Footers (M6): content slides get a copy of the layout's `sldNum` placeholder with a
   `slidenum` field, and its `ftr` placeholder if there is a footer text (Settings
   `Footer text`, else the layout's own text). Nothing on the title slide; `dt` is not used.
5b. Table fitting (`render.metrics`, M6): scale = content width / width of the built-in
   base, limited to 0.8–1.0 for all font sizes. The number of characters per column is
   estimated with an average character width of 0.52 em, and texts are shortened to one
   line. The header gets two lines if a header text would not fit. Row height comes from
   the value font (14 pt × scale). Rows per slide = min(setting, what fits), at least 3.
   `report.paginate` then splits the version tables. "Version total" spans the ID and
   Task columns (`gridSpan`).
6. Write the package and run the OPC checks from `tools/opc_lint.py` as a self-test
   before saving.

### 8.2 Slides

**Title slide:** title = project name, subtitle = `Status report · Customer · CW 46 / 2026 · 10 Nov 2026`.

**Progress over time:** title `Progress over time`, line chart with markers, one series
per version, y-axis 0–100 % fixed, x-axis `CW 41 … CW 46` (with the year when the range
crosses a year boundary), legend at the bottom. With only one snapshot the chart shows
points instead of lines, plus a note "Trend builds up from next week".

**Burndown (R-OUT-10, `burndown.py`):** title `Burndown: remaining effort`, line chart
with automatic y-axis in `#,##0" h"`. Two series: `Remaining effort (h)` (accent1, with
markers) for the past weeks with a snapshot, and `Ideal` (grey, dashed, no markers),
linear in calendar time from the first actual value to 0 h on the Monday of the week of
the latest Due. Future weeks up to that week (max. 26) are appended to the x-axis with
gaps in the actual series. If the latest Due lies before the first week shown, there is
no ideal line. A note line at the bottom gives the remaining and total hours and the
end date. Snapshots carry `"effort": {"total_h", "remaining_h"}`; for older ones the
value is derived (`derived_weeks`, shown in the console).

**Version table:** title `Version 1.2.0` (continued: `Version 1.2.0 (2/3)`).

| ID | Task | Progress (prev. week) | Status | Comment |
|---|---|---|---|---|
| **Version total** | | **62 %** (55 %) | ● (●) | |
| PRJ-123 | Login dialog | **65 %** (50 %) | ● (●) | waiting for API |

Owner and Due are not shown to the customer (R-OUT-3).

- Progress cell: run 1 = current value, bold, 14 pt; run 2 = ` (50 %)`, 10 pt grey.
- Status cell: run 1 = `●` 18 pt in the light colour, then ` (●)` with the dot at 13 pt in
  last week's colour and grey brackets, only if different from today. This makes changes
  stand out (R-OUT-6).
- Header row filled with `accent1`, summary row filled with `accent1` at 20 %, zebra
  rows light grey. Font from the theme (`+mn-lt`).
- Column widths as proportions of the content area: 11/40/17/10/22 %.
- The legend sits at the bottom of the content area, not under the last row: rows grow
  with their text, so the declared row heights are only a minimum (seen in the M1 spike).
- Titles longer than 62 characters and comments longer than 45 characters are shortened
  with `…` (tuned in the M5 check: 62 characters fit one line of the Task column).
- Row height 0.34" (header 0.36"), reduced down to 0.28" if `Rows per slide` would not
  fit the content area otherwise.
- The total row is repeated on every continuation slide of a version.
- Legend line at the bottom of the content area, 9 pt, using the thresholds from the
  settings: `● on track  ● behind by more than 10 pp  ● behind by more than 25 pp or overdue  ● not assessable  (…) = previous week`.

Light colours: green `#2E9E44`, orange `#F0A202`, red `#D62828`, grey `#A0A0A0`.

### 8.3 Chart (`pptx/chart.py`)

Parts per chart: `ppt/charts/chartN.xml` (`c:chartSpace` → `c:lineChart`, `c:dispBlanksAs val="gap"`,
cached `c:strCache`/`c:numCache`), `ppt/embeddings/Microsoft_Excel_Worksheet_N.xlsx`
(written by `xlsx_write.py`), rels slide → chart → embedding, and content types
`application/vnd.openxmlformats-officedocument.drawingml.chart+xml`. On the slide:
`p:graphicFrame` with `c:chart r:id`.

## 9. CLI

```
python statusreport.pyz --init Tasks.xlsx          # create the empty input template
python statusreport.pyz Tasks.xlsx                 # weekly report for today
python statusreport.pyz Tasks.xlsx --date 2026-11-06 --template Customer.pptx
python statusreport.pyz Tasks.xlsx --polarion export.xlsx --no-save
```

Exit codes: `0` OK, `1` OK with row errors (rows skipped), `2` structural error / no output.
An unexpected exception is caught in `main()`, printed with its traceback and ends with
exit code 2, so the window of `make_report.cmd` shows what happened.

`make_report.cmd` uses `py -3` (or `python` if there is no launcher) to run
`%~dp0statusreport.pyz` with all arguments. It prints Done / rows skipped / nothing
written depending on the exit code, and keeps the window open (`pause`, skipped when
`SWSR_NO_PAUSE` is set, for tests).

Build (`tools/build_pyz.py`): the package is copied into a temp folder without
`__pycache__`, a root `__main__.py` calls `statusreport.__main__.main`, and
`zipapp.create_archive(compressed=True)` writes the archive. `base.pptx` is read with
`pkgutil.get_data`, which works from inside the zip.

## 10. Test strategy

| Level | What | How |
|---|---|---|
| Unit | parse, calc, history, merge | `unittest`, table-driven cases incl. thresholds at their limits |
| xlsx reader | files saved by Excel, by LibreOffice, by Polarion; written by our writer | fixtures in `tests/fixtures/` |
| pptx structure | rels, content types, unique IDs, valid XML, layout references | `opc_lint` in tests |
| Visual / compatibility | opens without repair prompt, looks right | manual check in PowerPoint (365) at milestones M5, M6, M7 |
| End to end | example Excel + 3 weeks history → deck | `tests/test_e2e.py`, deck kept as artefact |

## 11. Risks

| Risk | Mitigation |
|---|---|
| PowerPoint reports "repair" because of a small XML error | D-2 (one tested code path), `opc_lint`, early manual check (task T-5.4) |
| Customer template has unusual layouts (no body placeholder, odd names) | Configurable layout names, fallback content area, warning instead of abort |
| Polarion export varies (column names, `.xls`, merged header cells) | Mapping sheet, clear errors, real sample export (OP-1) before M2 is done |
| Company laptop has an old Python | Target 3.9: no `match`, union types only in annotations with `from __future__ import annotations`; test with 3.9 |
| Many tasks per version | Table split across slides (R-OUT-1), text shortening |
