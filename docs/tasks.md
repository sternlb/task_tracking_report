# Implementation plan — SW Status Report

Status: **draft v0.1** (2026-11-10) · Based on [requirements.md](requirements.md) and [design.md](design.md)

Each milestone ends with something you can run and check. Task IDs reference the
requirements they fulfil. `[ ]` open, `[x]` done.

## M0 — Spec sign-off

- [ ] T-0.1 Leo reviews requirements.md and design.md, and the open points are decided or deferred
- [ ] T-0.2 Get an anonymised Polarion export (OP-1)

## M1 — Foundation (risk first: can we write a pptx PowerPoint accepts?)

- [x] T-1.1 Repo skeleton (no packaging metadata needed), `python -m unittest` runs, `.gitignore`
- [x] T-1.2 `pptx/opc.py`: read/write a package, rels, content types
- [x] T-1.3 `tools/build_base_pptx.py`: neutral 16:9 `base.pptx` (master, layouts
      `Title Slide` + `Title Only` + `Title and Content`, theme) — D-2
- [x] T-1.4 `tools/opc_lint.py` + tests
- [x] T-1.5 **Spike:** one title slide + one table + one line chart in `base.pptx`, open in
      PowerPoint without a repair prompt. Go/no-go for the whole approach.
- [x] T-1.6 `.github/copilot-instructions.md` so GitHub Copilot understands the project

**M1 result (2026-11-10): GO.** `tools/spike.py` builds a deck from `base.pptx`.
PowerPoint 365 opens it in a normal window without a repair prompt. The chart has 3
series and is editable, the table is 8×5, and the shape list via `tools/ppt_check.ps1`
is as expected. 16 tests pass on Python 3.11 and 3.13 (3.9 is not installed here, see T-7.3).

Lessons:
- PowerPoint is *tolerant*. A copy with an unknown element and one with a missing
  table grid column both opened without a prompt (the second was silently "repaired"
  by dropping a column). Neither automation nor a normal open reliably shows small
  schema mistakes, so `lint.py` and the visual PNG check stay important, and element
  order has to be right by construction.
- Table row heights are minimums. The legend is anchored to the bottom of the content
  area (design 8.2 updated).
- Changed on 2026-11-10: no Owner/Due columns on the customer slides; the header is
  `Progress (prev. week)` (R-OUT-3, R-OUT-4).

## M2 — Excel in (R-IN-*)

- [x] T-2.1 `xlsx_read.py` incl. shared strings, inline strings, dates, 1904, empty cells
- [x] T-2.2 `xlsx_write.py` (styles, column widths, frozen header, drop-down validation)
- [x] T-2.3 `--init`: input template with Settings, Tasks, Import, Mapping and an example row (R-IN-1…5)
- [x] T-2.4 `parse.py`: dates, durations, percent, lights, natural sort (R-IN-6, R-IN-8)
- [x] T-2.5 `loader.py`: settings, mapping, merge by ID, validation and issues (R-IN-7, R-OP-3)
- [x] T-2.6 Fixtures: template filled in Excel, Polarion export, broken files
      (`tools/make_fixtures.ps1`; the Polarion export is **synthetic** until OP-1 delivers a real one)
- [ ] T-2.7 Check the mapping against a real Polarion export (OP-1)

**M2 result (2026-11-10):**
- `python -m statusreport --init X.xlsx` creates the input workbook. It opens in Excel
  without a repair prompt, and the drop-down and the frozen header survive Excel's
  open/save round trip.
- `python -m statusreport X.xlsx --check [--polarion export.xlsx] [--date …]` reads and
  validates the input and lists errors and warnings with sheet and row. Exit codes:
  0 OK, 1 row errors, 2 not usable.
- Fixtures saved by real Excel cover shared strings, Excel styles, dates as date cells
  and as text, a percent cell, the 1904 date system and an old `.xls` (rejected with a
  hint). 43 tests pass on Python 3.11 and 3.13.

Lessons:
- Excel is strict, unlike PowerPoint: a repair prompt there would have been visible.
  None appeared.
- Added on the way: R-IN-9 (`#` rows ignored), R-IN-10 (header row detection in exports),
  `--check` (R-OP-1).

## M3 — Calculation (R-CA-*)

- [x] T-3.1 Task progress with the precedence rules (R-CA-1)
- [x] T-3.2 Expected % and traffic light incl. limits and overdue (R-CA-2, R-CA-3)
- [x] T-3.3 Version aggregation and version light (R-CA-4, R-CA-5)
- [x] T-3.4 Table-driven tests for all rules
- [x] T-3.5 OP-5 decided and built: empty Start = Due − estimate in working days (R-CA-6)
- [x] T-3.6 OP-6 decided and built: empty imported Time Spent = 0 h (R-IN-11)

**M3 result (2026-11-10):**
- `calc.py` implements R-CA-1 … R-CA-5, and `--check` prints the computed values per
  version and task.
- 58 tests pass on Python 3.11 and 3.13. They are table-driven, include every threshold
  limit and float noise, and run end to end on the Excel-made fixture.
- Clarified in the requirements: an overdue task with unknown progress is red; the
  version light compares only tasks with both progress and plan; overrides do not change
  the version light.
- Found while checking the Polarion fixture: without a Start date every imported task is
  grey (OP-5), and an empty Polarion `Time Spent` gives no progress (OP-6).

## M4 — History (R-HI-*)

- [x] T-4.1 Write snapshots atomically, replace for the same week (R-HI-1, R-HI-2)
- [x] T-4.2 Previous week with gaps, trend series (R-HI-3)
- [x] T-4.3 `--no-save` (R-HI-4)

**M4 result (2026-11-10):**
- `history.py` writes one JSON snapshot per ISO week next to the input workbook
  (temp file + rename), and a second run in the same week replaces it.
- "Previous week" is the latest earlier snapshot, so skipped weeks are fine.
- The trend data covers the last N weeks that have a snapshot. The in-memory current
  week replaces a stored one, and snapshots *after* the report week are ignored (for
  runs with an earlier `--date`).
- A normal run stores the snapshot; `--check` and `--no-save` write nothing.
- `--check` now shows the previous values in brackets as on the slide (`(50 %)`,
  `(new)`, `(-)`), plus light changes (`orange (was green)`).
- Broken snapshot files become warnings and are skipped.
- 78 tests pass on Python 3.11 and 3.13, including a simulated run over three weeks
  through the CLI.

## M5 — Deck without template (R-OUT-1…7, R-OUT-9)

- [x] T-5.1 `report.py`: report model (slides, rows, values incl. previous week)
- [x] T-5.2 Title slide and version tables with row splitting, highlighting, lights, legend
- [x] T-5.3 Trend chart with embedded workbook
- [x] T-5.4 **Manual check in PowerPoint** with example data covering 3 weeks
      (`tools/make_demo.py`: 22 tasks, 3 versions, 7 weeks; rendered by PowerPoint 365)
- [x] T-5.5 Output file name and overwrite protection (R-OUT-9)

**M5 result (2026-11-10):**
- A normal run writes `<Project> Status <YYYY>-W<WW>.pptx` next to the workbook, then
  stores the snapshot. If the deck cannot be written, the week is not stored either.
- The deck has a title slide, the trend chart, and per version a table (continued
  slides repeat the total row) with the legend at the bottom.
- An existing deck gets a ` (2)` suffix unless `--force` is given. `--out` and
  `--template` are available, and the Settings `Template file` is resolved relative to
  the workbook.
- No tasks → nothing written, exit 2. A deck that is open in PowerPoint gives a clear
  message.
- PowerPoint 365 opens the demo deck and renders all 6 slides as designed: trend over
  7 weeks, 1.2.0 on 2 slides, `(new)`, light changes, grey rows, the shortened title.
- 95 tests pass on Python 3.11 and 3.13; the demo runs end to end as a test.

Seen in the check, acceptable for now:
- Comments longer than about 38 characters wrap to a second line; the row grows and
  the page still fits with 12 rows.
- The version light can be green while several tasks are orange, because it compares
  weighted totals (R-CA-5). That works as specified, but it is worth explaining in the
  first customer meeting.

## M6 — Customer template (R-OUT-8)

- [x] T-6.1 Remove slides from a foreign template cleanly (incl. notes, charts, media)
- [x] T-6.2 Layouts by name, fallback, content area from the placeholder
- [x] T-6.3 Test with 2–3 public templates (different sizes, 4:3, without `Title Only`)
      — instead of downloads: 4 templates built by real PowerPoint (`tools/make_templates.ps1`)
- [ ] T-6.4 **As soon as the template arrives:** set the layout names, run a manual
      check, adjust colours/columns if needed (OP-2)

**M6 result (2026-11-10):**
- Test templates saved by PowerPoint 365 (German Office):
  - Office default 16:9, as `.pptx` and `.potx`, with slides, notes, a picture,
    sections and a custom show;
  - "Facet" in 4:3;
  - "Dividend" reduced to two renamed layouts ("Cover", "Content").
- Found and fixed with them:
  1. `.potx` was rejected → now accepted.
  2. German layout names made the fallback pick a random narrow layout → fallback by
     layout type.
  3. Tables overflowed on smaller content areas (4:3, large title bands) → table fitting
     (fewer rows, smaller fonts, shortening by column width).
  4. "Version total" wrapped in narrow ID columns → spans two columns.
  5. The template's slide numbers and footers were missing → copied onto content slides
     (new setting `Footer text`).
- Rendered in PowerPoint on all four templates plus the built-in base: the tables fit,
  the legend sits below the table, theme colours and fonts are taken over, and slide
  numbers appear.
- 109 tests pass on Python 3.11 and 3.13.
- Not covered by the generated templates: old charts on template slides (AddChart2 does
  not work without a window). Removing chart parts is tested with our own deck as the
  template.

## M7 — Delivery

- [x] T-7.1 CLI complete (R-OP-1), exit codes, console output
- [x] T-7.2 `tools/build_pyz.py` → `dist/statusreport.pyz`, `make_report.cmd` (R-OP-2)
- [x] T-7.3 Test with Python 3.9 and 3.12, performance with 500 tasks (NF-1)
      — done: performance (0.6 s for 500 tasks and 52 weeks), Python 3.11 and 3.13, a static
      3.9 check (grammar, no newer APIs, stdlib-only imports). A run on a real Python 3.9
      was **skipped by decision (2026-11-10)**: it would need a download. The static check
      stays the guard. If the target laptop turns out to have 3.9 or 3.10, run the test
      suite there once.
- [x] T-7.4 README as a user guide (NF-4): Excel, Polarion export view, weekly routine, template
- [ ] T-7.5 Acceptance: Leo runs a real week end to end

**M7 result (2026-11-10):**
- `python tools/build_pyz.py` builds `dist/sw-status-report/` (`statusreport.pyz` 57 KB,
  `make_report.cmd`, an empty `Tasks.xlsx`, `README.md`) and
  `dist/sw-status-report-0.9.0.zip`.
- Unpacked away from the sources, it runs with `python -I` on 3.11 and 3.13. The
  built-in design loads from inside the zipapp, and `make_report.cmd` finds `py`/`python`,
  passes the workbook on and reports the result. These are now also automated tests.
- CLI: the normal run shows one line per version; `-v`/`--verbose` and `--check` list
  every task. An unexpected error ends with a readable message (exit 2) instead of a
  bare traceback.
- README rewritten as a user guide: getting started, weekly routine, the workbook,
  calculation rules, Polarion, template, options, messages.
- Version 0.9.0; 1.0.0 after acceptance (T-7.5).
- 118 tests pass on Python 3.11 and 3.13.
- Lesson: PowerShell 5.1 `Set-Content -Encoding utf8` writes a BOM; a test now fails on
  that.

## After M7

- [x] T-8.1 German tutorial (offline HTML, real slides, traffic-light calculator) and an
      80-second video built with PowerPoint
- [x] T-8.2 Burndown slide over all tasks (R-OUT-10), version 0.9.1: open share of the
      effort in % (no hours for the customer), effort stored in the snapshots and derived
      for older ones; checked in PowerPoint on the base and on the 4:3 Facet template;
      131 tests pass

## Order and estimate

| Milestone | Effort (sessions) | Depends on |
|---|---|---|
| M1 | 1–2 | M0 |
| M2 | 1–2 | M1 (xlsx writer shares the zip/XML helpers) |
| M3 | 1 | M2 |
| M4 | 0.5 | M3 |
| M5 | 2 | M1, M4 |
| M6 | 1 (+ follow-up when the template arrives) | M5 |
| M7 | 1 | M6 |
