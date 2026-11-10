# Copilot instructions — SW Status Report

## What this project is

A command-line tool that turns a list of software work packages (an Excel file, typed
in by hand or exported from Polarion) into a **weekly PowerPoint status deck** for
customer meetings. The deck has a title slide, one "Progress over time" line chart
(one line per software version) and one status table per software version. Each task
row shows its progress in % with last week's value in brackets and a green, orange,
red or grey traffic light that says whether the task is on schedule.

## Hard rules (do not break these)

1. **Python standard library only, Python 3.9+.** No `pip install`, no third-party
   packages (no python-pptx, openpyxl, lxml, pandas), no Office/COM automation, no
   network access at runtime. Office files are read and written directly as Office
   Open XML with `zipfile` and `xml.etree.ElementTree`. If a feature seems to need a
   library, implement the small part that is needed instead.
2. **Python 3.9 syntax:** start each module with `from __future__ import annotations`.
   Do not use `match`, and use union types (`X | None`) only in annotations.
3. **The spec is the source of truth.** Read `docs/requirements.md` (what, with IDs such
   as `R-CA-3`), `docs/design.md` (how, decisions `D-1`…) and `docs/tasks.md` (plan,
   milestones `M0`…`M7`) before changing behaviour. If code and spec disagree, ask, or
   update the spec in the same change. Reference requirement IDs in docstrings and tests.
4. **Never modify the user's input Excel** (R-OP-4). Output: a new `.pptx`, plus one
   history snapshot per calendar week in `history/`.
5. **Deterministic output** (NF-2): same input gives the same bytes. Do not add
   timestamps or random IDs to package parts.
6. **The customer's PowerPoint template is unknown.** Never hard-code positions, fonts
   or colours of a specific template. Take layouts by name, the content area from the
   layout's placeholders, and colours from theme slots (`accent1`, `tx1`, `bg1`). Only
   the traffic-light colours are fixed RGB values.
7. Customer slides show **no Owner and no Due date** (R-OUT-3). Both exist only in the
   Excel input; Due feeds the traffic light.

## How the code is organised

```
statusreport/
  __main__.py        CLI: --init, --check (deck generation follows in M5)
  model.py           dataclasses Settings, Task, Issue, LoadResult
  labels.py          all user-visible texts: sheet names, column headers, help texts
  defaults.py        setting definitions and default Polarion column names
  parse.py           cell parsers: date, hours ("1w 2d 4h"), percent, light, text, natural_key
  xlsx_read.py       minimal .xlsx reader -> rows of Python values (Percent, date, str, float)
  loader.py          input workbook -> tasks + issues (settings, mapping, import merge by ID)
  calc.py            progress, expected %, traffic lights, version aggregation (R-CA-*)
  history.py         weekly JSON snapshots in history/, previous week, trend series (R-HI-*)
  report.py          presentation-neutral report model: slide texts, rows, "(50 %)"/"(new)"/"(–)"
  render.py          report model -> slides on a Deck (table styling, lights, legend, chart)
  output.py          output file name, " (2)" suffix instead of overwriting, template path
  template.py        the empty input workbook for --init
  xmlutil.py         namespaces, XML escaping, XML declaration
  xlsx_write.py      minimal .xlsx writer: styles, widths, frozen header, drop-downs
  pptx/
    opc.py           zip package: parts, relationships (.rels), [Content_Types].xml
    lint.py          structural self-check of a package (run before every save)
    deck.py          Deck: open a base/template .pptx, remove its slides, add slides on named layouts
    shapes.py        XML builders: Run/Para/Cell/Row/Box, placeholders, text boxes, tables, chart frames
    chart.py         native line chart part + embedded workbook ("Edit Data" works)
    base.pptx        built-in neutral 16:9 template (generated, do not edit by hand)
tools/
  build_base_pptx.py regenerates statusreport/pptx/base.pptx
  spike.py           M1 sample deck with every XML construct the real slides need
  opc_lint.py        CLI around pptx/lint.py
  ppt_check.ps1      dev only: opens a deck in installed PowerPoint, exports PNGs
  make_fixtures.ps1  dev only: creates tests/fixtures/*.xlsx with real Excel
  make_demo.py       demo input + 6 weeks of history (python tools/make_demo.py out/demo)
  make_templates.ps1 dev only: foreign test templates (German names, .potx, 4:3) with real PowerPoint
  build_pyz.py       delivery: dist/sw-status-report/ (statusreport.pyz, make_report.cmd, Tasks.xlsx, README, tutorial) + zip
  build_tutorial.py  docs/tutorial/tutorial.html + img/ -> one self-contained HTML (images as data URIs)
docs/tutorial/       German step-by-step tutorial; keep it in line with the tool when behaviour or texts change
make_report.cmd      Windows launcher shipped with the .pyz (drag the workbook onto it)
tests/               unittest only; fixtures/ holds files saved by real Excel (do not regenerate casually)
docs/                requirements.md, design.md, tasks.md
```

Data flow of a run: `loader.load` → `calc.compute` → `history` (previous week, trend) →
`report.build_report` → `render.render` onto `Deck.default()` or the customer template →
`deck.save` → `history.save`. Keep slide content decisions in `report.py` and the
PowerPoint details in `render.py`/`pptx/`.

### Input handling in short

- The loader never raises for problems in a single row. It records an `Issue`
  (`error` = row skipped, `warning` = kept) with sheet and Excel row number. Only
  structural problems (missing sheet or mapped column) raise `InputError`, and unreadable
  files raise `XlsxError`.
- Parsers in `parse.py` raise `ValueError` with a message written for the end user.
- A filled `Tasks` cell wins over the Polarion import for the same ID (R-IN-7). Rows with
  an ID starting with `#` are ignored (R-IN-9).
- An empty mapped Time Spent in the import is read as `0.0` h (R-IN-11); in a
  `Tasks`-only row an empty Spent stays `None` (unknown).
- Plain numbers in Progress are percentage points (65 = 65 %). Only cells with a
  percent format arrive as `Percent(0.65)`.

### Key ideas

- **One code path for both modes (D-2).** A deck is always built by opening a base
  `.pptx`, removing all its slides (`Deck._strip_slides`, then unreachable parts are
  garbage-collected) and adding new slides. Without a customer template the built-in
  `base.pptx` is used.
- **Part names** have no leading slash (`ppt/slides/slide1.xml`). Use `Package.add_part`,
  `add_rel` and `next_name`; never write zip entries directly.
- **XML is written as text** with f-strings. Always pass user text through `esc()`.
  Keep the schema's **element order** (e.g. in `a:rPr` the fill comes before `a:latin`;
  in `a:tcPr` the borders come before the fill); PowerPoint is strict about it.
- **Units:** EMU (914400 per inch, 12700 per point). Use `shapes.inch()`. Font sizes in
  `Run.size` are in points.
- **Colours** in `shapes.color_xml`: `"#RRGGBB"` fixed, `"accent1"` theme slot,
  `"accent1/20"` 20 % tint of a theme slot.
- Table row heights are minimums; rows grow with their text. Do not position things
  below a table by adding up row heights.
- Templates (M6): never rely on layout names alone, since they are localised ("Nur Titel").
  `Deck.layout(name, role)` falls back to the layout type. Everything on a slide has
  to fit `Deck.content_area(layout)`, which already excludes the footer zone; the
  table sizes itself in `render.metrics`. Test any change against
  `tests/fixtures/templates/*` (`tests/test_templates.py`), not only against the base.

## Domain rules in short (full text in docs/requirements.md §4.2)

- Task progress: manual % → else done state = 100 % → else spent/estimate hours,
  capped at 99 % → else no value (`–`).
- Empty Start = Due minus the estimate in working days (`calc.derive_start`, R-CA-6);
  `Task.start_derived` marks it. A Start entered by the user is never replaced.
- Expected % = elapsed share between Start and Due (calendar days), limited to 0–100.
- Light: override → 100 % is green → past Due and not done is red (also with unknown
  progress) → missing data is grey → gap = expected − progress: ≤ 10 pp green,
  ≤ 25 pp orange, else red (thresholds come from the Settings sheet; limits included).
- Version progress = average weighted by estimated hours (plain average if an estimate
  is missing). The version light compares only tasks with both progress and plan;
  task overrides do not affect it.
- Lights are the strings `green`, `orange`, `red`, `grey` (constants in `calc.py`).
- "Previous week" = the latest snapshot in `history/` before the report week.
- Week keys are ISO weeks (`2026-W46`, via `history.week_key`); string order = time order.
  The trend uses the weeks that have snapshots, not every calendar week.
- `--check` and `--no-save` must never write to `history/`; tests copy fixtures to a
  temp folder before a normal run.

## Workflow

- Run the tests: `python -m unittest discover -s tests`
- Create an input workbook: `python -m statusreport --init out/Tasks.xlsx`
- Validate an input: `python -m statusreport tests/fixtures/excel_with_import.xlsx --check`
- Full demo deck: `python tools/make_demo.py out/demo` then `python -m statusreport out/demo/Tasks.xlsx`
- Look at it: `powershell -File tools\ppt_check.ps1 "out\demo\Project Falcon Status 2026-W46.pptx" -PngDir out\png`
- After changing the base design: `python tools/build_base_pptx.py`
- Sample deck: `python tools/spike.py`, then `python tools/opc_lint.py out/spike.pptx`
- Build the delivery: `python tools/build_pyz.py` (output in `dist/`, not committed)
- `tests/test_delivery.py` guards the hard rules: Python 3.9 grammar, no newer APIs,
  standard-library imports only, the .pyz runs isolated (`python -I`), performance.
  Keep it green.
- Save source files as UTF-8 **without BOM** (PowerShell 5.1 `Set-Content -Encoding utf8`
  adds one; the tests catch it).
- When finishing a task from `docs/tasks.md`, tick its checkbox in the same commit.
- All user-visible text is English. From M5 on it belongs in `labels.py`, not inline.
- Keep functions small and typed; dataclasses for data; no global state.
