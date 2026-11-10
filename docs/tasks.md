# Implementation plan — SW Status Report

Status: **draft v0.1** (2026-11-10) · Based on [requirements.md](requirements.md) and [design.md](design.md)

Each milestone ends with something you can run and check. Task IDs reference the
requirements they fulfil. `[ ]` open, `[x]` done.

## M0 — Spec sign-off

- [ ] T-0.1 Leo reviews requirements.md and design.md, and the open points are decided or deferred
- [ ] T-0.2 Get an anonymised Polarion export (OP-1)

## M1 — Foundation (risk first: can we write a pptx PowerPoint accepts?)

- [ ] T-1.1 Repo skeleton (no packaging metadata needed), `python -m unittest` runs, `.gitignore`
- [ ] T-1.2 `pptx/opc.py`: read/write a package, rels, content types
- [ ] T-1.3 `tools/build_base_pptx.py`: neutral 16:9 `base.pptx` (master, layouts
      `Title Slide` + `Title Only` + `Title and Content`, theme) — D-2
- [ ] T-1.4 `tools/opc_lint.py` + tests
- [ ] T-1.5 **Spike:** one title slide + one table + one line chart in `base.pptx`, open in
      PowerPoint without a repair prompt. Go/no-go for the whole approach.

## M2 — Excel in (R-IN-*)

- [ ] T-2.1 `xlsx_read.py` incl. shared strings, inline strings, dates, 1904, empty cells
- [ ] T-2.2 `xlsx_write.py` (styles, column widths, frozen header, drop-down validation)
- [ ] T-2.3 `--init`: input template with Settings, Tasks, Import, Mapping and an example row (R-IN-1…5)
- [ ] T-2.4 `parse.py`: dates, durations, percent, lights, natural sort (R-IN-6, R-IN-8)
- [ ] T-2.5 `loader.py`: settings, mapping, merge by ID, validation and issues (R-IN-7, R-OP-3)
- [ ] T-2.6 Fixtures: template filled in Excel, Polarion export, broken files

## M3 — Calculation (R-CA-*)

- [ ] T-3.1 Task progress with the precedence rules (R-CA-1)
- [ ] T-3.2 Expected % and traffic light incl. limits and overdue (R-CA-2, R-CA-3)
- [ ] T-3.3 Version aggregation and version light (R-CA-4, R-CA-5)
- [ ] T-3.4 Table-driven tests for all rules

## M4 — History (R-HI-*)

- [ ] T-4.1 Write snapshots atomically, replace for the same week (R-HI-1, R-HI-2)
- [ ] T-4.2 Previous week with gaps, trend series (R-HI-3)
- [ ] T-4.3 `--no-save` (R-HI-4)

## M5 — Deck without template (R-OUT-1…7, R-OUT-9)

- [ ] T-5.1 `report.py`: report model (slides, rows, values incl. previous week)
- [ ] T-5.2 Title slide and version tables with row splitting, highlighting, lights, legend
- [ ] T-5.3 Trend chart with embedded workbook
- [ ] T-5.4 **Manual check in PowerPoint** with example data covering 3 weeks
- [ ] T-5.5 Output file name and overwrite protection (R-OUT-9)

## M6 — Customer template (R-OUT-8)

- [ ] T-6.1 Remove slides from a foreign template cleanly (incl. notes, charts, media)
- [ ] T-6.2 Layouts by name, fallback, content area from the placeholder
- [ ] T-6.3 Test with 2–3 public templates (different sizes, 4:3, without `Title Only`)
- [ ] T-6.4 **As soon as the template arrives:** set the layout names, run a manual
      check, adjust colours/columns if needed (OP-2)

## M7 — Delivery

- [ ] T-7.1 CLI complete (R-OP-1), exit codes, console output
- [ ] T-7.2 `tools/build_pyz.py` → `dist/statusreport.pyz`, `make_report.cmd` (R-OP-2)
- [ ] T-7.3 Test with Python 3.9 and 3.12, performance with 500 tasks (NF-1)
- [ ] T-7.4 README as a user guide (NF-4): Excel, Polarion export view, weekly routine, template
- [ ] T-7.5 Acceptance: Leo runs a real week end to end

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
