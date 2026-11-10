# Requirements — SW Status Report

Status: **draft v0.1** (2026-11-10) · Phase: spec, not implemented

## 1. Purpose

A tool for weekly customer meetings that turns a list of software work packages
(Excel, typed in by hand or exported from Polarion) into a PowerPoint deck. The deck
shows how far each work package and each software version has come, how that changed
since last week, and whether each task is on schedule.

## 2. Users and context

| Role | Does |
|---|---|
| Project lead (report author) | Keeps the task list up to date, runs the tool once a week, presents the deck |
| Customer | Sees the deck in the meeting, never touches the tool |

The tool runs on a locked-down company laptop. The customer's PowerPoint template
arrives later and must not be needed to build the tool.

## 3. Constraints

- **C-1** The tool runs with only a Python 3.9+ interpreter. It uses the standard
  library only: no pip, no third-party packages, no Office automation, no other
  programs, no network access.
- **C-2** No installation step. Copying a folder (or a single file) is enough.
- **C-3** Input is `.xlsx`. Output is `.pptx` that opens without repair prompts in
  PowerPoint 2016 or later and in Microsoft 365.
- **C-4** Language of the slides and the Excel template: English. All visible labels sit
  in one place so that a second language can be added later.
- **C-5** Task data never leaves the machine.

## 4. Functional requirements

IDs are stable. Priority: **M** = must for v1, **S** = should, **C** = could.

### 4.1 Excel input template

- **R-IN-1 (M)** The tool shall create an empty Excel input template on request
  (`--init`). The template contains the sheets `Settings`, `Tasks`, `Import` and `Mapping`,
  plus one example row.
- **R-IN-2 (M)** The `Tasks` sheet shall hold one row per work package with these columns:

  | Column | Required | Meaning |
  |---|---|---|
  | ID | yes | Unique key, e.g. Polarion work item ID |
  | Title | yes | Short task name |
  | Version | yes | Software version / release the task belongs to |
  | Owner | no | Responsible person |
  | Start | for the traffic light | Planned start date |
  | Due | for the traffic light | Planned finish date |
  | Estimate (h) | no | Estimated effort in hours |
  | Spent (h) | no | Hours used so far |
  | Progress (%) | no | Manually entered completion, 0–100 |
  | State | no | Work item state, e.g. `open`, `in progress`, `done` |
  | Light override | no | `Green`, `Orange` or `Red` (drop-down) |
  | Comment | no | Short remark shown on the slide |

- **R-IN-3 (M)** The `Settings` sheet shall hold: project name, customer name, report
  date (empty = today), the orange and red thresholds (default 10 and 25 percentage
  points), hours per day for duration parsing (default 8), the states that count as done
  (default `done, closed, resolved, verified`), the maximum number of table rows per
  slide (default 12), the number of weeks in the trend chart (default 12), and the
  version order (empty = alphabetical).
- **R-IN-4 (M)** The `Import` sheet shall accept a Polarion Excel export pasted as-is
  (header row plus data). As an alternative, the tool shall accept the path to a separate
  Polarion export file on the command line.
- **R-IN-5 (M)** The `Mapping` sheet shall assign each field from R-IN-2 to a column
  header of the Polarion export. Unmapped fields stay empty.
- **R-IN-6 (M)** Durations shall be read both as plain numbers (hours) and in the
  Polarion format (`1w 2d 4h 30m`, `3d`, `12h`). 1 d = hours-per-day setting,
  1 w = 5 d.
- **R-IN-7 (M)** When both `Import` and `Tasks` contain a row with the same ID, a
  non-empty cell in `Tasks` shall win over the import. This lets the user add comments,
  overrides or a manual % on top of a Polarion export. `Tasks` rows without a match in
  the import are added as additional tasks.
- **R-IN-8 (M)** Dates shall be read as Excel date cells and as text in the formats
  `YYYY-MM-DD`, `DD.MM.YYYY` and `DD/MM/YYYY`.

### 4.2 Calculation

- **R-CA-1 (M)** Progress % of a task, first rule that applies:
  1. `Progress (%)` filled in → that value
  2. `State` is a done state → 100 %
  3. `Estimate (h)` > 0 and `Spent (h)` given → spent / estimate, at most 99 %
     (a task is only 100 % when it is done or set to 100 manually)
  4. otherwise → no value (shown as `–`)
- **R-CA-2 (M)** Expected % of a task on the report date = elapsed share of the time
  between Start and Due, limited to 0–100 % (calendar days).
- **R-CA-3 (M)** Traffic light of a task, first rule that applies:
  1. `Light override` filled in → that colour
  2. Progress = 100 % → green
  3. Report date is after Due and progress < 100 % → red
  4. Start, Due or progress missing → grey ("not assessable")
  5. Gap = expected − progress. Gap ≤ orange threshold → green;
     gap ≤ red threshold → orange; else → red
- **R-CA-4 (M)** Progress % of a version = average of its tasks' progress weighted by
  Estimate (h). If any task with a progress value has no estimate, the plain average is
  used. Tasks without a progress value are left out and counted as a warning.
- **R-CA-5 (S)** Traffic light of a version: the same rules as R-CA-3, applied to the
  version's progress against the version's expected % (weighted the same way, with the
  earliest Start and latest Due of its tasks).

### 4.3 History

- **R-HI-1 (M)** Every run shall save a snapshot for the ISO calendar week of the report
  date: per task its progress, light and version, per version its progress and light.
- **R-HI-2 (M)** Snapshots shall be stored in a folder `history/` next to the input
  Excel, one human-readable file per week (`2026-W46.json`). Running the tool again in
  the same week replaces that week's snapshot.
- **R-HI-3 (M)** "Previous week" means the latest snapshot before the report week, even
  if one or more weeks were skipped.
- **R-HI-4 (S)** A `--no-save` option shall create the deck without writing a snapshot
  (for trial runs).

### 4.4 PowerPoint output

- **R-OUT-1 (M)** The deck shall contain, in this order:
  1. Title slide: project, customer, calendar week, report date
  2. One "Progress over time" slide
  3. Per software version: one status table (continued on more slides if it has more
     rows than the per-slide limit)
- **R-OUT-2 (M)** The "Progress over time" slide shall show a native, editable
  PowerPoint line chart: x-axis = calendar weeks (the last N weeks with snapshots),
  y-axis = 0–100 %, one line per version.
- **R-OUT-3 (M)** Each version table shall contain per task: ID, Title, Progress, Light,
  Comment. A summary row at the top shows the version's progress and light. Owner and
  Due are input only (Due feeds the traffic light) and are **not** shown on the customer
  slides (decided 2026-11-10).
- **R-OUT-4 (M)** The progress cell shall show the current value highlighted (bold, larger
  font) followed by the previous week's value in brackets in a smaller grey font, e.g.
  **65 %** (50 %). Tasks without a previous value show `(new)`, and the first report ever
  shows `(–)`. The column header explains the bracket: `Progress (prev. week)`.
- **R-OUT-5 (M)** The light shall be shown as a green, orange, red or grey circle. A
  short legend on each table slide explains the rule.
- **R-OUT-6 (S)** A change in the light since last week (e.g. green → red) shall be
  visible, e.g. by a small arrow or marker.
- **R-OUT-7 (M)** Without a customer template the tool shall produce a clean, neutral
  16:9 deck.
- **R-OUT-8 (M)** With a template (`--template file.pptx`) the tool shall use its slide
  master, layouts, theme colours and fonts. The layouts used are chosen by name and
  can be configured. The template's existing slides are not copied.
- **R-OUT-9 (M)** Default output file name: `<Project> Status <YYYY>-W<WW>.pptx` next to
  the input Excel. An existing file is replaced only with `--force`; otherwise a suffix is
  added.

### 4.5 Operation

- **R-OP-1 (M)** Command line: `python statusreport.pyz <input.xlsx> [options]`
  (`--init`, `--template`, `--polarion`, `--date`, `--out`, `--no-save`, `--force`).
- **R-OP-2 (S)** Windows convenience: dragging the Excel onto `make_report.cmd` creates
  the deck.
- **R-OP-3 (M)** Data problems (missing ID, duplicate ID, date that cannot be read,
  % outside 0–100, unknown version in the version order) shall be listed in the console
  with sheet, row and reason. Rows with errors are skipped and the report is still
  created. Structural errors (sheet or required column missing) stop the run with a
  clear message.
- **R-OP-4 (M)** The tool shall never modify the input Excel.

## 5. Non-functional requirements

- **NF-1** Runtime under 5 s for 500 tasks and 52 weeks of history.
- **NF-2** Same inputs produce the same deck (deterministic; only the timestamp differs).
- **NF-3** Automated tests with `unittest` (stdlib) for the parser, the calculations,
  the history and the structure of the generated XML.
- **NF-4** User guide (README) for a non-developer: preparing the Excel, the Polarion
  export, the weekly run and the template.

## 6. Out of scope (v1)

- Editing the deck after generation in a round-trip
- Direct Polarion API access (only Excel export)
- Gantt charts, burn-down charts, resource or cost views
- Public holidays and working-day calendars (calendar days only)
- PDF export (PowerPoint does this itself)

## 7. Open points

- **OP-1** A real Polarion export (anonymised) for testing the mapping and the duration
  formats.
- **OP-2** The customer PowerPoint template: layout names, content area, fonts.
- **OP-3** Is `Remaining Estimate` from Polarion more reliable than `Spent / Estimate`?
  If so, progress = spent / (spent + remaining) as an alternative rule.
- **OP-4** Should "not assessable" (grey) tasks appear on the customer slide or only in
  the console?
