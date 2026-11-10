# SW Status Report

Turns your list of software work packages into a **weekly PowerPoint status deck** for
the customer meeting:

- **Title slide** with project, customer, calendar week and date
- **Progress over time**: one line per software version over the last weeks
- **One table per software version**: every task with its progress in %, last week's
  value in brackets, e.g. **65 %** (50 %), and a traffic light that shows whether it is on
  schedule: ● green, ● orange, ● red, ● grey (not assessable)

The task list is an Excel file you fill in by hand, or a Polarion export you paste in.
The deck can use your customer's PowerPoint template.

Version 0.9.0 · English · runs with plain Python, nothing to install.

---

## 1. What you need

- **Windows** with **Python 3.9 or newer**. Check in a command prompt: `py --version`
  (or `python --version`). There is nothing else to install: no add-ons, no internet, no
  admin rights.
- **PowerPoint** to look at the result, and **Excel** to edit the task list.

## 2. Getting started

1. Unzip `sw-status-report-0.9.0.zip` somewhere, e.g. `C:\Tools\sw-status-report`. It
   contains `statusreport.pyz` (the tool), `make_report.cmd`, `Tasks.xlsx` (an empty
   input workbook) and this guide.
2. Copy `Tasks.xlsx` into your project folder and open it. The sheet **Read me**
   explains it in short.
3. In **Settings**, enter the project and customer names.
4. In **Tasks**, enter your work packages, or paste a Polarion export into **Import**
   (see section 5).
5. Save, close Excel, and **drag `Tasks.xlsx` onto `make_report.cmd`**. A window shows
   what was read and ends with "Done."
6. The deck is next to your workbook, e.g. `Project Falcon Status 2026-W46.pptx`.

The first deck has no previous-week values yet (`(–)`), and the chart starts with a
single week. Both build up from the second week on.

## 3. The weekly routine

1. Update the task list (or paste the new Polarion export into **Import**).
2. Drag the workbook onto `make_report.cmd`.
3. Check the deck, then present it.

Each run stores the week's values in a folder **`history`** next to your workbook (one
small file per calendar week, e.g. `2026-W46.json`). This is where last week's values
and the chart come from. **Keep this folder**: if it is deleted, the brackets and the
trend start over. Running again in the same week replaces that week's values. If you
skip a week, the brackets show the most recent earlier week.

An existing deck is never overwritten: a second run in the same week creates
`… (2).pptx`.

## 4. The input workbook

### Sheet "Tasks": one row per work package

| Column | Needed | Meaning |
|---|---|---|
| ID | yes | Unique key, e.g. the Polarion ID. A row whose ID starts with `#` is ignored. |
| Title | yes | Task name shown on the slide |
| Version | yes | Software version the task belongs to. Keep the column formatted as text, so 1.10 stays 1.10. |
| Owner | no | Not shown on the slides |
| Start | for the light | Planned start. Empty = Due minus Estimate (working days). |
| Due | for the light | Planned finish. Not shown on the slides. |
| Estimate (h) | no | Hours, or a duration such as `2d 4h` (1 d = 8 h, 1 w = 5 d) |
| Spent (h) | no | Hours used so far, same format |
| Progress (%) | no | Completion 0–100; wins over the calculation from hours |
| State | no | Work item state; `done`, `closed`, `resolved` and `verified` count as 100 % |
| Light override | no | Green / Orange / Red: sets the light by hand |
| Comment | no | Short remark shown on the slide |

Dates can be Excel dates or text like `2026-11-27` or `27.11.2026`.

### How the values are calculated

- **Progress of a task**, using the first rule that applies: the value in Progress (%) →
  100 % if the State is a done state → Spent ÷ Estimate (at most 99 % until the task is
  done) → no value (`–`).
- **Expected progress** = the share of the time from Start to Due that has passed on the
  report date. Example: 10 of 25 days gone = 40 % expected.
- **Traffic light**, using the first rule that applies: override → 100 % is green → past
  Due and not done is red → no Due or no progress is grey → otherwise it compares how
  far the task is behind its expected progress: up to 10 percentage points green, up
  to 25 orange, more red. The thresholds can be set in Settings.
- **Version total** = the average of its tasks weighted by estimated hours (a plain
  average if an estimate is missing). The version light compares the weighted totals,
  so a version can be green while single tasks are orange.

### Sheet "Settings"

| Setting | Default | Meaning |
|---|---|---|
| Project, Customer | | Shown on the title slide; the project is also in the file name |
| Report date | today | Sets the calendar week and the expected progress |
| Orange / Red threshold (pp) | 10 / 25 | Percentage points behind schedule |
| Hours per day | 8 | For durations like `2d` |
| Done states | done, closed, resolved, verified | States that count as 100 % |
| Rows per slide | 12 | Maximum; fewer if the template leaves less room |
| Trend weeks | 12 | Weeks shown in the chart |
| Version order | natural (1.2 before 1.10) | Order of the version slides, comma-separated |
| Template file | built-in design | Customer template (.pptx or .potx), relative to the workbook |
| Title layout / Content layout | Title Slide / Title Only | Layout names in the template |
| Footer text | | Footer on chart and table slides, if the template has a footer field |

## 5. Using a Polarion export

1. In Polarion, export the work items to **Excel (.xlsx)** with the columns you need
   (at least ID, title, version / planned in; ideally due date, initial estimate, time
   spent and status).
2. Paste the whole table including its header row into the sheet **Import**. As an
   alternative, keep the export as a separate file and run it from the command line:
   `py statusreport.pyz Tasks.xlsx --polarion export.xlsx`.
3. In **Mapping**, check that the column names match your export. They are pre-filled
   with typical Polarion names: ID, Title, Planned In, Assignee, Due Date, Initial
   Estimate, Time Spent, Status. Clear a name you do not export.

Details:
- A filled cell in **Tasks** wins over the import for the same ID. Use it for comments,
  overrides or a manual %; leave the other columns empty.
- Polarion has no standard start date, so an empty Start is derived from Due minus
  the estimate.
- An empty Time Spent in the export counts as 0 h.
- Old `.xls` exports are not supported: open them in Excel and save as `.xlsx`.

## 6. Using the customer's PowerPoint template

1. Put the template (`.pptx` or `.potx`) next to the workbook and enter its file name in
   Settings → **Template file**.
2. Run as usual. The deck uses the template's master, colours, fonts and slide numbers.
   Old slides in the template are not copied.
3. If the window shows `warning template: layout "Title Only" not found, using …`, the
   tool picked a suitable layout itself. To choose one yourself, copy a name from the
   list in the warning into **Title layout** / **Content layout**.
4. Optional: **Footer text**, e.g. `ACME confidential`.

## 7. Command line (optional)

`make_report.cmd` passes extra options on to the tool, and you can also call the tool
directly: `py statusreport.pyz <workbook> [options]`.

| Option | Effect |
|---|---|
| `--check` | Only read, calculate and list every task and problem; writes nothing |
| `-v`, `--verbose` | List every task with its values during a normal run |
| `--no-save` | Trial run: writes the deck but does not store the week in `history` |
| `--date 2026-11-06` | Report for another date (overrides Settings) |
| `--polarion export.xlsx` | Use a separate Polarion export instead of the Import sheet |
| `--template customer.pptx` | Use this template (overrides Settings) |
| `--out deck.pptx` | Output file name |
| `--force` | Overwrite an existing deck instead of creating `… (2).pptx` |
| `--init new.xlsx` | Create a new empty input workbook |

Exit codes: `0` done · `1` done, but rows with errors were skipped · `2` nothing written.

## 8. Messages and what to do

| Message | Meaning / fix |
|---|---|
| `error Tasks row 7: PRJ-105: Due: "31.02.2026" is not a date` | Fix the cell; the row is skipped until then, and the deck is still created |
| `duplicate ID … row skipped` | The ID appears twice; the first row is used |
| `warning … no progress …; not counted in version …` | The task has no %, no done state and no hours; it shows `–` and a grey light |
| `sheet "Tasks" is missing` | This is not the input workbook; start from `Tasks.xlsx` |
| `mapped column(s) "…" not found` | A column in Mapping is not in the export; fix or clear it in Mapping |
| `old .xls file` | Save the export as `.xlsx` |
| `cannot write … is it open in PowerPoint?` | Close the deck in PowerPoint, or run again (a new name is used) |
| `Python 3 was not found` | Install Python 3.9 or newer, or ask your IT |

The tool only reads your workbook and never changes it. Save and close Excel before a run,
because unsaved changes are not seen.

---

## For developers

- Specification (spec-driven): [docs/requirements.md](docs/requirements.md),
  [docs/design.md](docs/design.md), [docs/tasks.md](docs/tasks.md)
- For GitHub Copilot and other assistants: [.github/copilot-instructions.md](.github/copilot-instructions.md)
- Standard library only, Python 3.9+. From a checkout: `python -m statusreport …`

```
python -m unittest discover -s tests          # all tests
python tools/build_pyz.py                     # dist/sw-status-report/ and the zip
python tools/make_demo.py out/demo            # demo input with 6 weeks of history
python -m statusreport out/demo/Tasks.xlsx    # demo deck
python tools/opc_lint.py <file.pptx>          # structural check
powershell -File tools\ppt_check.ps1 <file.pptx> -PngDir out\png   # render in PowerPoint (dev only)
```
