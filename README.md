# SW Status Report

Weekly status deck for customer meetings: turns a list of software work packages
(Excel, typed in by hand or exported from Polarion) into a PowerPoint with a progress
trend per version and one status table per version, showing progress (previous week)
and a traffic light per task.

Runs with plain Python 3.9+, standard library only, no installation.

**Status:** M2 (Excel input) done, the deck follows in M5 — see
[docs/requirements.md](docs/requirements.md), [docs/design.md](docs/design.md) and
[docs/tasks.md](docs/tasks.md).

## Usage so far

```
python -m statusreport --init Tasks.xlsx                       # create the input workbook
python -m statusreport Tasks.xlsx --check                      # validate it, list problems
python -m statusreport Tasks.xlsx --polarion export.xlsx --check
```

## Development

```
python -m unittest discover -s tests     # tests
powershell -File tools\make_fixtures.ps1 # regenerate Excel-made fixtures (needs Excel)
python tools/spike.py                    # sample deck -> out/spike.pptx
python tools/opc_lint.py out/spike.pptx  # structural check
python tools/build_base_pptx.py          # regenerate the built-in base template
```

For GitHub Copilot and other AI assistants: [.github/copilot-instructions.md](.github/copilot-instructions.md).
