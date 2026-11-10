# SW Status Report

Weekly status deck for customer meetings: turns a list of software work packages
(Excel, typed in by hand or exported from Polarion) into a PowerPoint with a progress
trend per version and one status table per version, showing progress (previous week)
and a traffic light per task.

Runs with plain Python 3.9+, standard library only, no installation.

**Status:** M5 (deck) done; customer template (M6) and packaging (M7) follow — see
[docs/requirements.md](docs/requirements.md), [docs/design.md](docs/design.md) and
[docs/tasks.md](docs/tasks.md).

## Usage

```
python -m statusreport --init Tasks.xlsx                       # create the input workbook
python -m statusreport Tasks.xlsx --check                      # validate and calculate, write nothing
python -m statusreport Tasks.xlsx                              # weekly run: deck + history/YYYY-Www.json
python -m statusreport Tasks.xlsx --polarion export.xlsx       # with a separate Polarion export
python -m statusreport Tasks.xlsx --no-save                    # trial run: deck only, week not stored
python -m statusreport Tasks.xlsx --template customer.pptx     # use a customer template
```

Demo with seven weeks of data: `python tools/make_demo.py out/demo`, then
`python -m statusreport out/demo/Tasks.xlsx`.

## Development

```
python -m unittest discover -s tests     # tests
powershell -File tools\make_fixtures.ps1 # regenerate Excel-made fixtures (needs Excel)
python tools/spike.py                    # sample deck -> out/spike.pptx
python tools/opc_lint.py out/spike.pptx  # structural check
python tools/build_base_pptx.py          # regenerate the built-in base template
```

For GitHub Copilot and other AI assistants: [.github/copilot-instructions.md](.github/copilot-instructions.md).
