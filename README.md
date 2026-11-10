# SW Status Report

Weekly status deck for customer meetings: turns a list of software work packages
(Excel, typed in by hand or exported from Polarion) into a PowerPoint with a progress
trend per version and one status table per version, showing progress (previous week)
and a traffic light per task.

Runs with plain Python 3.9+, standard library only, no installation.

**Status:** M1 (PowerPoint spike) done — see [docs/requirements.md](docs/requirements.md),
[docs/design.md](docs/design.md) and [docs/tasks.md](docs/tasks.md).

## Development

```
python -m unittest discover -s tests     # tests
python tools/spike.py                    # sample deck -> out/spike.pptx
python tools/opc_lint.py out/spike.pptx  # structural check
python tools/build_base_pptx.py          # regenerate the built-in base template
```

For GitHub Copilot and other AI assistants: [.github/copilot-instructions.md](.github/copilot-instructions.md).
