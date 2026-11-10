"""Check a .pptx/.xlsx package for structural problems (dev tool).

    python tools/opc_lint.py out/spike.pptx
"""

from __future__ import annotations

import os
import sys

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, ROOT)

from statusreport.pptx.lint import lint_package  # noqa: E402
from statusreport.pptx.opc import Package  # noqa: E402


def main() -> int:
    if len(sys.argv) < 2:
        print(__doc__)
        return 2
    status = 0
    for path in sys.argv[1:]:
        problems = lint_package(Package.open(path))
        print(f"{path}: {'OK' if not problems else f'{len(problems)} problem(s)'}")
        for p in problems:
            print("  " + p)
        status = status or (1 if problems else 0)
    return status


if __name__ == "__main__":
    sys.exit(main())
