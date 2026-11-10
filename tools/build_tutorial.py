"""Build the German tutorial as one self-contained HTML file (dev tool).

    python tools/build_tutorial.py [out.html] [--fragment]

Source: docs/tutorial/tutorial.html (references img/*.png). The images are embedded
as data URIs, so the result works offline and can be mailed or put on a share. By
default it is wrapped into a complete HTML document; --fragment writes the page
content only (for publishing as a claude.ai artifact, which adds its own skeleton).
"""

from __future__ import annotations

import base64
import os
import re
import sys

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
SRC = os.path.join(ROOT, "docs", "tutorial", "tutorial.html")
DEFAULT_OUT = os.path.join(ROOT, "dist", "Tutorial (Deutsch).html")


def inline_images(html: str, base: str) -> str:
    def repl(m: "re.Match[str]") -> str:
        with open(os.path.join(base, m.group(1)), "rb") as fh:
            data = base64.b64encode(fh.read()).decode("ascii")
        return f'src="data:image/png;base64,{data}"'

    return re.sub(r'src="(img/[^"]+\.png)"', repl, html)


def build(out: str = DEFAULT_OUT, fragment: bool = False) -> str:
    with open(SRC, encoding="utf-8") as fh:
        html = inline_images(fh.read(), os.path.dirname(SRC))
    if not fragment:
        html = (
            '<!doctype html>\n<html lang="de">\n<head>\n<meta charset="utf-8">\n'
            '<meta name="viewport" content="width=device-width, initial-scale=1, viewport-fit=cover">\n'
            "</head>\n<body>\n" + html + "\n</body>\n</html>\n"
        )
    os.makedirs(os.path.dirname(os.path.abspath(out)), exist_ok=True)
    with open(out, "w", encoding="utf-8", newline="\n") as fh:
        fh.write(html)
    return out


if __name__ == "__main__":
    args = [a for a in sys.argv[1:] if not a.startswith("--")]
    path = build(args[0] if args else DEFAULT_OUT, fragment="--fragment" in sys.argv)
    print(f"written {path} ({os.path.getsize(path) // 1024} KB)")
