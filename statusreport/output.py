"""Output file name, overwrite protection and template lookup — R-OUT-8, R-OUT-9."""

from __future__ import annotations

import os
import re
from typing import Optional

from . import labels

_BAD_CHARS = re.compile(r'[<>:"/\\|?*\x00-\x1f]')


def default_name(project: str, week: str) -> str:
    """``<Project> Status 2026-W46.pptx``; characters Windows forbids in file names are dropped."""
    clean = " ".join(_BAD_CHARS.sub("", project).split()).strip(". ")
    if not clean:
        return labels.OUTPUT_NAME_NO_PROJECT.format(week=week)
    return labels.OUTPUT_NAME.format(project=clean, week=week)


def free_path(path: str, force: bool = False) -> str:
    """``path`` itself if it is free or ``force`` is set, else ``name (2).pptx``, ``name (3).pptx`` …"""
    if force or not os.path.exists(path):
        return path
    base, ext = os.path.splitext(path)
    n = 2
    while os.path.exists(f"{base} ({n}){ext}"):
        n += 1
    return f"{base} ({n}){ext}"


def template_path(input_path: str, setting: str, cli: Optional[str]) -> Optional[str]:
    """The customer template: ``--template`` wins over the Settings value, which is relative to the workbook."""
    if cli:
        return cli
    if not setting:
        return None
    if os.path.isabs(setting):
        return setting
    return os.path.join(os.path.dirname(os.path.abspath(input_path)), setting)


__all__ = ["default_name", "free_path", "template_path"]
