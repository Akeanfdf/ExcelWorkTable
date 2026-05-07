"""Folder glob scan for import phase."""

from __future__ import annotations

import glob
import os
from typing import List


def scan_folder(folder: str, pattern: str = "*.xlsx") -> List[str]:
    folder = os.path.abspath(folder)
    if not os.path.isdir(folder):
        return []
    pat = pattern.strip() or "*"
    paths = sorted(glob.glob(os.path.join(folder, pat)), key=str.lower)
    return [p for p in paths if os.path.isfile(p)]


def scan_spreadsheets(folder: str) -> List[str]:
    """扫描常见表格：xlsx / xls / csv（去重）。"""
    seen: set[str] = set()
    out: List[str] = []
    for pat in ("*.xlsx", "*.xls", "*.csv"):
        for p in scan_folder(folder, pat):
            ap = os.path.abspath(p)
            if ap not in seen:
                seen.add(ap)
                out.append(ap)
    return sorted(out, key=str.lower)
