from __future__ import annotations

import json
import os
import re
from typing import Any, List, Tuple


def sanitize_filename(name: str) -> str:
    s = str(name)
    for ch in r'\/:*?"<>|':
        s = s.replace(ch, "_")
    return s.strip() or "output"


def parse_mapping_json(s: str) -> List[Tuple[str, str]]:
    if not (s or "").strip():
        return []
    data = json.loads(s)
    out: List[Tuple[str, str]] = []
    for row in data:
        if isinstance(row, (list, tuple)) and len(row) >= 2:
            a, b = str(row[0]).strip(), str(row[1]).strip()
            if a and b:
                out.append((a, b.upper()))
    return out


def load_secrets() -> dict:
    p = os.path.join(os.path.expanduser("~"), ".excel_workflow", "secrets.json")
    if not os.path.isfile(p):
        return {}
    try:
        with open(p, encoding="utf-8") as f:
            return json.load(f)
    except Exception:
        return {}
