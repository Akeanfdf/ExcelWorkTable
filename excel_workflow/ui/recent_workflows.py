"""Persist recently opened workflow paths (user profile)."""

from __future__ import annotations

import json
from pathlib import Path
from typing import List

from excel_workflow.simple.paths import excel_workflow_data_root

_MAX = 5


def _path() -> Path:
    return excel_workflow_data_root() / "recent.json"


def load_recent() -> List[str]:
    p = _path()
    if not p.is_file():
        return []
    try:
        with open(p, encoding="utf-8") as f:
            data = json.load(f)
        if isinstance(data, list):
            return [str(x) for x in data if x][: _MAX]
    except (json.JSONDecodeError, OSError):
        pass
    return []


def push_recent(path: str) -> None:
    path = os.path.abspath(path)
    cur = load_recent()
    cur = [p for p in cur if p != path]
    cur.insert(0, path)
    cur = cur[:_MAX]
    try:
        with open(_path(), "w", encoding="utf-8") as f:
            json.dump(cur, f, ensure_ascii=False, indent=2)
    except OSError:
        pass
