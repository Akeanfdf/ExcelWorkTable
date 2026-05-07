from __future__ import annotations

import json
import os
from pathlib import Path
from typing import Any, Dict, List

from excel_workflow.simple.paths import excel_workflow_data_root


def templates_root() -> Path:
    p = excel_workflow_data_root() / "simple_templates"
    p.mkdir(parents=True, exist_ok=True)
    return p


def save_template(name: str, glob_pattern: str, steps: List[Dict[str, Any]]) -> Path:
    safe = "".join(c for c in name if c.isalnum() or c in (" ", "-", "_")).strip() or "template"
    path = templates_root() / f"{safe}.json"
    data = {"version": 1, "name": name, "glob": glob_pattern, "steps": steps}
    with open(path, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)
    return path


def load_template(path: str) -> Dict[str, Any]:
    with open(path, encoding="utf-8") as f:
        return json.load(f)


def list_templates() -> List[Path]:
    root = templates_root()
    return sorted(root.glob("*.json"), key=lambda p: p.stat().st_mtime, reverse=True)
