"""
应用数据根目录与 staging：默认尽量不用 C 盘。

1) EXCEL_WORKFLOW_DATA_ROOT
2) Windows: D:～Z: 下 .excel_workflow
3) Windows: 当前工作目录非 C: 时使用该盘
4) 回退：用户主目录 .excel_workflow
"""

from __future__ import annotations

import os
from pathlib import Path


def excel_workflow_data_root() -> Path:
    env = (os.environ.get("EXCEL_WORKFLOW_DATA_ROOT") or "").strip()
    if env:
        p = Path(os.path.expandvars(env)).expanduser().resolve()
        p.mkdir(parents=True, exist_ok=True)
        return p

    if os.name == "nt":
        for letter in "DEFGHIJKLMNOPQRSTUVWXYZ":
            root = f"{letter}:\\"
            if os.path.isdir(root):
                try:
                    p = Path(root) / ".excel_workflow"
                    p.mkdir(parents=True, exist_ok=True)
                    return p.resolve()
                except OSError:
                    continue
        try:
            anchor = Path.cwd().resolve().anchor
            if anchor and len(anchor) >= 2:
                drive = anchor[0].upper()
                if drive != "C":
                    p = Path(anchor) / ".excel_workflow"
                    p.mkdir(parents=True, exist_ok=True)
                    return p.resolve()
        except OSError:
            pass

    p = Path.home() / ".excel_workflow"
    p.mkdir(parents=True, exist_ok=True)
    return p.resolve()


def staging_sessions_dir() -> Path:
    d = excel_workflow_data_root() / "staging"
    d.mkdir(parents=True, exist_ok=True)
    return d
