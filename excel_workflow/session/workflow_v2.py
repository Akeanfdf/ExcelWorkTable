from __future__ import annotations

import json
from datetime import datetime, timezone
from pathlib import Path
from typing import Any, Dict, List, Optional


def _now_name() -> str:
    return datetime.now(timezone.utc).replace(microsecond=0).isoformat()


def save_workflow_v2(
    path: str,
    name: str,
    ops: List[Dict[str, Any]],
    legacy_mapping_config: Optional[dict] = None,
) -> None:
    data = {
        "version": "2.0",
        "name": name,
        "created_at": _now_name(),
        "ops": ops,
        "legacy_mapping_config": legacy_mapping_config,
    }
    p = Path(path)
    p.parent.mkdir(parents=True, exist_ok=True)
    with open(p, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)


def load_workflow_v2(path: str) -> Dict[str, Any]:
    with open(path, encoding="utf-8") as f:
        return json.load(f)


def mapping_config_to_template_ops(cfg: Dict[str, Any]) -> List[Dict[str, Any]]:
    """旧 mapping_config.json → v2 ops（首步为 template_fill 参数包）。"""
    import json as _json

    mapping = cfg.get("mapping") or []
    clean = [[a, b] for a, b in mapping if (a or "").strip() and (b or "").strip()]
    return [
        {
            "op_type": "template_fill",
            "params": {
                "data_file": cfg.get("data_file", ""),
                "tmpl_file": cfg.get("tmpl_file", ""),
                "output_dir": cfg.get("output_dir", ""),
                "name_col": cfg.get("name_col", ""),
                "sheet_name": cfg.get("sheet_name", "Sheet1"),
                "mapping_json": _json.dumps(clean, ensure_ascii=False),
            },
        }
    ]
