from __future__ import annotations

import json
import os
from pathlib import Path
from typing import Any, Dict, Optional


def save_workflow(graph, path: str) -> None:
    """保存 NodeGraph 会话到 JSON 文件。"""
    ap = os.path.abspath(path)
    parent = os.path.dirname(ap)
    if parent:
        os.makedirs(parent, exist_ok=True)
    graph.save_session(ap)


def load_workflow(graph, path: str) -> None:
    graph.clear_session()
    graph.load_session(os.path.abspath(path))
    from excel_workflow.core.node_theme import apply_all_node_colors

    apply_all_node_colors(graph)


def import_mapping_config(graph, path: str, pos: tuple = (0, 0)) -> None:
    """
    将旧版 mapping_config.json 转为画布上的 ExcelReader + TemplateMapper 并连线。
    """
    with open(path, encoding="utf-8") as f:
        cfg: Dict[str, Any] = json.load(f)

    x, y = pos
    reader = graph.create_node("excel_workflow.source.ExcelReader", pos=(x, y))
    mapper = graph.create_node(
        "excel_workflow.template.TemplateMapper", pos=(x + 320, y)
    )

    reader.set_property("file_path", cfg.get("data_file", ""))
    reader.set_property("sheet_name", "")

    mapper.set_property("template_path", cfg.get("tmpl_file", ""))
    mapper.set_property("output_dir", cfg.get("output_dir", ""))
    mapper.set_property("name_column", cfg.get("name_col", ""))
    mapper.set_property("sheet_name", cfg.get("sheet_name", "Sheet1"))
    mapping = cfg.get("mapping") or []
    # 过滤空行
    clean = [[a, b] for a, b in mapping if (a or "").strip() and (b or "").strip()]
    mapper.set_property("mapping_json", json.dumps(clean, ensure_ascii=False))

    # 连接 dataframe -> dataframe
    out_p = reader.output(0)
    in_p = mapper.input(0)
    out_p.connect_to(in_p)
