from __future__ import annotations

import fnmatch
import json
import os
import shutil
from typing import Any, Dict, List

from NodeGraphQt import BaseNode

from excel_workflow.core.registry import register_node, register_runner


def _as_file_list(x) -> List[str]:
    if x is None:
        return []
    if isinstance(x, str):
        return [x] if x else []
    if isinstance(x, (list, tuple)):
        return [str(p) for p in x if str(p)]
    return []


@register_node
class Rename(BaseNode):
    __identifier__ = "excel_workflow.fileops"
    NODE_NAME = "批量重命名"

    def __init__(self):
        super().__init__()
        self.set_color(100, 116, 139)
        self.add_input("files")
        self.add_output("files")
        self.add_text_input(
            "replacements_json",
            '文件名替换 JSON {"旧":"新"}',
            tab="参数",
        )


@register_runner("excel_workflow.fileops.Rename")
def run_rename_files(node, inputs, ctx):
    files = _as_file_list(inputs.get("files"))
    raw = (node.get_property("replacements_json") or "{}").strip() or "{}"
    mp = json.loads(raw)
    if not isinstance(mp, dict):
        raise ValueError("replacements_json 须为对象")
    out: List[str] = []
    for p in files:
        if not os.path.isfile(p):
            continue
        d, base = os.path.dirname(p), os.path.basename(p)
        new_base = base
        for k, v in mp.items():
            new_base = new_base.replace(str(k), str(v))
        new_path = os.path.join(d, new_base)
        if new_path != p:
            os.rename(p, new_path)
            ctx.log(f"Rename: {p} -> {new_path}")
            out.append(new_path)
        else:
            out.append(p)
    return {"files": out}


@register_node
class Move(BaseNode):
    __identifier__ = "excel_workflow.fileops"
    NODE_NAME = "移动文件"

    def __init__(self):
        super().__init__()
        self.set_color(100, 116, 139)
        self.add_input("files")
        self.add_output("files")
        self.add_text_input("dest_dir", "目标目录", tab="参数")


@register_runner("excel_workflow.fileops.Move")
def run_move(node, inputs, ctx):
    files = _as_file_list(inputs.get("files"))
    dest = (node.get_property("dest_dir") or "").strip()
    if not dest:
        raise ValueError("Move: 目标目录未设置")
    os.makedirs(dest, exist_ok=True)
    out: List[str] = []
    for p in files:
        if os.path.isfile(p):
            t = os.path.join(dest, os.path.basename(p))
            shutil.move(p, t)
            ctx.log(f"Move: {t}")
            out.append(t)
    return {"files": out}


@register_node
class GroupByFolder(BaseNode):
    __identifier__ = "excel_workflow.fileops"
    NODE_NAME = "按规则分目录"

    def __init__(self):
        super().__init__()
        self.set_color(100, 116, 139)
        self.add_input("files")
        self.add_output("files")
        self.add_text_input("base_dir", "根目录", tab="参数")
        self.add_text_input(
            "rules_json",
            '[{"pattern":"*检验*","folder":"检验批"}]',
            tab="参数",
        )


@register_runner("excel_workflow.fileops.GroupByFolder")
def run_group_folder(node, inputs, ctx):
    files = _as_file_list(inputs.get("files"))
    base = (node.get_property("base_dir") or "").strip()
    raw = (node.get_property("rules_json") or "[]").strip() or "[]"
    rules = json.loads(raw)
    if not isinstance(rules, list):
        raise ValueError("rules_json 须为数组")
    if not base:
        raise ValueError("base_dir 未设置")
    os.makedirs(base, exist_ok=True)
    out: List[str] = []
    for p in files:
        if not os.path.isfile(p):
            continue
        name = os.path.basename(p)
        sub = ""
        for r in rules:
            pat = r.get("pattern") or "*"
            folder = (r.get("folder") or "other").strip() or "other"
            if fnmatch.fnmatch(name, pat):
                sub = folder
                break
        if not sub:
            sub = "other"
        target_dir = os.path.join(base, sub)
        os.makedirs(target_dir, exist_ok=True)
        t = os.path.join(target_dir, name)
        shutil.copy2(p, t)
        out.append(t)
        ctx.log(f"GroupByFolder: {t}")
    return {"files": out}


@register_node
class Archive(BaseNode):
    __identifier__ = "excel_workflow.fileops"
    NODE_NAME = "归档复制"

    def __init__(self):
        super().__init__()
        self.set_color(100, 116, 139)
        self.add_input("files")
        self.add_output("files")
        self.add_text_input("archive_dir", "归档目录", tab="参数")


@register_runner("excel_workflow.fileops.Archive")
def run_archive(node, inputs, ctx):
    files = _as_file_list(inputs.get("files"))
    adir = (node.get_property("archive_dir") or "").strip()
    if not adir:
        raise ValueError("archive_dir 未设置")
    os.makedirs(adir, exist_ok=True)
    out: List[str] = []
    for p in files:
        if os.path.isfile(p):
            t = os.path.join(adir, os.path.basename(p))
            shutil.copy2(p, t)
            out.append(t)
    return {"files": out}
