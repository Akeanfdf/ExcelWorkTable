from __future__ import annotations

import os
from pathlib import Path
from typing import Any, Dict

import pandas as pd
from NodeGraphQt import BaseNode

from excel_workflow.core.registry import register_node, register_runner


@register_node
class ExcelReader(BaseNode):
    __identifier__ = "excel_workflow.source"
    NODE_NAME = "Excel 读取"

    def __init__(self):
        super().__init__()
        self.set_color(59, 130, 246)
        self.add_output("dataframe")
        self.add_text_input("file_path", "文件路径", tab="参数")
        self.add_text_input("sheet_name", "工作表(空=默认)", tab="参数")


@register_runner("excel_workflow.source.ExcelReader")
def run_excel_reader(node, inputs: Dict[str, Any], ctx) -> Dict[str, Any]:
    path = (node.get_property("file_path") or "").strip()
    sheet = (node.get_property("sheet_name") or "").strip()
    if not path or not os.path.isfile(path):
        raise FileNotFoundError(f"ExcelReader: 文件不存在: {path}")
    if sheet:
        df = pd.read_excel(path, sheet_name=sheet)
    else:
        df = pd.read_excel(path)
    ctx.log(f"ExcelReader: 读取 {len(df)} 行 × {len(df.columns)} 列")
    return {"dataframe": df}


@register_node
class CSVReader(BaseNode):
    __identifier__ = "excel_workflow.source"
    NODE_NAME = "CSV 读取"

    def __init__(self):
        super().__init__()
        self.set_color(59, 130, 246)
        self.add_output("dataframe")
        self.add_text_input("file_path", "文件路径", tab="参数")
        self.add_text_input("encoding", "编码 utf-8", tab="参数")


@register_runner("excel_workflow.source.CSVReader")
def run_csv_reader(node, inputs, ctx):
    path = (node.get_property("file_path") or "").strip()
    enc = (node.get_property("encoding") or "utf-8").strip() or "utf-8"
    if not path or not os.path.isfile(path):
        raise FileNotFoundError(f"CSVReader: 文件不存在: {path}")
    df = pd.read_csv(path, encoding=enc)
    ctx.log(f"CSVReader: 读取 {len(df)} 行")
    return {"dataframe": df}


@register_node
class FolderWatcher(BaseNode):
    __identifier__ = "excel_workflow.source"
    NODE_NAME = "文件夹扫描"

    def __init__(self):
        super().__init__()
        self.set_color(59, 130, 246)
        self.add_output("paths")
        self.add_text_input("folder", "文件夹", tab="参数")
        self.add_text_input("pattern", "通配符 *.xlsx", tab="参数")
        self.add_checkbox("recursive", "", "递归子目录", False)


@register_runner("excel_workflow.source.FolderWatcher")
def run_folder_watcher(node, inputs, ctx):
    folder = (node.get_property("folder") or "").strip()
    pattern = (node.get_property("pattern") or "*").strip() or "*"
    rec = bool(node.get_property("recursive"))
    if not folder or not os.path.isdir(folder):
        raise FileNotFoundError(f"FolderWatcher: 目录无效: {folder}")
    paths = []
    base = Path(folder)
    if rec:
        for p in base.rglob(pattern):
            if p.is_file():
                paths.append(str(p.resolve()))
    else:
        for p in base.glob(pattern):
            if p.is_file():
                paths.append(str(p.resolve()))
    paths.sort()
    ctx.log(f"FolderWatcher: 匹配 {len(paths)} 个文件")
    return {"paths": paths}


@register_node
class Schedule(BaseNode):
    __identifier__ = "excel_workflow.source"
    NODE_NAME = "计划触发"

    def __init__(self):
        super().__init__()
        self.set_color(59, 130, 246)
        self.add_output("tick")
        self.add_text_input("cron", "Cron 表达式", tab="参数")
        self.add_text_input("note", "说明", tab="参数")


@register_runner("excel_workflow.source.Schedule")
def run_schedule(node, inputs, ctx):
    ctx.log(
        "Schedule: 手动运行仅输出 tick=True；后台定时请用工具栏「启动调度」。"
    )
    return {"tick": True}
