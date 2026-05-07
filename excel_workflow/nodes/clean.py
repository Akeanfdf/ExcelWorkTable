from __future__ import annotations

from typing import Any, Dict

import pandas as pd
from NodeGraphQt import BaseNode

from excel_workflow.core.registry import register_node, register_runner


def _df(inputs):
    d = inputs.get("dataframe")
    if d is None:
        raise ValueError("需要上游 dataframe")
    if isinstance(d, dict):
        raise ValueError("清洗节点暂不支持 dict 分组输入，请先展开或接扁平表")
    return d.copy()


@register_node
class Deduplicate(BaseNode):
    __identifier__ = "excel_workflow.clean"
    NODE_NAME = "去重"

    def __init__(self):
        super().__init__()
        self.set_color(20, 184, 166)
        self.add_input("dataframe")
        self.add_output("dataframe")
        self.add_text_input("subset", "列名(逗号分隔，空=整行)", tab="参数")
        self.add_checkbox("keep", "", "保留 first", True)


@register_runner("excel_workflow.clean.Deduplicate")
def run_dedup(node, inputs, ctx):
    df = _df(inputs)
    sub = (node.get_property("subset") or "").strip()
    subset = [c.strip() for c in sub.split(",") if c.strip()] if sub else None
    keep = "first" if node.get_property("keep") else "last"
    out = df.drop_duplicates(subset=subset, keep=keep)
    ctx.log(f"Deduplicate: {len(df)} -> {len(out)}")
    return {"dataframe": out}


@register_node
class FillEmpty(BaseNode):
    __identifier__ = "excel_workflow.clean"
    NODE_NAME = "空值填充"

    def __init__(self):
        super().__init__()
        self.set_color(20, 184, 166)
        self.add_input("dataframe")
        self.add_output("dataframe")
        self.add_text_input("fill_value", "填充值", tab="参数")
        self.add_text_input("columns", "列(逗号，空=全部)", tab="参数")


@register_runner("excel_workflow.clean.FillEmpty")
def run_fill_empty(node, inputs, ctx):
    df = _df(inputs)
    val = node.get_property("fill_value")
    cols = (node.get_property("columns") or "").strip()
    if cols:
        use = [c.strip() for c in cols.split(",") if c.strip()]
        for c in use:
            if c in df.columns:
                df[c] = df[c].fillna(val)
    else:
        df = df.fillna(val)
    return {"dataframe": df}


@register_node
class FormatStandardize(BaseNode):
    __identifier__ = "excel_workflow.clean"
    NODE_NAME = "格式标准化"

    def __init__(self):
        super().__init__()
        self.set_color(20, 184, 166)
        self.add_input("dataframe")
        self.add_output("dataframe")
        self.add_text_input("columns", "文本列(逗号)", tab="参数")
        self.add_checkbox("strip", "", "strip 首尾空格", True)


@register_runner("excel_workflow.clean.FormatStandardize")
def run_fmt(node, inputs, ctx):
    df = _df(inputs)
    cols = [c.strip() for c in (node.get_property("columns") or "").split(",") if c.strip()]
    strip = bool(node.get_property("strip"))
    for c in cols:
        if c in df.columns:
            s = df[c].astype(str)
            if strip:
                s = s.str.strip()
            df[c] = s
    return {"dataframe": df}


@register_node
class RenameColumns(BaseNode):
    __identifier__ = "excel_workflow.clean"
    NODE_NAME = "列重命名"

    def __init__(self):
        super().__init__()
        self.set_color(20, 184, 166)
        self.add_input("dataframe")
        self.add_output("dataframe")
        self.add_text_input(
            "mapping_json",
            'JSON 对象 {"旧":"新"}',
            tab="参数",
        )


@register_runner("excel_workflow.clean.RenameColumns")
def run_rename(node, inputs, ctx):
    import json

    df = _df(inputs)
    raw = (node.get_property("mapping_json") or "{}").strip() or "{}"
    mp = json.loads(raw)
    if not isinstance(mp, dict):
        raise ValueError("mapping_json 必须是 JSON 对象")
    df = df.rename(columns={str(k): str(v) for k, v in mp.items()})
    return {"dataframe": df}
