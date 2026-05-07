from __future__ import annotations

import json
from typing import Any, Dict

import pandas as pd
from NodeGraphQt import BaseNode

from excel_workflow.core.registry import register_node, register_runner


def _df(inputs):
    d = inputs.get("dataframe")
    if d is None:
        raise ValueError("需要上游 dataframe")
    if isinstance(d, dict):
        raise ValueError("此节点需要扁平 DataFrame")
    return d.copy()


@register_node
class FilterRows(BaseNode):
    __identifier__ = "excel_workflow.filter"
    NODE_NAME = "行过滤"

    def __init__(self):
        super().__init__()
        self.set_color(234, 179, 8)
        self.add_input("dataframe")
        self.add_output("dataframe")
        self.add_text_input("query", "pandas query 表达式", tab="参数")


@register_runner("excel_workflow.filter.FilterRows")
def run_filter(node, inputs, ctx):
    df = _df(inputs)
    q = (node.get_property("query") or "").strip()
    if not q:
        return {"dataframe": df}
    out = df.query(q, engine="python")
    ctx.log(f"FilterRows: {len(df)} -> {len(out)}")
    return {"dataframe": out}


@register_node
class GroupSplit(BaseNode):
    __identifier__ = "excel_workflow.filter"
    NODE_NAME = "分组字典"

    def __init__(self):
        super().__init__()
        self.set_color(234, 179, 8)
        self.add_input("dataframe")
        self.add_output("dataframe")
        self.add_text_input("group_column", "分组列名", tab="参数")


@register_runner("excel_workflow.filter.GroupSplit")
def run_group_split(node, inputs, ctx):
    df = _df(inputs)
    col = (node.get_property("group_column") or "").strip()
    if not col or col not in df.columns:
        raise ValueError(f"GroupSplit: 无效列 {col}")
    parts = {str(k): g.copy() for k, g in df.groupby(col, dropna=False)}
    ctx.log(f"GroupSplit: {len(parts)} 组")
    return {"dataframe": parts}


@register_node
class Sample(BaseNode):
    __identifier__ = "excel_workflow.filter"
    NODE_NAME = "抽样"

    def __init__(self):
        super().__init__()
        self.set_color(234, 179, 8)
        self.add_input("dataframe")
        self.add_output("dataframe")
        self.add_spinbox("n", "行数", 100, 1, 10_000_000, tab="参数", double=False)


@register_runner("excel_workflow.filter.Sample")
def run_sample(node, inputs, ctx):
    df = _df(inputs)
    n = int(node.get_property("n") or 100)
    out = df.head(n)
    return {"dataframe": out}


@register_node
class Branch(BaseNode):
    __identifier__ = "excel_workflow.filter"
    NODE_NAME = "条件分支"

    def __init__(self):
        super().__init__()
        self.set_color(234, 179, 8)
        self.add_input("dataframe")
        self.add_output("true_df")
        self.add_output("false_df")
        self.add_text_input("mask_expr", "布尔掩码 pandas 表达式", tab="参数")


@register_runner("excel_workflow.filter.Branch")
def run_branch(node, inputs, ctx):
    df = _df(inputs)
    expr = (node.get_property("mask_expr") or "").strip()
    if not expr:
        raise ValueError("Branch: 需要 mask_expr")
    mask = df.eval(expr, engine="python")
    t = df.loc[mask].copy()
    f = df.loc[~mask].copy()
    return {"true_df": t, "false_df": f}
