from __future__ import annotations

from typing import Any, Dict

import pandas as pd
from NodeGraphQt import BaseNode

from excel_workflow.core.registry import register_node, register_runner


def _df(inputs, key="dataframe"):
    d = inputs.get(key)
    if d is None:
        raise ValueError(f"需要输入 {key}")
    if isinstance(d, dict):
        raise ValueError("合并节点需要 DataFrame")
    return d.copy()


@register_node
class VerticalMerge(BaseNode):
    __identifier__ = "excel_workflow.merge"
    NODE_NAME = "纵向合并"

    def __init__(self):
        super().__init__()
        self.set_color(34, 197, 94)
        self.add_input("left")
        self.add_input("right")
        self.add_output("dataframe")
        self.add_checkbox("ignore_index", "", "重置索引", True)


@register_runner("excel_workflow.merge.VerticalMerge")
def run_vmerge(node, inputs, ctx):
    a = _df(inputs, "left")
    b = _df(inputs, "right")
    ig = bool(node.get_property("ignore_index"))
    out = pd.concat([a, b], axis=0, ignore_index=ig)
    ctx.log(f"VerticalMerge: {len(out)} 行")
    return {"dataframe": out}


@register_node
class Join(BaseNode):
    __identifier__ = "excel_workflow.merge"
    NODE_NAME = "表关联"

    def __init__(self):
        super().__init__()
        self.set_color(34, 197, 94)
        self.add_input("left")
        self.add_input("right")
        self.add_output("dataframe")
        self.add_text_input("on_columns", "关联列(逗号)", tab="参数")
        self.add_combo_menu(
            "how",
            "方式",
            items=["inner", "left", "right", "outer"],
            tab="参数",
        )


@register_runner("excel_workflow.merge.Join")
def run_join(node, inputs, ctx):
    left = _df(inputs, "left")
    right = _df(inputs, "right")
    on = [c.strip() for c in (node.get_property("on_columns") or "").split(",") if c.strip()]
    if not on:
        raise ValueError("Join: 需要 on 列")
    how = (node.get_property("how") or "inner").strip()
    out = pd.merge(left, right, on=on, how=how)
    ctx.log(f"Join: {len(out)} 行")
    return {"dataframe": out}


@register_node
class Pivot(BaseNode):
    __identifier__ = "excel_workflow.merge"
    NODE_NAME = "透视"

    def __init__(self):
        super().__init__()
        self.set_color(34, 197, 94)
        self.add_input("dataframe")
        self.add_output("dataframe")
        self.add_text_input("index", "index 列", tab="参数")
        self.add_text_input("columns", "columns 列", tab="参数")
        self.add_text_input("values", "values 列", tab="参数")
        self.add_combo_menu(
            "aggfunc",
            "聚合",
            items=["mean", "sum", "count", "max", "min"],
            tab="参数",
        )


@register_runner("excel_workflow.merge.Pivot")
def run_pivot(node, inputs, ctx):
    df = _df(inputs, "dataframe")
    idx = (node.get_property("index") or "").strip()
    cols = (node.get_property("columns") or "").strip()
    vals = (node.get_property("values") or "").strip()
    agg = (node.get_property("aggfunc") or "mean").strip()
    if not idx or not cols or not vals:
        raise ValueError("Pivot: 需要 index/columns/values")
    aggfunc = agg if agg in ("mean", "sum", "count", "max", "min") else "mean"
    out = pd.pivot_table(df, index=idx, columns=cols, values=vals, aggfunc=aggfunc)
    out = out.reset_index()
    return {"dataframe": out}
