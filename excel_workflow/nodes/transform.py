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
        raise ValueError("变换节点需要扁平 DataFrame")
    return d.copy()


@register_node
class Expression(BaseNode):
    __identifier__ = "excel_workflow.transform"
    NODE_NAME = "表达式列"

    def __init__(self):
        super().__init__()
        self.set_color(168, 85, 247)
        self.add_input("dataframe")
        self.add_output("dataframe")
        self.add_text_input("assignments", "如: total=数量*单价", tab="参数")


@register_runner("excel_workflow.transform.Expression")
def run_expr(node, inputs, ctx):
    df = _df(inputs)
    line = (node.get_property("assignments") or "").strip()
    if not line:
        return {"dataframe": df}
    # 支持分号分隔多赋值
    for part in line.split(";"):
        part = part.strip()
        if not part or "=" not in part:
            continue
        lhs, rhs = part.split("=", 1)
        lhs = lhs.strip()
        rhs = rhs.strip()
        df[lhs] = df.eval(rhs, engine="python")
    return {"dataframe": df}


@register_node
class DateFormat(BaseNode):
    __identifier__ = "excel_workflow.transform"
    NODE_NAME = "日期格式化"

    def __init__(self):
        super().__init__()
        self.set_color(168, 85, 247)
        self.add_input("dataframe")
        self.add_output("dataframe")
        self.add_text_input("column", "列名", tab="参数")
        self.add_text_input("fmt", "输出 strftime %Y-%m-%d", tab="参数")


@register_runner("excel_workflow.transform.DateFormat")
def run_date_fmt(node, inputs, ctx):
    df = _df(inputs)
    col = (node.get_property("column") or "").strip()
    fmt = (node.get_property("fmt") or "%Y-%m-%d").strip() or "%Y-%m-%d"
    if not col or col not in df.columns:
        raise ValueError("DateFormat: 列无效")
    s = pd.to_datetime(df[col], errors="coerce")
    df[col] = s.dt.strftime(fmt)
    return {"dataframe": df}


@register_node
class UnitConvert(BaseNode):
    __identifier__ = "excel_workflow.transform"
    NODE_NAME = "数值换算"

    def __init__(self):
        super().__init__()
        self.set_color(168, 85, 247)
        self.add_input("dataframe")
        self.add_output("dataframe")
        self.add_text_input("column", "列名", tab="参数")
        self.add_text_input("factor", "乘数", tab="参数")


@register_runner("excel_workflow.transform.UnitConvert")
def run_unit(node, inputs, ctx):
    df = _df(inputs)
    col = (node.get_property("column") or "").strip()
    fac = float((node.get_property("factor") or "1").strip() or "1")
    if not col or col not in df.columns:
        raise ValueError("UnitConvert: 列无效")
    df[col] = pd.to_numeric(df[col], errors="coerce") * fac
    return {"dataframe": df}


@register_node
class Concat(BaseNode):
    __identifier__ = "excel_workflow.transform"
    NODE_NAME = "列拼接"

    def __init__(self):
        super().__init__()
        self.set_color(168, 85, 247)
        self.add_input("dataframe")
        self.add_output("dataframe")
        self.add_text_input("columns", "列(逗号)", tab="参数")
        self.add_text_input("target", "新列名", tab="参数")
        self.add_text_input("sep", "分隔符", tab="参数")


@register_runner("excel_workflow.transform.Concat")
def run_concat(node, inputs, ctx):
    df = _df(inputs)
    cols = [c.strip() for c in (node.get_property("columns") or "").split(",") if c.strip()]
    tgt = (node.get_property("target") or "merged").strip() or "merged"
    sep = node.get_property("sep") or ""
    if len(cols) < 2:
        raise ValueError("Concat: 至少两列")
    parts = [df[c].astype(str) for c in cols if c in df.columns]
    df[tgt] = parts[0]
    for p in parts[1:]:
        df[tgt] = df[tgt] + sep + p
    return {"dataframe": df}


@register_node
class TypeCast(BaseNode):
    __identifier__ = "excel_workflow.transform"
    NODE_NAME = "类型转换"

    def __init__(self):
        super().__init__()
        self.set_color(168, 85, 247)
        self.add_input("dataframe")
        self.add_output("dataframe")
        self.add_text_input("column", "列名", tab="参数")
        self.add_combo_menu(
            "dtype",
            "目标类型",
            items=["str", "int", "float", "bool"],
            tab="参数",
        )


@register_runner("excel_workflow.transform.TypeCast")
def run_cast(node, inputs, ctx):
    df = _df(inputs)
    col = (node.get_property("column") or "").strip()
    dt = (node.get_property("dtype") or "str").strip()
    if not col or col not in df.columns:
        raise ValueError("TypeCast: 列无效")
    if dt == "str":
        df[col] = df[col].astype(str)
    elif dt == "int":
        df[col] = pd.to_numeric(df[col], errors="coerce").astype("Int64")
    elif dt == "float":
        df[col] = pd.to_numeric(df[col], errors="coerce")
    elif dt == "bool":
        df[col] = df[col].astype(bool)
    return {"dataframe": df}
