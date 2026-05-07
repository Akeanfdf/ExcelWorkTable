from __future__ import annotations

from typing import Any, Dict

import pandas as pd
from NodeGraphQt import BaseNode

from excel_workflow.core.registry import register_node, register_runner


@register_node
class ForEachRow(BaseNode):
    __identifier__ = "excel_workflow.control"
    NODE_NAME = "逐行(透传)"

    def __init__(self):
        super().__init__()
        self.set_color(30, 41, 59)
        self.add_input("dataframe")
        self.add_output("dataframe")


@register_runner("excel_workflow.control.ForEachRow")
def run_for_each(node, inputs, ctx):
    df = inputs.get("dataframe")
    if isinstance(df, dict):
        raise ValueError("ForEachRow: 不支持分组 dict")
    if df is None:
        raise ValueError("ForEachRow: 无输入")
    ctx.log(f"ForEachRow: 透传 {len(df)} 行（复杂子流程请用模板节点内建循环）")
    return {"dataframe": df}


@register_node
class TryCatch(BaseNode):
    __identifier__ = "excel_workflow.control"
    NODE_NAME = "Try 透传"

    def __init__(self):
        super().__init__()
        self.set_color(30, 41, 59)
        self.add_input("dataframe")
        self.add_output("dataframe")
        self.add_text_input("note", "备注", tab="参数")


@register_runner("excel_workflow.control.TryCatch")
def run_try_catch(node, inputs, ctx):
    df = inputs.get("dataframe")
    if df is None:
        return {"dataframe": pd.DataFrame()}
    return {"dataframe": df}


@register_node
class Output(BaseNode):
    __identifier__ = "excel_workflow.control"
    NODE_NAME = "输出(日志)"

    def __init__(self):
        super().__init__()
        self.set_color(30, 41, 59)
        self.add_input("payload")


@register_runner("excel_workflow.control.Output")
def run_output(node, inputs, ctx):
    p = inputs.get("payload")
    ctx.log(f"Output: {type(p).__name__} => {repr(p)[:500]}")
    return {}
