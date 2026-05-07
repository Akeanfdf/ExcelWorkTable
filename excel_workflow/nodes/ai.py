from __future__ import annotations

import json
import os
from typing import Any, Dict, List

import pandas as pd
from NodeGraphQt import BaseNode

from excel_workflow.core.registry import register_node, register_runner
from excel_workflow.nodes.utils import load_secrets


def _first_path(x) -> str:
    if x is None:
        return ""
    if isinstance(x, str):
        return x.strip()
    if isinstance(x, (list, tuple)) and x:
        return str(x[0]).strip()
    return ""


@register_node
class PDFExtract(BaseNode):
    __identifier__ = "excel_workflow.ai"
    NODE_NAME = "PDF 表格提取"

    def __init__(self):
        super().__init__()
        self.set_color(236, 72, 153)
        self.add_input("path")
        self.add_output("dataframe")
        self.add_spinbox("page", "页码 0=全部", 0, 0, 9999, tab="参数", double=False)


@register_runner("excel_workflow.ai.PDFExtract")
def run_pdf_extract(node, inputs, ctx):
    import pdfplumber

    path = _first_path(inputs.get("path"))
    if not path or not os.path.isfile(path):
        raise FileNotFoundError("PDFExtract: 文件无效")
    page_idx = int(node.get_property("page") or 0)
    frames: List[pd.DataFrame] = []
    with pdfplumber.open(path) as pdf:
        pages = pdf.pages if page_idx == 0 else [pdf.pages[page_idx - 1]]
        for pg in pages:
            t = pg.extract_table()
            if t and len(t) > 1:
                df = pd.DataFrame(t[1:], columns=t[0])
                frames.append(df)
    if not frames:
        ctx.log("PDFExtract: 未解析到表格，返回空表")
        return {"dataframe": pd.DataFrame()}
    out = pd.concat(frames, ignore_index=True)
    ctx.log(f"PDFExtract: {len(out)} 行")
    return {"dataframe": out}


@register_node
class AIFieldIdentify(BaseNode):
    __identifier__ = "excel_workflow.ai"
    NODE_NAME = "AI 字段识别"

    def __init__(self):
        super().__init__()
        self.set_color(236, 72, 153)
        self.add_input("dataframe")
        self.add_output("text")
        self.add_text_input("prompt", "额外说明", tab="参数")


@register_runner("excel_workflow.ai.AIFieldIdentify")
def run_ai_fields(node, inputs, ctx):
    df = inputs.get("dataframe")
    if df is None or isinstance(df, dict):
        raise ValueError("需要 DataFrame")
    cols = list(df.columns)[:40]
    sample = df.head(3).to_dict(orient="records")
    extra = (node.get_property("prompt") or "").strip()
    secrets = load_secrets()
    key = secrets.get("api_key") or os.environ.get("OPENAI_API_KEY")
    base = secrets.get("base_url") or os.environ.get("OPENAI_BASE_URL")
    if not key:
        ctx.log("AIFieldIdentify: 未配置 ~/.excel_workflow/secrets.json 的 api_key，跳过调用")
        summary = json.dumps({"columns": cols, "sample": sample}, ensure_ascii=False)
        return {"text": summary}
    try:
        from openai import OpenAI

        client = OpenAI(api_key=key, base_url=base or None)
        msg = (
            "根据以下列名与样本行，用中文简要说明每列可能含义，返回纯 JSON 数组"
            '[{"column":"","meaning":""}]\n'
            f"列: {cols}\n样本: {sample}\n{extra}"
        )
        r = client.chat.completions.create(
            model=secrets.get("model") or "gpt-4o-mini",
            messages=[{"role": "user", "content": msg}],
            temperature=0.2,
        )
        text = r.choices[0].message.content or ""
        ctx.log("AIFieldIdentify: 已调用模型")
        return {"text": text}
    except Exception as e:
        ctx.log(f"AIFieldIdentify: 调用失败 {e}")
        return {"text": json.dumps({"error": str(e)}, ensure_ascii=False)}


@register_node
class AIMappingSuggest(BaseNode):
    __identifier__ = "excel_workflow.ai"
    NODE_NAME = "AI 映射建议"

    def __init__(self):
        super().__init__()
        self.set_color(236, 72, 153)
        self.add_input("dataframe")
        self.add_output("text")
        self.add_text_input("cells_hint", "模板单元格说明", tab="参数")


@register_runner("excel_workflow.ai.AIMappingSuggest")
def run_ai_mapping(node, inputs, ctx):
    df = inputs.get("dataframe")
    if df is None or isinstance(df, dict):
        raise ValueError("需要 DataFrame")
    cols = list(df.columns)
    hint = (node.get_property("cells_hint") or "").strip()
    secrets = load_secrets()
    key = secrets.get("api_key") or os.environ.get("OPENAI_API_KEY")
    base = secrets.get("base_url") or os.environ.get("OPENAI_BASE_URL")
    if not key:
        ctx.log("AIMappingSuggest: 无 api_key，返回占位")
        return {"text": json.dumps([[c, ""] for c in cols], ensure_ascii=False)}
    try:
        from openai import OpenAI

        client = OpenAI(api_key=key, base_url=base or None)
        msg = (
            "数据源列名如下，请建议「列名 -> 模板单元格地址」映射，输出 JSON 二维数组"
            '例如 [["分项工程名称","Y6"]]，只输出 JSON。\n'
            f"列: {cols}\n模板说明: {hint}"
        )
        r = client.chat.completions.create(
            model=secrets.get("model") or "gpt-4o-mini",
            messages=[{"role": "user", "content": msg}],
            temperature=0.1,
        )
        text = r.choices[0].message.content or ""
        return {"text": text}
    except Exception as e:
        ctx.log(f"AIMappingSuggest: {e}")
        return {"text": json.dumps({"error": str(e)}, ensure_ascii=False)}
