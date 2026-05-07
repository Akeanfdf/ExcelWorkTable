from __future__ import annotations

import os
import shutil
import zipfile
from typing import Any, Dict, List

import pandas as pd
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
class ExcelWriter(BaseNode):
    __identifier__ = "excel_workflow.export"
    NODE_NAME = "写出 Excel"

    def __init__(self):
        super().__init__()
        self.set_color(249, 115, 22)
        self.add_input("dataframe")
        self.add_output("path")
        self.add_text_input("output_path", "输出 .xlsx 路径", tab="参数")
        self.add_text_input("sheet_name", "工作表名 Sheet1", tab="参数")


@register_runner("excel_workflow.export.ExcelWriter")
def run_excel_writer(node, inputs, ctx):
    df = inputs.get("dataframe")
    if df is None or isinstance(df, dict):
        raise ValueError("ExcelWriter: 需要 DataFrame")
    path = (node.get_property("output_path") or "").strip()
    sname = (node.get_property("sheet_name") or "Sheet1").strip() or "Sheet1"
    if not path:
        raise ValueError("ExcelWriter: 未设置输出路径")
    os.makedirs(os.path.dirname(os.path.abspath(path)) or ".", exist_ok=True)
    with pd.ExcelWriter(path, engine="openpyxl") as w:
        df.to_excel(w, sheet_name=sname, index=False)
    ctx.log(f"ExcelWriter: {path}")
    return {"path": path}


@register_node
class PDFExport(BaseNode):
    __identifier__ = "excel_workflow.export"
    NODE_NAME = "导出 PDF"

    def __init__(self):
        super().__init__()
        self.set_color(249, 115, 22)
        self.add_input("files")
        self.add_output("files")
        self.add_checkbox("use_com", "", "使用 Excel COM (Windows)", True)


@register_runner("excel_workflow.export.PDFExport")
def run_pdf_export(node, inputs, ctx):
    files = _as_file_list(inputs.get("files"))
    if not files:
        raise ValueError("PDFExport: 无输入文件")
    use_com = bool(node.get_property("use_com"))
    out_pdfs: List[str] = []
    if use_com:
        try:
            import win32com.client  # type: ignore
        except ImportError:
            ctx.log("PDFExport: 未安装 pywin32，跳过 COM 导出")
            use_com = False
    if use_com:
        excel = win32com.client.Dispatch("Excel.Application")
        excel.Visible = False
        try:
            for xlsx in files:
                if not xlsx.lower().endswith((".xlsx", ".xls")):
                    continue
                pdf = os.path.splitext(xlsx)[0] + ".pdf"
                ab = os.path.abspath(xlsx)
                wb = excel.Workbooks.Open(ab)
                wb.ExportAsFixedFormat(0, os.path.abspath(pdf))
                wb.Close(False)
                out_pdfs.append(pdf)
                ctx.log(f"PDFExport: {pdf}")
        finally:
            excel.Quit()
    else:
        ctx.log("PDFExport: COM 关闭，未生成 PDF")
    return {"files": out_pdfs or files}


@register_node
class ZipPack(BaseNode):
    __identifier__ = "excel_workflow.export"
    NODE_NAME = "ZIP 打包"

    def __init__(self):
        super().__init__()
        self.set_color(249, 115, 22)
        self.add_input("files")
        self.add_output("archive")
        self.add_text_input("zip_path", "输出 .zip 路径", tab="参数")


@register_runner("excel_workflow.export.ZipPack")
def run_zip_pack(node, inputs, ctx):
    files = _as_file_list(inputs.get("files"))
    zpath = (node.get_property("zip_path") or "").strip()
    if not files:
        raise ValueError("ZipPack: 无文件")
    if not zpath:
        raise ValueError("ZipPack: 未设置 zip_path")
    parent = os.path.dirname(os.path.abspath(zpath))
    if parent:
        os.makedirs(parent, exist_ok=True)
    with zipfile.ZipFile(zpath, "w", zipfile.ZIP_DEFLATED) as zf:
        for p in files:
            if os.path.isfile(p):
                zf.write(p, arcname=os.path.basename(p))
    ctx.log(f"ZipPack: {zpath} ({len(files)} 项)")
    return {"archive": zpath}


@register_node
class PrintToPrinter(BaseNode):
    __identifier__ = "excel_workflow.export"
    NODE_NAME = "打印"

    def __init__(self):
        super().__init__()
        self.set_color(249, 115, 22)
        self.add_input("files")
        self.add_output("files")
        self.add_text_input("printer_name", "打印机名(空=默认)", tab="参数")


@register_runner("excel_workflow.export.PrintToPrinter")
def run_print(node, inputs, ctx):
    files = _as_file_list(inputs.get("files"))
    try:
        import win32com.client  # type: ignore
    except ImportError:
        ctx.log("Print: 需要 pywin32")
        return {"files": files}
    printer = (node.get_property("printer_name") or "").strip() or None
    excel = win32com.client.Dispatch("Excel.Application")
    excel.Visible = False
    try:
        for xlsx in files:
            if not xlsx.lower().endswith((".xlsx", ".xls")):
                continue
            wb = excel.Workbooks.Open(os.path.abspath(xlsx))
            if printer:
                wb.PrintOut(ActivePrinter=printer)
            else:
                wb.PrintOut()
            wb.Close(False)
            ctx.log(f"Print: {xlsx}")
    finally:
        excel.Quit()
    return {"files": files}
