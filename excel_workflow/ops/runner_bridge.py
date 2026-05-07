"""
将线性流程的 op_type + params 映射到已注册的节点 runner（simple.ops.run_registered）。
"""

from __future__ import annotations

import json
import os
from datetime import datetime
from pathlib import Path
from typing import Any, Dict, List, TYPE_CHECKING

import pandas as pd

from excel_workflow.core.context import ExecContext
from excel_workflow.session.state import AppliedOp, LinearSession
from excel_workflow.simple.ops import run_registered

if TYPE_CHECKING:
    pass

T_READER = "excel_workflow.source.ExcelReader"
T_FMT = "excel_workflow.clean.FormatStandardize"
T_DEDUP = "excel_workflow.clean.Deduplicate"
T_FILL = "excel_workflow.clean.FillEmpty"
T_RENAME = "excel_workflow.clean.RenameColumns"
T_FILTER = "excel_workflow.filter.FilterRows"
T_MERGE = "excel_workflow.merge.VerticalMerge"
T_TMPL = "excel_workflow.template.TemplateMapper"
T_WRITE = "excel_workflow.export.ExcelWriter"
T_ZIP = "excel_workflow.export.ZipPack"


def _is_tabular_path(path: Path) -> bool:
    return path.suffix.lower() in (".xlsx", ".xls", ".csv")


def tabular_source_paths(session: LinearSession) -> List[Path]:
    return [p for p in session.source_files if _is_tabular_path(p)]


def _parse_csv_cols(s: str) -> List[str]:
    return [c.strip() for c in (s or "").split(",") if c.strip()]


_OP_TITLE_ZH = {
    "format_std": "格式标准化",
    "dedup": "去重",
    "fill_empty": "空值填充",
    "rename_cols": "列重命名",
    "filter_rows": "筛选行",
    "sort": "排序",
    "column_pick": "选列",
    "replace": "查找替换",
    "merge_vertical": "纵向合并",
    "template_fill": "填充模板",
    "write_xlsx": "写出 xlsx",
    "zip_pack": "打包 ZIP",
}


def _op_title_zh(op_type: str) -> str:
    return _OP_TITLE_ZH.get(op_type, op_type)


def missing_columns_for_op(op: AppliedOp, colset: set) -> List[str]:
    p = op.params
    t = op.op_type
    missing: List[str] = []
    if t == "format_std":
        raw = (p.get("columns") or "").strip()
        if raw:
            for c in _parse_csv_cols(raw):
                if c not in colset:
                    missing.append(c)
    elif t == "dedup":
        raw = (p.get("subset") or "").strip()
        for c in _parse_csv_cols(raw):
            if c not in colset:
                missing.append(c)
    elif t == "fill_empty":
        raw = (p.get("columns") or "").strip()
        if raw:
            for c in _parse_csv_cols(raw):
                if c not in colset:
                    missing.append(c)
    elif t == "rename_cols":
        try:
            m = json.loads((p.get("mapping_json") or "{}").strip() or "{}")
            if isinstance(m, dict):
                for k in m.keys():
                    ks = str(k)
                    if ks and ks not in colset:
                        missing.append(ks)
        except (json.JSONDecodeError, TypeError):
            pass
    elif t == "sort":
        c = (p.get("column") or "").strip()
        if c and c not in colset:
            missing.append(c)
    elif t == "column_pick":
        for c in _parse_csv_cols(p.get("columns") or ""):
            if c not in colset:
                missing.append(c)
    elif t == "replace":
        c = (p.get("column") or "").strip()
        if c and c not in colset:
            missing.append(c)
    elif t == "template_fill":
        c = (p.get("name_column") or "").strip()
        if c and c not in colset:
            missing.append(c)
    elif t == "filter_rows":
        q = (p.get("query") or "").strip()
        if not q:
            return missing
        if not colset:
            missing.append("（该表无任何列）")
            return missing
        try:
            probe = pd.DataFrame({c: [1] for c in colset})
            probe.query(q)
        except Exception as ex:
            missing.append(f"筛选条件在该表上校验失败（列名或写法与数据不一致）：{ex}")
    return missing


def validate_ops_across_all_sources(session: LinearSession, ctx: ExecContext) -> None:
    paths = tabular_source_paths(session)
    if len(paths) < 2:
        return
    ops = list(session.ops_history)
    if not ops:
        return
    primary = paths[0]
    problems: List[str] = []
    for path in paths:
        load_table_from_path(session, path, ctx)
        df = session.dataframe
        if df is None:
            continue
        colset = set(df.columns.astype(str))
        for op in ops:
            miss = missing_columns_for_op(op, colset)
            if miss:
                step = _op_title_zh(op.op_type)
                problems.append(
                    f"「{path.name}」· 步骤「{step}」({op.op_type})：{'; '.join(miss)}"
                )
    load_table_from_path(session, primary, ctx)
    if problems:
        raise ValueError(
            "批量处理已取消。\n\n"
            "【原因】已载入多张表时，会对每张表依次重放「步骤流程」中的全部操作；"
            "因此每张表都必须具备各步骤所需的列名（与主表一致），且<strong>筛选行</strong>步骤里填写的"
            "条件式在<strong>每一张表</strong>上都必须能正确计算（列名一致、写法合法）。\n"
            "若只有部分表缺少某列，或条件式只在主表上成立，则无法继续批量。\n\n"
            "逐表检查结果：\n\n" + "\n".join(problems)
        )


def load_table_from_path(
    session: LinearSession, path: Path, ctx: ExecContext
) -> str:
    p = str(path)
    ext = os.path.splitext(p)[1].lower()
    if ext == ".csv":
        session.dataframe = pd.read_csv(p)
        return f"已载入 CSV: {os.path.basename(p)}"
    props = {"file_path": p, "sheet_name": ""}
    out = run_registered(T_READER, props, {}, ctx)
    session.dataframe = out.get("dataframe")
    return f"已载入: {os.path.basename(p)}"


def load_primary_table(session: LinearSession, ctx: ExecContext) -> str:
    if not session.source_files:
        raise ValueError("没有源文件")
    return load_table_from_path(session, session.source_files[0], ctx)


def batch_apply_ops_to_all_sources(
    session: LinearSession, ctx: ExecContext
) -> str:
    ops = list(session.ops_history)
    if not ops:
        raise ValueError("还没有已应用的操作步骤")
    paths = tabular_source_paths(session)
    if len(paths) < 2:
        raise ValueError("请至少载入 2 个表格文件后再使用批量处理")
    validate_ops_across_all_sources(session, ctx)
    order = paths[1:] + [paths[0]]
    for path in order:
        msg = load_table_from_path(session, path, ctx)
        ctx.log(f"[批量] {msg}")
        for op in ops:
            apply_linear_op(session, op.op_type, dict(op.params), ctx)
    names = "、".join(p.name for p in paths)
    return f"已对 {len(paths)} 个文件各执行 {len(ops)} 步（{names}）"


def apply_linear_op(
    session: LinearSession, op_type: str, params: Dict[str, Any], ctx: ExecContext
) -> str:
    st = session.staging_dir_str()
    if not st:
        raise RuntimeError("未初始化 staging 目录")

    df = session.dataframe
    if df is None and op_type not in ("zip_pack",):
        raise RuntimeError("当前没有表格，请先完成导入")

    if op_type == "format_std":
        props = {
            "columns": (params.get("columns") or "").strip(),
            "strip": bool(params.get("strip", True)),
        }
        out = run_registered(T_FMT, props, {"dataframe": df}, ctx)
        session.dataframe = out["dataframe"]
        return "格式标准化"

    if op_type == "dedup":
        props = {
            "subset": (params.get("subset") or "").strip(),
            "keep": bool(params.get("keep", True)),
        }
        out = run_registered(T_DEDUP, props, {"dataframe": df}, ctx)
        session.dataframe = out["dataframe"]
        return f"去重 ({props['subset'] or '整行'})"

    if op_type == "fill_empty":
        props = {
            "fill_value": params.get("fill_value", ""),
            "columns": (params.get("columns") or "").strip(),
        }
        out = run_registered(T_FILL, props, {"dataframe": df}, ctx)
        session.dataframe = out["dataframe"]
        return "空值填充"

    if op_type == "rename_cols":
        props = {"mapping_json": (params.get("mapping_json") or "{}").strip()}
        out = run_registered(T_RENAME, props, {"dataframe": df}, ctx)
        session.dataframe = out["dataframe"]
        return "列重命名"

    if op_type == "filter_rows":
        props = {"query": (params.get("query") or "").strip()}
        out = run_registered(T_FILTER, props, {"dataframe": df}, ctx)
        session.dataframe = out["dataframe"]
        q = props["query"] or "（未筛选）"
        return f"筛选行: {q[:40]}"

    if op_type == "sort":
        col = (params.get("column") or "").strip()
        if not col:
            raise ValueError("请填写用于排序的列名")
        if col not in df.columns:
            raise ValueError(f"列不存在: {col}")
        ascending = bool(params.get("ascending", True))
        session.dataframe = df.sort_values(by=col, ascending=ascending).reset_index(
            drop=True
        )
        ctx.log(f"排序: {col}, ascending={ascending}")
        return f"排序 · {col}"

    if op_type == "column_pick":
        raw = (params.get("columns") or "").strip()
        if not raw:
            raise ValueError("请填写要保留的列名（逗号分隔）")
        cols = [c.strip() for c in raw.split(",") if c.strip()]
        missing = [c for c in cols if c not in df.columns]
        if missing:
            raise ValueError(f"列不存在: {missing}")
        session.dataframe = df.loc[:, cols].copy()
        ctx.log(f"选列: 保留 {len(cols)} 列")
        return f"选列 · {len(cols)} 列"

    if op_type == "replace":
        col = (params.get("column") or "").strip()
        find = params.get("find", "")
        repl = params.get("replace", "")
        if not col:
            raise ValueError("请填写列名")
        if col not in df.columns:
            raise ValueError(f"列不存在: {col}")
        out_df = df.copy()
        out_df[col] = out_df[col].astype(str).str.replace(
            str(find), str(repl), regex=False
        )
        session.dataframe = out_df
        ctx.log(f"查找替换: 列={col}")
        return f"替换 · {col}"

    if op_type == "merge_vertical":
        path2 = (params.get("second_file") or "").strip()
        if not path2 or not os.path.isfile(path2):
            raise ValueError("请选择有效的第二个 Excel 文件")
        props_b = {"file_path": path2, "sheet_name": ""}
        out_b = run_registered(T_READER, props_b, {}, ctx)
        bdf = out_b["dataframe"]
        props = {"ignore_index": bool(params.get("ignore_index", True))}
        out = run_registered(T_MERGE, props, {"left": df, "right": bdf}, ctx)
        session.dataframe = out["dataframe"]
        return f"纵向合并: {os.path.basename(path2)}"

    if op_type == "template_fill":
        props = {
            "template_path": (params.get("template_path") or "").strip(),
            "output_dir": st,
            "name_column": (params.get("name_column") or "").strip(),
            "sheet_name": (params.get("sheet_name") or "Sheet1").strip() or "Sheet1",
            "mapping_json": (params.get("mapping_json") or "[]").strip() or "[]",
        }
        if not props["template_path"]:
            raise ValueError("请填写模板路径")
        out = run_registered(T_TMPL, props, {"dataframe": df}, ctx)
        for f in out.get("files") or []:
            session.add_artifact(str(f))
        n = len(out.get("files") or [])
        return f"模板填充 → {n} 个文件"

    if op_type == "write_xlsx":
        name = (params.get("filename") or "").strip()
        if not name:
            name = datetime.now().strftime("export_%Y%m%d_%H%M%S.xlsx")
        if not name.lower().endswith(".xlsx"):
            name += ".xlsx"
        out_path = os.path.join(st, name)
        props = {
            "output_path": out_path,
            "sheet_name": (params.get("sheet_name") or "Sheet1"),
        }
        run_registered(T_WRITE, props, {"dataframe": df}, ctx)
        session.add_artifact(out_path)
        return f"写出: {name}"

    if op_type == "zip_pack":
        if not session.artifacts:
            raise ValueError("没有可打包的缓存文件")
        name = datetime.now().strftime("export_%Y%m%d_%H%M%S.zip")
        zpath = os.path.join(st, name)
        props = {"zip_path": zpath}
        run_registered(T_ZIP, props, {"files": list(session.artifacts)}, ctx)
        session.add_artifact(zpath)
        return f"ZIP: {name}"

    raise ValueError(f"未知或未实现的操作: {op_type}")


def replay_ops(session: LinearSession, ops: List[AppliedOp], ctx: ExecContext) -> None:
    session.reset_staging_keep_sources()
    summary = load_primary_table(session, ctx)
    ctx.log(summary)
    for op in ops:
        s = apply_linear_op(session, op.op_type, op.params, ctx)
        session.ops_history.append(
            AppliedOp(
                op_type=op.op_type,
                params=dict(op.params),
                timestamp=op.timestamp,
                result_summary=s,
            )
        )
    session.current_files = list(session.source_files)
    session.current_files.extend(Path(a) for a in session.artifacts)
