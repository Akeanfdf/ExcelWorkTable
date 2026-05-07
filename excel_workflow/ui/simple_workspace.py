from __future__ import annotations

import glob
import os
import shutil
from datetime import datetime
from typing import TYPE_CHECKING, Callable, List, Optional

from PySide6.QtCore import Qt
from PySide6.QtWidgets import (
    QAbstractItemView,
    QDialog,
    QDialogButtonBox,
    QFileDialog,
    QFormLayout,
    QGroupBox,
    QHBoxLayout,
    QInputDialog,
    QLabel,
    QLineEdit,
    QListWidget,
    QListWidgetItem,
    QMessageBox,
    QPushButton,
    QTextEdit,
    QVBoxLayout,
    QWidget,
)

from excel_workflow.core.context import ExecContext
from excel_workflow.simple import template_store
from excel_workflow.simple.ops import run_registered
from excel_workflow.simple.session import CacheSession

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


class SimpleWorkspace(QWidget):
    """
    简易批量：导入文件夹 → 按钮执行各步骤 → 结果留在缓存目录 → 用户点导出复制到目标文件夹；
    可将本序列保存为模板下次一键重放。
    """

    def __init__(self, append_log: Callable[[str], None]):
        super().__init__()
        self._append_log = append_log
        self._session = CacheSession()
        self._folder = ""
        self._glob_pat = "*.xlsx"
        self._loaded_template_steps: Optional[List[dict]] = None

        root = QVBoxLayout(self)
        root.setSpacing(10)

        intro = QLabel(
            "流程：① 选文件夹并扫描 → ② 选中文件「载入工作区」→ ③ 点击下方按钮处理（结果写入缓存目录）"
            "→ ④ 确认无误后「导出到文件夹」→ ⑤ 可将本批操作「保存为模板」。"
        )
        intro.setWordWrap(True)
        root.addWidget(intro)

        imp = QGroupBox("1. 批量导入文件夹")
        il = QHBoxLayout(imp)
        self._folder_edit = QLineEdit()
        self._folder_edit.setPlaceholderText("选择要批量扫描的文件夹…")
        b_f = QPushButton("浏览文件夹…")
        b_f.clicked.connect(self._pick_folder)
        self._glob_edit = QLineEdit("*.xlsx")
        self._glob_edit.setMaximumWidth(120)
        b_scan = QPushButton("扫描文件")
        b_scan.clicked.connect(self._scan)
        il.addWidget(self._folder_edit, 1)
        il.addWidget(QLabel("通配符"))
        il.addWidget(self._glob_edit)
        il.addWidget(b_f)
        il.addWidget(b_scan)
        root.addWidget(imp)

        self._file_list = QListWidget()
        self._file_list.setSelectionMode(QAbstractItemView.ExtendedSelection)
        self._file_list.setMinimumHeight(120)
        root.addWidget(QLabel("扫描结果（可多选）"))
        root.addWidget(self._file_list, 1)

        row_load = QHBoxLayout()
        b_load = QPushButton("将选中 Excel 载入工作区（表）")
        b_load.clicked.connect(self._load_selected_excel)
        b_clear = QPushButton("清空缓存工作区")
        b_clear.clicked.connect(self._clear_workspace)
        row_load.addWidget(b_load)
        row_load.addWidget(b_clear)
        row_load.addStretch(1)
        root.addLayout(row_load)

        ops = QGroupBox("2. 处理（按钮 = 原节点图里的能力，结果进缓存区）")
        gl = QVBoxLayout(ops)
        row1 = QHBoxLayout()
        for text, slot in [
            ("格式标准化", self._op_format),
            ("去重", self._op_dedup),
            ("空值填充", self._op_fill),
            ("列重命名(JSON)", self._op_rename),
            ("行过滤(query)", self._op_filter),
        ]:
            b = QPushButton(text)
            b.clicked.connect(slot)
            row1.addWidget(b)
        gl.addLayout(row1)
        row2 = QHBoxLayout()
        for text, slot in [
            ("纵向合并另一张表…", self._op_merge),
            ("模板填充到缓存…", self._op_template),
            ("当前表写出到缓存", self._op_write_df),
            ("把缓存文件打成 ZIP", self._op_zip),
        ]:
            b = QPushButton(text)
            b.clicked.connect(slot)
            row2.addWidget(b)
        gl.addLayout(row2)
        root.addWidget(ops)

        cache = QGroupBox("3. 缓存工作区状态")
        cl = QVBoxLayout(cache)
        self._cache_info = QLabel("尚未载入表格。")
        self._cache_info.setWordWrap(True)
        self._artifact_list = QListWidget()
        self._artifact_list.setMaximumHeight(100)
        cl.addWidget(self._cache_info)
        cl.addWidget(QLabel("已生成的缓存文件："))
        cl.addWidget(self._artifact_list)
        root.addWidget(cache)

        out = QGroupBox("4. 导出与模板")
        ol = QHBoxLayout(out)
        b_export = QPushButton("导出到用户文件夹…")
        b_export.setStyleSheet("font-weight: 600;")
        b_export.clicked.connect(self._export_to_folder)
        b_save_tpl = QPushButton("保存本次按钮序列为模板")
        b_save_tpl.clicked.connect(self._save_template)
        b_load_tpl = QPushButton("加载模板…")
        b_load_tpl.clicked.connect(self._load_template_file)
        b_run_tpl = QPushButton("一键执行已加载模板")
        b_run_tpl.clicked.connect(self._run_loaded_template)
        ol.addWidget(b_export)
        ol.addWidget(b_save_tpl)
        ol.addWidget(b_load_tpl)
        ol.addWidget(b_run_tpl)
        root.addWidget(out)

        self._refresh_cache_ui()

    def _log(self, msg: str) -> None:
        self._append_log(f"[简易] {msg}")

    def _ctx(self) -> ExecContext:
        def log(m: str):
            self._log(m)

        def prog(_p: float, _m: str):
            pass

        return ExecContext(log=log, progress=prog)

    def _pick_folder(self):
        d = QFileDialog.getExistingDirectory(self, "选择文件夹", self._folder or "")
        if d:
            self._folder = d
            self._folder_edit.setText(d)

    def _scan(self):
        self._folder = self._folder_edit.text().strip()
        pat = self._glob_edit.text().strip() or "*"
        if not self._folder or not os.path.isdir(self._folder):
            QMessageBox.warning(self, "提示", "请先选择有效文件夹")
            return
        paths = sorted(
            glob.glob(os.path.join(self._folder, pat)), key=str.lower
        )
        paths = [p for p in paths if os.path.isfile(p)]
        self._file_list.clear()
        self._session.imported_files = paths
        for p in paths:
            self._file_list.addItem(QListWidgetItem(p))
        self._log(f"扫描到 {len(paths)} 个文件")

    def _selected_paths(self) -> List[str]:
        out = []
        for it in self._file_list.selectedItems():
            out.append(it.text())
        return out

    def _load_selected_excel(self):
        paths = self._selected_paths()
        if not paths and self._session.imported_files:
            paths = [self._session.imported_files[0]]
        if not paths:
            QMessageBox.warning(self, "提示", "请先扫描并选中至少一个 Excel")
            return
        path = paths[0]
        ctx = self._ctx()
        props = {"file_path": path, "sheet_name": ""}
        try:
            out = run_registered(T_READER, props, {}, ctx)
            self._session.dataframe = out.get("dataframe")
            self._remember(T_READER, props)
            self._log(f"已载入: {path}")
            self._refresh_cache_ui()
        except Exception as e:
            QMessageBox.critical(self, "错误", str(e))

    def _require_df(self):
        if self._session.dataframe is None:
            QMessageBox.warning(self, "提示", "请先将 Excel 载入工作区")
            return None
        return self._session.dataframe

    def _remember(self, type_id: str, props: dict) -> None:
        self._session.remember_op(type_id, props)

    def _apply_df(self, new_df, type_id: str, props: dict) -> None:
        self._session.dataframe = new_df
        self._remember(type_id, props)
        self._refresh_cache_ui()

    def _op_format(self):
        df = self._require_df()
        if df is None:
            return
        cols, ok = QInputDialog.getText(
            self, "格式标准化", "文本列名（逗号分隔，留空=不处理）:", text=""
        )
        if not ok:
            return
        props = {"columns": cols.strip(), "strip": True}
        ctx = self._ctx()
        try:
            out = run_registered(T_FMT, props, {"dataframe": df}, ctx)
            self._apply_df(out["dataframe"], T_FMT, props)
        except Exception as e:
            QMessageBox.critical(self, "错误", str(e))

    def _op_dedup(self):
        df = self._require_df()
        if df is None:
            return
        sub, ok = QInputDialog.getText(
            self, "去重", "按列去重（逗号分隔，留空=整行）:", text=""
        )
        if not ok:
            return
        props = {"subset": sub.strip(), "keep": True}
        ctx = self._ctx()
        try:
            out = run_registered(T_DEDUP, props, {"dataframe": df}, ctx)
            self._apply_df(out["dataframe"], T_DEDUP, props)
        except Exception as e:
            QMessageBox.critical(self, "错误", str(e))

    def _op_fill(self):
        df = self._require_df()
        if df is None:
            return
        val, ok = QInputDialog.getText(self, "空值填充", "填充值:", text="")
        if not ok:
            return
        cols, ok2 = QInputDialog.getText(
            self, "空值填充", "列名（逗号，留空=全部列）:", text=""
        )
        if not ok2:
            return
        props = {"fill_value": val, "columns": cols.strip()}
        ctx = self._ctx()
        try:
            out = run_registered(T_FILL, props, {"dataframe": df}, ctx)
            self._apply_df(out["dataframe"], T_FILL, props)
        except Exception as e:
            QMessageBox.critical(self, "错误", str(e))

    def _op_rename(self):
        df = self._require_df()
        if df is None:
            return
        j, ok = QInputDialog.getText(
            self,
            "列重命名",
            'JSON 对象，例如 {"旧列名":"新列名"}',
            text='{"旧列":"新列"}',
        )
        if not ok:
            return
        props = {"mapping_json": j.strip()}
        ctx = self._ctx()
        try:
            out = run_registered(T_RENAME, props, {"dataframe": df}, ctx)
            self._apply_df(out["dataframe"], T_RENAME, props)
        except Exception as e:
            QMessageBox.critical(self, "错误", str(e))

    def _op_filter(self):
        df = self._require_df()
        if df is None:
            return
        q, ok = QInputDialog.getText(
            self, "行过滤", "pandas query（留空=不过滤）:", text=""
        )
        if not ok:
            return
        props = {"query": q.strip()}
        ctx = self._ctx()
        try:
            out = run_registered(T_FILTER, props, {"dataframe": df}, ctx)
            self._apply_df(out["dataframe"], T_FILTER, props)
        except Exception as e:
            QMessageBox.critical(self, "错误", str(e))

    def _op_merge(self):
        df = self._require_df()
        if df is None:
            return
        path, _ = QFileDialog.getOpenFileName(
            self, "选择要纵向合并的另一张 Excel", "", "Excel (*.xlsx *.xls)"
        )
        if not path:
            return
        ctx = self._ctx()
        props_b = {"file_path": path, "sheet_name": ""}
        try:
            out_b = run_registered(T_READER, props_b, {}, ctx)
            bdf = out_b["dataframe"]
        except Exception as e:
            QMessageBox.critical(self, "错误", str(e))
            return
        props = {"ignore_index": True}
        try:
            out = run_registered(
                T_MERGE, props, {"left": df, "right": bdf}, ctx
            )
            self._remember(T_READER, props_b)
            self._apply_df(out["dataframe"], T_MERGE, props)
        except Exception as e:
            QMessageBox.critical(self, "错误", str(e))

    def _op_template(self):
        df = self._require_df()
        if df is None:
            return
        dlg = QDialog(self)
        dlg.setWindowTitle("模板填充（输出到缓存目录）")
        form = QFormLayout(dlg)
        t_path = QLineEdit()
        name_col = QLineEdit()
        sheet = QLineEdit("Sheet1")
        mapping = QTextEdit()
        mapping.setPlaceholderText('[["列名","A1"],["金额","B2"]]')
        mapping.setMinimumHeight(80)
        form.addRow("模板 xlsx", t_path)
        form.addRow("命名列", name_col)
        form.addRow("工作表名", sheet)
        form.addRow("映射 JSON", mapping)
        bb = QDialogButtonBox(
            QDialogButtonBox.Ok | QDialogButtonBox.Cancel
        )
        bb.accepted.connect(dlg.accept)
        bb.rejected.connect(dlg.reject)
        form.addRow(bb)
        if dlg.exec() != QDialog.DialogCode.Accepted:
            return
        props = {
            "template_path": t_path.text().strip(),
            "output_dir": self._session.staging_dir,
            "name_column": name_col.text().strip(),
            "sheet_name": sheet.text().strip() or "Sheet1",
            "mapping_json": mapping.toPlainText().strip() or "[]",
        }
        ctx = self._ctx()
        try:
            out = run_registered(T_TMPL, props, {"dataframe": df}, ctx)
            files = out.get("files") or []
            for f in files:
                self._session.add_artifact(f)
            tpl_props = dict(props)
            tpl_props["output_dir"] = "<STAGING>"
            self._remember(T_TMPL, tpl_props)
            self._log(f"模板已生成 {len(files)} 个文件到缓存")
            self._refresh_cache_ui()
        except Exception as e:
            QMessageBox.critical(self, "错误", str(e))

    def _op_write_df(self):
        df = self._require_df()
        if df is None:
            return
        name = datetime.now().strftime("export_%Y%m%d_%H%M%S.xlsx")
        out_path = os.path.join(self._session.staging_dir, name)
        props = {"output_path": out_path, "sheet_name": "Sheet1"}
        ctx = self._ctx()
        try:
            run_registered(T_WRITE, props, {"dataframe": df}, ctx)
            self._session.add_artifact(out_path)
            tpl_props = dict(props)
            tpl_props["output_path"] = "<STAGING>/" + name
            self._remember(T_WRITE, tpl_props)
            self._log(f"已写出: {out_path}")
            self._refresh_cache_ui()
        except Exception as e:
            QMessageBox.critical(self, "错误", str(e))

    def _op_zip(self):
        if not self._session.artifacts:
            QMessageBox.warning(self, "提示", "缓存区还没有可打包的文件")
            return
        name = datetime.now().strftime("export_%Y%m%d_%H%M%S.zip")
        zpath = os.path.join(self._session.staging_dir, name)
        props = {"zip_path": zpath}
        ctx = self._ctx()
        try:
            run_registered(T_ZIP, props, {"files": list(self._session.artifacts)}, ctx)
            self._session.add_artifact(zpath)
            tpl_props = {"zip_path": "<STAGING>/" + name}
            self._remember(T_ZIP, tpl_props)
            self._log(f"ZIP: {zpath}")
            self._refresh_cache_ui()
        except Exception as e:
            QMessageBox.critical(self, "错误", str(e))

    def _export_to_folder(self):
        if not self._session.artifacts:
            QMessageBox.warning(self, "提示", "缓存区没有文件，请先执行写出或模板填充等操作")
            return
        dest = QFileDialog.getExistingDirectory(self, "导出到文件夹", "")
        if not dest:
            return
        n = 0
        for src in self._session.artifacts:
            if os.path.isfile(src):
                shutil.copy2(src, os.path.join(dest, os.path.basename(src)))
                n += 1
        self._log(f"已导出 {n} 个文件到: {dest}")
        QMessageBox.information(self, "完成", f"已复制 {n} 个文件到:\n{dest}")

    def _save_template(self):
        if not self._session.op_history:
            QMessageBox.warning(self, "提示", "还没有执行过任何步骤，无法保存模板")
            return
        name, ok = QInputDialog.getText(self, "保存模板", "模板名称:", text="我的批量流程")
        if not ok or not name.strip():
            return
        glob_pat = self._glob_edit.text().strip() or "*.xlsx"
        try:
            p = template_store.save_template(
                name.strip(), glob_pat, list(self._session.op_history)
            )
            self._log(f"模板已保存: {p}")
            QMessageBox.information(self, "完成", f"已保存:\n{p}")
        except Exception as e:
            QMessageBox.critical(self, "错误", str(e))

    def _load_template_file(self):
        path, _ = QFileDialog.getOpenFileName(
            self, "加载模板", str(template_store.templates_root()), "JSON (*.json)"
        )
        if not path:
            return
        try:
            data = template_store.load_template(path)
            self._loaded_template_steps = data.get("steps") or []
            g = data.get("glob", "*.xlsx")
            self._glob_edit.setText(g)
            self._log(f"已加载模板「{data.get('name', path)}」共 {len(self._loaded_template_steps)} 步")
            QMessageBox.information(
                self, "模板",
                f"已加载 {len(self._loaded_template_steps)} 步。\n请先「载入工作区」当前表，再点「一键执行已加载模板」。",
            )
        except Exception as e:
            QMessageBox.critical(self, "错误", str(e))

    def _resolve_props(self, props: dict) -> dict:
        out = {}
        for k, v in props.items():
            if isinstance(v, str) and "<STAGING>" in v:
                out[k] = v.replace("<STAGING>", self._session.staging_dir).replace(
                    "/", os.sep
                )
            else:
                out[k] = v
        return out

    def _run_loaded_template(self):
        """Replay saved steps on current dataframe (reader steps use props paths)."""
        if not self._loaded_template_steps:
            QMessageBox.warning(self, "提示", "请先「加载模板」")
            return
        ctx = self._ctx()
        cur = self._session.dataframe
        if cur is None:
            QMessageBox.warning(self, "提示", "请先将 Excel 载入工作区")
            return
        try:
            for step in self._loaded_template_steps:
                tid = step["type_id"]
                props = self._resolve_props(step.get("props") or {})
                if tid == T_READER:
                    out = run_registered(tid, props, {}, ctx)
                    cur = out.get("dataframe")
                elif tid in (T_FMT, T_DEDUP, T_FILL, T_RENAME, T_FILTER):
                    out = run_registered(tid, props, {"dataframe": cur}, ctx)
                    cur = out.get("dataframe")
                elif tid == T_MERGE:
                    QMessageBox.warning(
                        self, "提示", "此模板含纵向合并，请在节点图或手动合并后再导出。"
                    )
                    return
                elif tid == T_TMPL:
                    out = run_registered(tid, props, {"dataframe": cur}, ctx)
                    for f in out.get("files") or []:
                        self._session.add_artifact(f)
                elif tid == T_WRITE:
                    props2 = dict(props)
                    outp = (props2.get("output_path") or "").strip()
                    if not outp:
                        continue
                    run_registered(tid, props2, {"dataframe": cur}, ctx)
                    self._session.add_artifact(os.path.abspath(outp))
                elif tid == T_ZIP:
                    zprops = dict(props)
                    zp = (zprops.get("zip_path") or "").strip()
                    if not zp:
                        continue
                    run_registered(
                        tid, zprops, {"files": list(self._session.artifacts)}, ctx
                    )
                    self._session.add_artifact(os.path.abspath(zp))
                else:
                    self._log(f"跳过未知步骤: {tid}")
            self._session.dataframe = cur
            self._log("模板步骤执行完毕")
            self._refresh_cache_ui()
        except Exception as e:
            QMessageBox.critical(self, "模板执行失败", str(e))

    def _clear_workspace(self):
        if QMessageBox.question(
            self, "确认", "清空缓存工作区？未导出的缓存文件将被删除。", QMessageBox.Yes | QMessageBox.No
        ) != QMessageBox.Yes:
            return
        self._session.reset()
        self._loaded_template_steps = None
        self._log("已清空缓存工作区")
        self._refresh_cache_ui()

    def _refresh_cache_ui(self):
        df = self._session.dataframe
        if df is None:
            self._cache_info.setText(
                f"缓存目录:\n{self._session.staging_dir}\n\n当前表:（无）"
            )
        else:
            try:
                r, c = len(df), len(df.columns)
                self._cache_info.setText(
                    f"缓存目录:\n{self._session.staging_dir}\n\n当前表: {r} 行 × {c} 列"
                )
            except Exception:
                self._cache_info.setText(
                    f"缓存目录:\n{self._session.staging_dir}\n\n当前表:（已载入）"
                )
        self._artifact_list.clear()
        for p in self._session.artifacts:
            self._artifact_list.addItem(QListWidgetItem(p))
