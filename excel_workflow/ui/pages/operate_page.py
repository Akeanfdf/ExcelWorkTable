"""线性流程 — 操作阶段：选功能、填参数、时间线、追加导入与 workflow 重放。"""

from __future__ import annotations

import json
import os
from pathlib import Path
from typing import Any, Callable, Dict, List, Optional

from PySide6.QtCore import Signal, Qt
from PySide6.QtWidgets import (
    QCheckBox,
    QComboBox,
    QFileDialog,
    QFormLayout,
    QFrame,
    QGroupBox,
    QHBoxLayout,
    QLabel,
    QLineEdit,
    QMessageBox,
    QPushButton,
    QScrollArea,
    QTextEdit,
    QVBoxLayout,
    QWidget,
)

from excel_workflow.core.context import ExecContext
from excel_workflow.ops.runner_bridge import (
    apply_linear_op,
    batch_apply_ops_to_all_sources,
    load_primary_table,
    replay_ops,
    tabular_source_paths,
)
from excel_workflow.session.state import AppliedOp, LinearSession
from excel_workflow.session.workflow_io import load_workflow_v2
from excel_workflow.ui import linear_feature_docs as docs
from excel_workflow.ui.linear_manual_dialog import show_full_manual
from excel_workflow.ui.widgets.column_widgets import (
    attach_column_completer,
    column_hint_label,
    column_options_from_dataframe,
    combo_column_text,
    make_column_combo,
    make_column_list_picker,
)
from excel_workflow.ui.widgets.data_preview import DataPreviewPanel
from excel_workflow.ui.widgets.file_list_card import FileListCard
from excel_workflow.ui.widgets.op_button_panel import OpButtonPanel
from excel_workflow.ui.widgets.op_timeline import OpTimeline


def _normalize_template_fill_params(p: Dict[str, Any]) -> Dict[str, Any]:
    mj = p.get("mapping_json")
    if isinstance(mj, (list, dict)):
        mj = json.dumps(mj, ensure_ascii=False)
    else:
        mj = (str(mj) if mj is not None else "").strip() or "[]"
    return {
        "template_path": str(
            p.get("template_path") or p.get("tmpl_file") or ""
        ).strip(),
        "name_column": str(p.get("name_column") or p.get("name_col") or "").strip(),
        "sheet_name": str(p.get("sheet_name") or "Sheet1").strip() or "Sheet1",
        "mapping_json": mj,
    }


class OperatePage(QWidget):
    """操作页：左侧功能、参数区、时间线、预览。"""

    ops_changed = Signal()

    def __init__(
        self,
        append_log: Callable[[str], None],
        on_back_to_import: Optional[Callable[[], None]] = None,
        parent=None,
    ):
        super().__init__(parent)
        self._append_log = append_log
        self._on_back_to_import = on_back_to_import
        self._session: Optional[LinearSession] = None
        self._current_op: Optional[str] = None
        self._current_label = ""
        self._fields: Dict[str, Any] = {}
        self._column_hint: Optional[QLabel] = None
        self._list_collect: Optional[Callable[[], str]] = None
        self._fmt_columns_collect: Optional[Callable[[], str]] = None
        self._dedup_subset_collect: Optional[Callable[[], str]] = None
        self._fill_columns_collect: Optional[Callable[[], str]] = None

        # 左侧固定按钮栏 + 右侧整栏共用一个纵向滚动条（不再给各块单独套 QScrollArea）
        root = QHBoxLayout(self)
        root.setContentsMargins(0, 0, 0, 0)
        root.setSpacing(10)

        self._op_panel = OpButtonPanel(self._on_op_clicked)
        root.addWidget(self._op_panel)

        right_scroll = QScrollArea()
        right_scroll.setObjectName("OperateRightScroll")
        right_scroll.setWidgetResizable(True)
        right_scroll.setFrameShape(QFrame.Shape.NoFrame)
        right_scroll.setHorizontalScrollBarPolicy(
            Qt.ScrollBarPolicy.ScrollBarAsNeeded
        )
        right_scroll.setVerticalScrollBarPolicy(
            Qt.ScrollBarPolicy.ScrollBarAsNeeded
        )

        right_inner = QWidget()
        self._mid_layout = QVBoxLayout(right_inner)
        self._mid_layout.setContentsMargins(0, 0, 0, 0)
        self._mid_layout.setSpacing(10)

        tool_row = QHBoxLayout()
        b_json = QPushButton("导入 workflow.json…")
        b_json.setToolTip(docs.tooltip("import_workflow_json"))
        b_json.clicked.connect(self._import_workflow_json)
        b_batch = QPushButton("对全部表格执行当前步骤")
        b_batch.setToolTip(docs.tooltip("batch_all_sources"))
        b_batch.clicked.connect(self._run_batch_all)
        b_man = QPushButton("全文说明书…")
        b_man.clicked.connect(lambda: show_full_manual(self.window()))
        tool_row.addWidget(b_json)
        tool_row.addWidget(b_batch)
        if on_back_to_import:
            b_imp = QPushButton("返回导入")
            b_imp.setToolTip(docs.tooltip("back_to_import_step"))
            b_imp.clicked.connect(on_back_to_import)
            tool_row.addWidget(b_imp)
        tool_row.addStretch(1)
        tool_row.addWidget(b_man)
        self._mid_layout.addLayout(tool_row)

        self._file_card = FileListCard(
            on_append=self._append_import,
            on_remove_selected=self._on_remove_sources,
        )
        self._mid_layout.addWidget(self._file_card)

        self._timeline = OpTimeline(self._delete_at)
        steps_box = QGroupBox("步骤流程")
        steps_box.setObjectName("LinearCard")
        steps_lay = QVBoxLayout(steps_box)
        steps_lay.setContentsMargins(8, 10, 8, 8)
        steps_lay.addWidget(self._timeline)
        self._mid_layout.addWidget(steps_box)

        self._op_title = QLabel("请从左侧选择一项操作")
        self._op_title.setStyleSheet("font-weight:600;font-size:14px;")
        self._mid_layout.addWidget(self._op_title)

        self._param_host = QFrame()
        self._param_host.setObjectName("LinearCard")
        self._param_layout = QFormLayout(self._param_host)
        self._mid_layout.addWidget(self._param_host)

        btn_row = QHBoxLayout()
        self._btn_apply = QPushButton("应用")
        self._btn_apply.setToolTip(docs.tooltip("param_apply"))
        self._btn_apply.clicked.connect(self._apply)
        self._btn_cancel = QPushButton("取消草稿")
        self._btn_cancel.setToolTip(docs.tooltip("param_cancel"))
        self._btn_cancel.clicked.connect(self._cancel_draft)
        btn_row.addWidget(self._btn_apply)
        btn_row.addWidget(self._btn_cancel)
        btn_row.addStretch(1)
        self._mid_layout.addLayout(btn_row)

        self._preview = DataPreviewPanel()
        self._mid_layout.addWidget(self._preview)

        right_scroll.setWidget(right_inner)
        root.addWidget(right_scroll, 1)

    def _ctx(self) -> ExecContext:
        def log(m: str) -> None:
            self._append_log(f"[线性] {m}")

        return ExecContext(log=log, progress=lambda _p, _m: None)

    def set_session(self, session: LinearSession) -> None:
        self._session = session
        self._file_card.set_files(list(session.source_files))
        self._preview.set_dataframe(session.dataframe)
        self._refresh_timeline()
        self._current_op = None
        self._op_title.setText("请从左侧选择一项操作")
        self._clear_param_form()

    def apply_mapping_prefill(self, payload: Dict[str, Any]) -> None:
        if not self._session:
            return
        cfg = payload.get("legacy_mapping_config")
        if cfg is not None:
            self._session.legacy_mapping_config = cfg
        ops = payload.get("prefill_ops") or []
        if not ops:
            return
        first = ops[0]
        op_type = str(first.get("op_type", ""))
        params = dict(first.get("params") or {})
        if op_type == "template_fill":
            params = _normalize_template_fill_params(params)
        self._on_op_clicked(op_type, "填充模板")
        self._prefill_fields(op_type, params)

    def _append_import(self) -> None:
        if not self._session:
            return
        paths, _ = QFileDialog.getOpenFileNames(
            self,
            "追加表格",
            "",
            "表格 (*.xlsx *.xls *.csv)",
        )
        if not paths:
            return
        for p in paths:
            self._session.source_files.append(Path(os.path.abspath(p)))
        self._file_card.set_files(list(self._session.source_files))
        try:
            msg = load_primary_table(self._session, self._ctx())
            self._append_log(msg)
        except Exception as e:
            QMessageBox.critical(self, "载入主表失败", str(e))
            return
        self._preview.set_dataframe(self._session.dataframe)
        self._session.is_dirty = True

    @staticmethod
    def _path_key(p: Path) -> str:
        try:
            return os.path.normcase(str(p.resolve()))
        except OSError:
            return os.path.normcase(os.path.abspath(str(p)))

    def _on_remove_sources(self, paths: List[Path]) -> None:
        if not self._session:
            return
        to_remove = {self._path_key(p) for p in paths}
        show_names = "、".join(p.name for p in paths[:8])
        if len(paths) > 8:
            show_names += f" 等共 {len(paths)} 个"
        r = QMessageBox.question(
            self,
            "移除已载入文件",
            f"从当前会话中移除下列文件？（不会删除磁盘上的文件）\n\n{show_names}",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
            QMessageBox.StandardButton.No,
        )
        if r != QMessageBox.StandardButton.Yes:
            return

        old_sources = list(self._session.source_files)
        if not old_sources:
            return
        primary_key = self._path_key(old_sources[0])
        primary_in_remove = primary_key in to_remove
        saved_ops = list(self._session.ops_history)
        new_sources = [p for p in old_sources if self._path_key(p) not in to_remove]

        if not new_sources:
            if self._on_back_to_import:
                if (
                    QMessageBox.question(
                        self,
                        "返回导入",
                        "移除后没有剩余表格，将清空流程并回到「导入」界面。确定？",
                        QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
                        QMessageBox.StandardButton.No,
                    )
                    == QMessageBox.StandardButton.Yes
                ):
                    self._on_back_to_import()
            else:
                QMessageBox.warning(self, "提示", "至少保留一个已载入文件。")
            return

        self._session.source_files = new_sources
        if primary_in_remove:
            try:
                replay_ops(self._session, saved_ops, self._ctx())
            except Exception as e:
                QMessageBox.critical(self, "切换主表后重放失败", str(e))
                return
        self._file_card.set_files(list(self._session.source_files))
        self._preview.set_dataframe(self._session.dataframe)
        self._refresh_timeline()
        if self._current_op == "merge_vertical" and "second_cb" in self._fields:
            self._on_op_clicked("merge_vertical", "合并文件")
        self._append_log(f"[线性] 已移除 {len(paths)} 个文件")
        self.ops_changed.emit()

    def _import_workflow_json(self) -> None:
        if not self._session:
            return
        path, _ = QFileDialog.getOpenFileName(
            self, "选择 workflow.json", "", "JSON (*.json)"
        )
        if not path:
            return
        try:
            data = load_workflow_v2(path)
        except Exception as e:
            QMessageBox.critical(self, "读取失败", str(e))
            return
        legacy = data.get("legacy_mapping_config")
        if legacy is not None:
            self._session.legacy_mapping_config = legacy
        raw_ops = data.get("ops") or []
        try:
            ops = LinearSession.from_ops_list(raw_ops)
            replay_ops(self._session, ops, self._ctx())
        except Exception as e:
            QMessageBox.critical(self, "重放失败", str(e))
            return
        self._append_log(f"[线性] 已导入并重放 workflow: {path}")
        self._preview.set_dataframe(self._session.dataframe)
        self._refresh_timeline()
        self.ops_changed.emit()
        QMessageBox.information(self, "完成", f"已载入并重放 {len(ops)} 步。")

    def _run_batch_all(self) -> None:
        if not self._session:
            return
        try:
            msg = batch_apply_ops_to_all_sources(self._session, self._ctx())
            self._append_log(msg)
            load_primary_table(self._session, self._ctx())
        except Exception as e:
            QMessageBox.warning(self, "批量处理", str(e))
            return
        self._preview.set_dataframe(self._session.dataframe)
        self._refresh_timeline()
        self.ops_changed.emit()

    def _on_op_clicked(self, op_key: str, label: str) -> None:
        self._current_op = op_key
        self._current_label = label
        self._op_title.setText(f"当前：{label}")
        self._build_param_form(op_key)

    def _cancel_draft(self) -> None:
        self._current_op = None
        self._op_title.setText("请从左侧选择一项操作")
        self._clear_param_form()

    def _clear_param_form(self) -> None:
        while self._param_layout.rowCount():
            self._param_layout.removeRow(0)
        self._fields.clear()
        self._list_collect = None
        self._fmt_columns_collect = None
        self._dedup_subset_collect = None
        self._fill_columns_collect = None
        self._column_hint = None

    def _names(self) -> List[str]:
        if not self._session:
            return []
        return column_options_from_dataframe(self._session.dataframe)

    def _add_hint_row(self) -> None:
        self._column_hint = column_hint_label(self._names())
        self._param_layout.addRow(self._column_hint)

    def _build_param_form(self, op: str) -> None:
        self._clear_param_form()
        self._add_hint_row()
        names = self._names()

        if op == "format_std":
            self._param_layout.addRow(
                QLabel("在下面勾选要标准化的列；不勾选表示<strong>全部列</strong>。")
            )
            lw, collect = make_column_list_picker(names, "", max_height=180)
            self._fields["columns_list"] = lw
            self._fmt_columns_collect = collect
            self._param_layout.addRow("列（多选）", lw)
            self._fields["strip"] = QCheckBox("去除首尾空白")
            self._fields["strip"].setChecked(True)
            self._param_layout.addRow("", self._fields["strip"])
        elif op == "dedup":
            self._param_layout.addRow(
                QLabel("勾选参与去重的列；<strong>不勾选</strong>表示按<strong>整行</strong>去重。")
            )
            lw, collect = make_column_list_picker(names, "", max_height=180)
            self._fields["dedup_subset_list"] = lw
            self._dedup_subset_collect = collect
            self._param_layout.addRow("依据列（多选）", lw)
            self._fields["keep"] = QCheckBox("保留首次出现")
            self._fields["keep"].setChecked(True)
            self._param_layout.addRow("", self._fields["keep"])
        elif op == "fill_empty":
            self._fields["fill_value"] = QLineEdit()
            self._param_layout.addRow("填充值", self._fields["fill_value"])
            self._param_layout.addRow(
                QLabel("勾选要填充的列；不勾选表示<strong>全部列</strong>。")
            )
            lw, collect = make_column_list_picker(names, "", max_height=180)
            self._fields["fill_columns_list"] = lw
            self._fill_columns_collect = collect
            self._param_layout.addRow("列（多选）", lw)
        elif op == "rename_cols":
            te = QTextEdit()
            te.setPlaceholderText('{"旧列名":"新列名"}')
            te.setMaximumHeight(100)
            self._fields["mapping_json"] = te
            self._param_layout.addRow("列映射 JSON", te)
            row_r = QWidget()
            hr = QHBoxLayout(row_r)
            hr.setContentsMargins(0, 0, 0, 0)
            old_cb = make_column_combo(names)
            new_le = QLineEdit()
            new_le.setPlaceholderText("新列名")
            btn_add = QPushButton("加入映射")

            def _add_rename() -> None:
                key = combo_column_text(old_cb).strip()
                val = new_le.text().strip()
                if not key:
                    return
                try:
                    m = json.loads(te.toPlainText().strip() or "{}")
                except json.JSONDecodeError:
                    m = {}
                if not isinstance(m, dict):
                    m = {}
                m[key] = val
                te.setPlainText(json.dumps(m, ensure_ascii=False, indent=2))

            btn_add.clicked.connect(_add_rename)
            hr.addWidget(QLabel("旧列"))
            hr.addWidget(old_cb, 1)
            hr.addWidget(QLabel("新列"))
            hr.addWidget(new_le, 1)
            hr.addWidget(btn_add)
            self._fields["rename_old_cb"] = old_cb
            self._fields["rename_new_le"] = new_le
            self._param_layout.addRow("从主表列构造", row_r)
        elif op == "filter_rows":
            hint = QLabel(
                "只保留「条件成立」的行。列名含中文或空格时请用英文反引号 ` 包起来；"
                "详见右侧「?」或工具栏说明书。"
            )
            hint.setWordWrap(True)
            hint.setStyleSheet("color:#5F6B7A;font-size:12px;")
            self._param_layout.addRow(hint)
            q_edit = QLineEdit()
            q_edit.setPlaceholderText(
                "示例：`销售金额` > 1000   或   (地区 == \"华东\") & (数量 >= 10)"
            )
            q_edit.setToolTip(docs.tooltip("filter_rows"))
            attach_column_completer(q_edit, names)
            self._fields["query"] = q_edit
            self._param_layout.addRow("筛选条件（条件式）", q_edit)
            ins_row = QWidget()
            hi = QHBoxLayout(ins_row)
            hi.setContentsMargins(0, 0, 0, 0)
            hi.addWidget(QLabel("从主表插入列名"))
            ins = QComboBox()
            ins.addItem("— 点击一行插入 —", None)
            for n in names:
                ins.addItem(n, n)

            def _on_ins(idx: int) -> None:
                if idx <= 0:
                    return
                col = ins.itemData(idx)
                if col:
                    q_edit.insert(str(col))
                ins.blockSignals(True)
                ins.setCurrentIndex(0)
                ins.blockSignals(False)

            ins.currentIndexChanged.connect(_on_ins)
            hi.addWidget(ins, 1)
            self._param_layout.addRow("", ins_row)
        elif op == "sort":
            self._fields["column"] = make_column_combo(names)
            self._fields["ascending"] = QCheckBox("升序")
            self._fields["ascending"].setChecked(True)
            self._param_layout.addRow("排序列", self._fields["column"])
            self._param_layout.addRow("", self._fields["ascending"])
        elif op == "column_pick":
            lw, collect = make_column_list_picker(names, "", max_height=220)
            self._fields["list"] = lw
            self._list_collect = collect
            self._param_layout.addRow("勾选要保留的列", lw)
        elif op == "replace":
            self._fields["column"] = make_column_combo(names)
            self._fields["find"] = QLineEdit()
            self._fields["replace"] = QLineEdit()
            self._param_layout.addRow("列", self._fields["column"])
            self._param_layout.addRow("查找", self._fields["find"])
            self._param_layout.addRow("替换为", self._fields["replace"])
        elif op == "merge_vertical":
            cb = QComboBox()
            cb.setEditable(True)
            cb.setMinimumWidth(200)
            if self._session:
                for p in tabular_source_paths(self._session)[1:]:
                    cb.addItem(p.name, str(p.resolve()))
            b = QPushButton("其他文件…")
            b.clicked.connect(self._pick_second_merge)
            row = QWidget()
            hl = QHBoxLayout(row)
            hl.setContentsMargins(0, 0, 0, 0)
            hl.addWidget(cb, 1)
            hl.addWidget(b)
            self._fields["second_cb"] = cb
            self._fields["ignore_index"] = QCheckBox("忽略原索引")
            self._fields["ignore_index"].setChecked(True)
            self._param_layout.addRow(
                QLabel("第二个表格（可从已载入列表选，或浏览其他文件）"),
            )
            self._param_layout.addRow("", row)
            self._param_layout.addRow("", self._fields["ignore_index"])
        elif op == "template_fill":
            self._fields["template_path"] = QLineEdit()
            b = QPushButton("模板…")
            b.clicked.connect(self._pick_template)
            row = QWidget()
            hl = QHBoxLayout(row)
            hl.setContentsMargins(0, 0, 0, 0)
            hl.addWidget(self._fields["template_path"], 1)
            hl.addWidget(b)
            self._fields["name_column"] = make_column_combo(names)
            self._fields["sheet_name"] = QLineEdit("Sheet1")
            te = QTextEdit()
            te.setPlaceholderText('[["列名","A1"],…]')
            te.setMaximumHeight(100)
            self._fields["mapping_json"] = te
            self._param_layout.addRow("模板 xlsx", row)
            self._param_layout.addRow("文件命名列", self._fields["name_column"])
            self._param_layout.addRow("工作表名", self._fields["sheet_name"])
            self._param_layout.addRow("映射 JSON", te)
        elif op == "write_xlsx":
            self._fields["filename"] = QLineEdit()
            self._fields["sheet_name"] = QLineEdit("Sheet1")
            self._param_layout.addRow("文件名（可空=自动）", self._fields["filename"])
            self._param_layout.addRow("工作表名", self._fields["sheet_name"])
        elif op == "zip_pack":
            self._param_layout.addRow(
                QLabel("将把当前缓存区已生成的文件打成 ZIP，直接点「应用」。")
            )
        else:
            self._param_layout.addRow(QLabel(f"（未知操作: {op}）"))

    def _pick_second_merge(self) -> None:
        path, _ = QFileDialog.getOpenFileName(
            self, "第二个表格", "", "Excel (*.xlsx *.xls)"
        )
        if not path or "second_cb" not in self._fields:
            return
        cb = self._fields["second_cb"]
        ab = os.path.abspath(path)
        for i in range(cb.count()):
            if str(cb.itemData(i) or "") == ab:
                cb.setCurrentIndex(i)
                return
        cb.addItem(os.path.basename(ab), ab)
        cb.setCurrentIndex(cb.count() - 1)

    def _pick_template(self) -> None:
        path, _ = QFileDialog.getOpenFileName(
            self, "模板 xlsx", "", "Excel (*.xlsx)"
        )
        if path and "template_path" in self._fields:
            self._fields["template_path"].setText(os.path.abspath(path))

    def _prefill_fields(self, op: str, params: Dict[str, Any]) -> None:
        if op != "template_fill":
            return
        if "template_path" in self._fields:
            self._fields["template_path"].setText(
                str(params.get("template_path", ""))
            )
        if "name_column" in self._fields:
            cb = self._fields["name_column"]
            nc = str(params.get("name_column", ""))
            idx = cb.findText(nc)
            if idx >= 0:
                cb.setCurrentIndex(idx)
            else:
                cb.setEditText(nc)
        if "sheet_name" in self._fields:
            self._fields["sheet_name"].setText(
                str(params.get("sheet_name", "Sheet1"))
            )
        if "mapping_json" in self._fields:
            mj = params.get("mapping_json", "[]")
            if not isinstance(mj, str):
                mj = json.dumps(mj, ensure_ascii=False)
            self._fields["mapping_json"].setPlainText(mj)

    def _collect_params(self, op: str) -> Dict[str, Any]:
        f = self._fields
        if op == "format_std":
            cols = (
                self._fmt_columns_collect()
                if self._fmt_columns_collect
                else ""
            )
            return {
                "columns": cols,
                "strip": f["strip"].isChecked(),
            }
        if op == "dedup":
            sub = (
                self._dedup_subset_collect()
                if self._dedup_subset_collect
                else ""
            )
            return {
                "subset": sub,
                "keep": f["keep"].isChecked(),
            }
        if op == "fill_empty":
            cols = (
                self._fill_columns_collect()
                if self._fill_columns_collect
                else ""
            )
            return {
                "fill_value": f["fill_value"].text(),
                "columns": cols,
            }
        if op == "rename_cols":
            return {"mapping_json": f["mapping_json"].toPlainText().strip()}
        if op == "filter_rows":
            return {"query": f["query"].text().strip()}
        if op == "sort":
            return {
                "column": combo_column_text(f["column"]),
                "ascending": f["ascending"].isChecked(),
            }
        if op == "column_pick":
            raw = self._list_collect() if self._list_collect else ""
            return {"columns": raw}
        if op == "replace":
            return {
                "column": combo_column_text(f["column"]),
                "find": f["find"].text(),
                "replace": f["replace"].text(),
            }
        if op == "merge_vertical":
            cb = f["second_cb"]
            i = cb.currentIndex()
            data = cb.itemData(i) if i >= 0 else None
            path = (str(data) if data is not None else "").strip()
            if not path:
                path = cb.currentText().strip()
            return {
                "second_file": path,
                "ignore_index": f["ignore_index"].isChecked(),
            }
        if op == "template_fill":
            return {
                "template_path": f["template_path"].text().strip(),
                "name_column": combo_column_text(f["name_column"]),
                "sheet_name": f["sheet_name"].text().strip(),
                "mapping_json": f["mapping_json"].toPlainText().strip() or "[]",
            }
        if op == "write_xlsx":
            return {
                "filename": f["filename"].text().strip(),
                "sheet_name": f["sheet_name"].text().strip() or "Sheet1",
            }
        if op == "zip_pack":
            return {}
        return {}

    def _apply(self) -> None:
        if not self._session or not self._current_op:
            QMessageBox.information(self, "提示", "请先在左侧选择一项操作。")
            return
        params = self._collect_params(self._current_op)
        try:
            summary = apply_linear_op(
                self._session, self._current_op, params, self._ctx()
            )
        except Exception as e:
            QMessageBox.critical(self, "执行失败", str(e))
            return
        self._session.ops_history.append(
            AppliedOp(
                op_type=self._current_op,
                params=params,
                result_summary=summary,
            )
        )
        self._append_log(f"[线性] {self._current_label}: {summary}")
        self._preview.set_dataframe(self._session.dataframe)
        self._file_card.set_files(list(self._session.source_files))
        self._refresh_timeline()
        self.ops_changed.emit()

    def _delete_at(self, idx: int) -> None:
        if not self._session:
            return
        hist = self._session.ops_history
        if idx < 0 or idx >= len(hist):
            return
        remaining = hist[:idx] + hist[idx + 1 :]
        try:
            replay_ops(self._session, remaining, self._ctx())
        except Exception as e:
            QMessageBox.critical(self, "重放失败", str(e))
            return
        self._append_log(f"[线性] 已删除第 {idx + 1} 步并重放")
        self._preview.set_dataframe(self._session.dataframe)
        self._refresh_timeline()
        self.ops_changed.emit()

    def _refresh_timeline(self) -> None:
        if not self._session:
            self._timeline.set_ops([], None)
            return
        pending = None
        if self._current_op:
            pending = f"草稿：{self._current_label}"
        self._timeline.set_ops(list(self._session.ops_history), pending)
