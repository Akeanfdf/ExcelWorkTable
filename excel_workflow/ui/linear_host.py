from __future__ import annotations

from pathlib import Path
from typing import Any, Callable, Dict, Optional

from PySide6.QtCore import Qt, Slot
from PySide6.QtWidgets import (
    QFileDialog,
    QFrame,
    QHBoxLayout,
    QLabel,
    QMessageBox,
    QProgressBar,
    QPushButton,
    QStackedWidget,
    QVBoxLayout,
    QWidget,
)

from excel_workflow.core.context import ExecContext
from excel_workflow.ops.runner_bridge import load_primary_table
from excel_workflow.session.state import LinearSession
from excel_workflow.session.workflow_io import save_workflow_v2
from excel_workflow.ui import linear_feature_docs as docs
from excel_workflow.ui.pages.export_page import ExportPage
from excel_workflow.ui.pages.import_page import ImportPage
from excel_workflow.ui.pages.operate_page import OperatePage
from excel_workflow.ui.widgets.flow_pills import FlowPills


class LinearFlowHost(QWidget):
    """线性主流程：导入 › 操作 › 导出（视觉与 Material 节点壳统一）。"""

    def __init__(self, append_log: Callable[[str], None], parent=None):
        super().__init__(parent)
        self._append_log = append_log
        self._session = LinearSession()
        self._pending_mapping: Optional[Dict[str, Any]] = None

        root = QVBoxLayout(self)
        root.setContentsMargins(10, 10, 10, 10)
        root.setSpacing(10)

        top_frame = QFrame()
        top_frame.setObjectName("LinearHostTop")
        top = QHBoxLayout(top_frame)
        top.setContentsMargins(12, 10, 12, 10)
        logo = QLabel("Excel 工作流")
        logo.setStyleSheet("font-weight: 700; font-size: 15px;")
        self._pills = FlowPills()
        top.addWidget(logo)
        top.addWidget(self._pills, 1)
        root.addWidget(top_frame)

        self._stack = QStackedWidget()
        self._stack.setObjectName("LinearStack")
        self._stack.setMinimumHeight(0)
        self._import_page = ImportPage()
        self._import_page.import_confirmed.connect(self._on_import_confirmed)
        self._import_page.mapping_legacy_loaded.connect(self._on_mapping_legacy)
        self._operate_page = OperatePage(append_log, on_back_to_import=self._back_to_import)
        self._operate_page.ops_changed.connect(self._refresh_bottom)
        self._export_page = ExportPage(append_log, on_back_to_import=self._back_to_import)
        self._export_page.export_done.connect(self._on_export_done)
        self._export_page.back_requested.connect(self._back_to_operate)
        self._stack.addWidget(self._import_page)
        self._stack.addWidget(self._operate_page)
        self._stack.addWidget(self._export_page)
        root.addWidget(self._stack, 1)

        bot_frame = QFrame()
        bot_frame.setObjectName("LinearHostBottom")
        bot_frame.setMinimumHeight(52)
        bot = QHBoxLayout(bot_frame)
        bot.setContentsMargins(12, 10, 12, 10)
        self._step_lab = QLabel("已执行 0 步")
        self._mini_prog = QProgressBar()
        self._mini_prog.setRange(0, 100)
        self._mini_prog.setValue(0)
        self._mini_prog.setFixedHeight(4)
        self._mini_prog.setTextVisible(False)
        self._mini_prog.setMaximumWidth(220)
        self._btn_save_flow = QPushButton("保存流程")
        self._btn_save_flow.setToolTip(docs.tooltip("bottom_save"))
        self._btn_save_flow.clicked.connect(self._save_workflow)
        self._btn_export = QPushButton("导出文件")
        self._btn_export.setToolTip(docs.tooltip("bottom_export"))
        self._btn_export.setStyleSheet("font-weight: 600;")
        self._btn_export.clicked.connect(self._on_export_clicked)
        self._btn_batch_export = QPushButton("批量导出")
        self._btn_batch_export.setToolTip(docs.tooltip("bottom_batch_export"))
        self._btn_batch_export.clicked.connect(self._on_batch_export_clicked)
        bot.addWidget(self._step_lab)
        bot.addWidget(self._mini_prog)
        bot.addStretch(1)
        bot.addWidget(self._btn_save_flow)
        bot.addWidget(self._btn_export)
        bot.addWidget(self._btn_batch_export)
        root.addWidget(bot_frame)

        self._stack.setCurrentIndex(0)
        self._pills.set_step(0)

    def _back_to_import(self) -> None:
        """从操作/导出阶段回到导入，清空会话。"""
        if self._session.ops_history or self._session.source_files:
            r = QMessageBox.question(
                self,
                "返回导入",
                "将清空已载入文件与已执行步骤，回到「导入」界面。\n\n确定继续？",
                QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
                QMessageBox.StandardButton.No,
            )
            if r != QMessageBox.StandardButton.Yes:
                return
        self._session.reset_for_new_import()
        self._pending_mapping = None
        self._pills.set_step(0)
        self._stack.setCurrentIndex(0)
        self._operate_page.set_session(self._session)
        self._refresh_bottom()
        self._append_log("[线性] 已返回导入界面")

    def _ctx(self) -> ExecContext:
        def log(m: str) -> None:
            self._append_log(f"[线性] {m}")

        return ExecContext(log=log, progress=lambda _p, _m: None)

    def _refresh_bottom(self) -> None:
        n = len(self._session.ops_history)
        self._step_lab.setText(f"已执行 {n} 步")
        self._mini_prog.setValue(min(100, n * 15))

    def _on_mapping_legacy(self, payload: Dict[str, Any]) -> None:
        self._pending_mapping = payload
        if self._session.source_files and self._stack.currentIndex() == 1:
            self._operate_page.apply_mapping_prefill(payload)
            self._pending_mapping = None

    def _on_import_confirmed(self, paths: list) -> None:
        if not paths:
            return
        self._session.reset_for_new_import()
        for p in paths:
            self._session.source_files.append(Path(p))
        self._session.new_staging()
        try:
            msg = load_primary_table(self._session, self._ctx())
            self._append_log(msg)
        except Exception as e:
            QMessageBox.critical(self, "载入失败", str(e))
            return
        self._session.phase = 1
        self._session.current_files = list(self._session.source_files)
        self._session.is_dirty = True
        self._pills.set_step(1)
        self._stack.setCurrentIndex(1)
        self._operate_page.set_session(self._session)
        self._export_page.prepare(self._session, select_all=False)
        self._refresh_bottom()
        if self._pending_mapping:
            self._operate_page.apply_mapping_prefill(self._pending_mapping)
            self._pending_mapping = None

    @Slot()
    def _on_export_clicked(self) -> None:
        self._goto_export(select_all=False)

    @Slot()
    def _on_batch_export_clicked(self) -> None:
        self._goto_export(select_all=True)

    def _goto_export(self, select_all: bool = False) -> None:
        if not self._session.source_files:
            QMessageBox.warning(self, "提示", "请先完成导入。")
            return
        self._session.phase = 2
        self._pills.set_step(2)
        self._export_page.prepare(self._session, select_all=select_all)
        self._stack.setCurrentIndex(2)
        self._export_page.setFocus()

    def _back_to_operate(self) -> None:
        self._session.phase = 1
        self._pills.set_step(1)
        self._stack.setCurrentIndex(1)
        self._operate_page.set_session(self._session)

    def _save_workflow(self) -> None:
        if not self._session.ops_history:
            QMessageBox.information(self, "提示", "还没有可保存的操作步骤。")
            return
        path, _ = QFileDialog.getSaveFileName(
            self, "保存 workflow.json", "", "JSON (*.json)"
        )
        if not path:
            return
        name = Path(path).stem
        try:
            save_workflow_v2(
                path,
                name,
                self._session.to_serializable_ops(),
                self._session.legacy_mapping_config,
            )
            self._append_log(f"[线性] 已保存流程: {path}")
            QMessageBox.information(self, "完成", f"已保存:\n{path}")
        except Exception as e:
            QMessageBox.critical(self, "错误", str(e))

    def _on_export_done(self, _dest: str) -> None:
        self._refresh_bottom()
        n = len(self._session.ops_history)
        summary_lines = [
            f"{i + 1}. {o.op_type} — {o.result_summary}"
            for i, o in enumerate(self._session.ops_history)
        ]
        body = "\n".join(summary_lines) if summary_lines else "（无步骤）"
        box = QMessageBox(self)
        box.setWindowTitle("本次流程可保存为模板")
        box.setText(f"共 {n} 步操作。\n\n{body[:800]}")
        btn_save = box.addButton("保存为模板…", QMessageBox.ActionRole)
        box.addButton("下次再说", QMessageBox.RejectRole)
        box.exec()
        if box.clickedButton() == btn_save:
            path, _ = QFileDialog.getSaveFileName(
                self, "保存 workflow.json", "", "JSON (*.json)"
            )
            if path:
                try:
                    save_workflow_v2(
                        path,
                        Path(path).stem,
                        self._session.to_serializable_ops(),
                        self._session.legacy_mapping_config,
                    )
                    self._append_log(f"[线性] 模板已保存: {path}")
                    QMessageBox.information(self, "完成", f"已保存:\n{path}")
                except Exception as e:
                    QMessageBox.critical(self, "错误", str(e))
