from __future__ import annotations

import json
import os
from typing import List

from PySide6.QtCore import Qt, Signal
from PySide6.QtGui import QDragEnterEvent, QDropEvent
from PySide6.QtWidgets import (
    QAbstractItemView,
    QDialog,
    QDialogButtonBox,
    QFileDialog,
    QFrame,
    QHBoxLayout,
    QLabel,
    QListWidget,
    QListWidgetItem,
    QMessageBox,
    QPushButton,
    QVBoxLayout,
    QWidget,
)

from excel_workflow.session.workflow_io import mapping_config_to_template_ops
from excel_workflow.ui import linear_feature_docs as docs
from excel_workflow.ui.linear_manual_dialog import show_import_help
from excel_workflow.utils.file_scanner import scan_spreadsheets


class _DropZone(QWidget):
    """虚线拖放区。"""

    files_dropped = Signal(list)

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setAcceptDrops(True)
        self.setMinimumHeight(180)
        self.setObjectName("ImportDropZone")
        self.setToolTip(docs.tooltip("import_drop"))
        self.setStyleSheet(
            """
            #ImportDropZone {
                border: 2px dashed #9CA3AF;
                border-radius: 12px;
                background: #FAFBFC;
            }
            """
        )
        lay = QVBoxLayout(self)
        lab = QLabel("将 xlsx / csv 拖入此区域\n或点击下方按钮选择")
        lab.setAlignment(Qt.AlignCenter)
        lay.addWidget(lab)

    def dragEnterEvent(self, event: QDragEnterEvent) -> None:
        if event.mimeData().hasUrls():
            event.acceptProposedAction()
        else:
            event.ignore()

    def dropEvent(self, event: QDropEvent) -> None:
        paths = []
        for u in event.mimeData().urls():
            p = u.toLocalFile()
            if p and os.path.isfile(p):
                ext = os.path.splitext(p)[1].lower()
                if ext in (".xlsx", ".xls", ".csv"):
                    paths.append(os.path.abspath(p))
        if paths:
            self.files_dropped.emit(paths)
        event.acceptProposedAction()


class ImportPage(QWidget):
    """导入阶段：全屏居中导入区。"""

    import_confirmed = Signal(list)  # list[str] paths
    mapping_legacy_loaded = Signal(object)  # dict template ops payload for operate page

    def __init__(self, parent=None):
        super().__init__(parent)
        root = QVBoxLayout(self)
        root.addStretch(2)
        box = QFrame()
        box.setObjectName("LinearCard")
        bl = QVBoxLayout(box)
        bl.setContentsMargins(20, 20, 20, 20)
        bl.setSpacing(16)
        self._drop = _DropZone()
        self._drop.files_dropped.connect(self._emit_paths)
        bl.addWidget(self._drop)
        row = QHBoxLayout()
        b1 = QPushButton("选择文件（xlsx / csv）…")
        b1.setToolTip(docs.tooltip("import_select_files"))
        b1.clicked.connect(self._pick_files)
        b2 = QPushButton("选择文件夹（扫描后勾选）…")
        b2.setToolTip(docs.tooltip("import_folder"))
        b2.clicked.connect(self._pick_folder_scan)
        row.addStretch(1)
        row.addWidget(b1)
        row.addWidget(b2)
        row.addStretch(1)
        bl.addLayout(row)
        b3 = QPushButton("导入旧版 mapping_config.json…")
        b3.setToolTip(docs.tooltip("import_mapping"))
        b3.clicked.connect(self._load_mapping_json)
        bl.addWidget(b3, alignment=Qt.AlignCenter)
        hint_row = QHBoxLayout()
        hint_row.addStretch(1)
        b_help = QPushButton("本页功能说明…")
        b_help.setToolTip("查看导入阶段各操作的详细说明")
        b_help.clicked.connect(lambda: show_import_help(self.window()))
        hint_row.addWidget(b_help)
        hint_row.addStretch(1)
        bl.addLayout(hint_row)
        root.addWidget(box, alignment=Qt.AlignCenter)
        root.addStretch(3)

    def _emit_paths(self, paths: List[str]) -> None:
        if paths:
            self.import_confirmed.emit(paths)

    def _pick_files(self) -> None:
        paths, _ = QFileDialog.getOpenFileNames(
            self,
            "选择表格文件",
            "",
            "表格 (*.xlsx *.xls *.csv)",
        )
        if paths:
            self.import_confirmed.emit([os.path.abspath(p) for p in paths])

    def _pick_folder_scan(self) -> None:
        d = QFileDialog.getExistingDirectory(self, "选择要扫描的文件夹", "")
        if not d:
            return
        found = scan_spreadsheets(d)
        if not found:
            QMessageBox.information(self, "提示", "该文件夹下未发现 xlsx/xls/csv")
            return
        dlg = QDialog(self)
        dlg.setWindowTitle("勾选要载入的文件")
        vl = QVBoxLayout(dlg)
        vl.addWidget(QLabel(f"共 {len(found)} 个文件，请勾选："))
        lw = QListWidget()
        lw.setSelectionMode(QAbstractItemView.NoSelection)
        for p in found:
            it = QListWidgetItem(os.path.basename(p))
            it.setData(Qt.UserRole, p)
            it.setCheckState(Qt.Checked)
            lw.addItem(it)
        vl.addWidget(lw, 1)
        bb = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        bb.accepted.connect(dlg.accept)
        bb.rejected.connect(dlg.reject)
        vl.addWidget(bb)
        if dlg.exec() != QDialog.DialogCode.Accepted:
            return
        chosen = []
        for i in range(lw.count()):
            it = lw.item(i)
            if it.checkState() == Qt.Checked:
                chosen.append(str(it.data(Qt.UserRole)))
        if not chosen:
            QMessageBox.warning(self, "提示", "请至少勾选一个文件")
            return
        self.import_confirmed.emit(chosen)

    def _load_mapping_json(self) -> None:
        path, _ = QFileDialog.getOpenFileName(
            self, "选择 mapping_config.json", "", "JSON (*.json)"
        )
        if not path:
            return
        try:
            with open(path, encoding="utf-8") as f:
                cfg = json.load(f)
        except Exception as e:
            QMessageBox.critical(self, "错误", str(e))
            return
        try:
            ops = mapping_config_to_template_ops(cfg)
        except Exception as e:
            QMessageBox.critical(self, "错误", f"转换失败: {e}")
            return
        self.mapping_legacy_loaded.emit({"legacy_mapping_config": cfg, "prefill_ops": ops})
        QMessageBox.information(
            self,
            "已导入",
            "旧版映射已读取。请先完成表格导入进入「操作」阶段后，"
            "可使用「填充模板」并核对自动填写的参数。",
        )
