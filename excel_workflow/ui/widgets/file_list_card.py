from __future__ import annotations

import os
from pathlib import Path
from typing import Callable, List, Optional

from PySide6.QtCore import Qt
from PySide6.QtWidgets import (
    QAbstractItemView,
    QGroupBox,
    QHBoxLayout,
    QLabel,
    QListWidget,
    QListWidgetItem,
    QMessageBox,
    QPushButton,
    QVBoxLayout,
)

from excel_workflow.ui import linear_feature_docs as docs

StatusFn = Callable[[Path], str]


class FileListCard(QGroupBox):
    """已载入文件列表：名称、大小、状态。"""

    def __init__(
        self,
        on_append: Optional[Callable[[], None]] = None,
        on_remove_selected: Optional[Callable[[List[Path]], None]] = None,
        parent=None,
    ):
        super().__init__("已载入文件", parent)
        self.setObjectName("LinearCard")
        self._on_append = on_append
        self._on_remove_selected = on_remove_selected
        root = QVBoxLayout(self)
        top = QHBoxLayout()
        top.addWidget(QLabel("支持追加导入；列表中可多选后删除"))
        top.addStretch(1)
        if on_remove_selected:
            br = QPushButton("删除所选")
            br.setToolTip(docs.tooltip("remove_source_files"))
            br.clicked.connect(self._emit_remove_selected)
            top.addWidget(br)
        if on_append:
            b = QPushButton("追加导入…")
            b.setToolTip(docs.tooltip("append_import"))
            b.clicked.connect(on_append)
            top.addWidget(b)
        root.addLayout(top)
        self._list = QListWidget()
        self._list.setMinimumHeight(100)
        self._list.setSelectionMode(QAbstractItemView.SelectionMode.ExtendedSelection)
        root.addWidget(self._list)
        self._status_fn: Optional[StatusFn] = None

    def set_status_fn(self, fn: Optional[StatusFn]) -> None:
        self._status_fn = fn

    def set_files(self, paths: List[Path]) -> None:
        self._list.clear()
        for p in paths:
            self._add_row(p)

    def _add_row(self, p: Path) -> None:
        try:
            sz = p.stat().st_size
            szs = f"{sz / 1024:.1f} KB" if sz < 1024 * 1024 else f"{sz / 1024 / 1024:.2f} MB"
        except OSError:
            szs = "?"
        st = "就绪"
        if self._status_fn:
            try:
                st = self._status_fn(p)
            except Exception:
                st = "就绪"
        it = QListWidgetItem(f"{p.name}  ·  {szs}  ·  {st}")
        it.setData(Qt.UserRole, str(p))
        self._list.addItem(it)

    def append_files(self, paths: List[Path]) -> None:
        for p in paths:
            self._add_row(p)

    def paths(self) -> List[Path]:
        out: List[Path] = []
        for i in range(self._list.count()):
            it = self._list.item(i)
            s = it.data(Qt.UserRole)
            if s:
                out.append(Path(str(s)))
        return out

    def selected_paths(self) -> List[Path]:
        out: List[Path] = []
        for it in self._list.selectedItems():
            s = it.data(Qt.UserRole)
            if s:
                out.append(Path(str(s)))
        return out

    def _emit_remove_selected(self) -> None:
        if not self._on_remove_selected:
            return
        paths = self.selected_paths()
        if not paths:
            QMessageBox.information(
                self, "提示", "请先在列表中选中要移除的文件（可按住 Ctrl 多选）。"
            )
            return
        self._on_remove_selected(paths)
