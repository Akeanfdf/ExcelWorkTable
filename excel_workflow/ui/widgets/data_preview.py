from __future__ import annotations

from typing import Any

from PySide6.QtCore import Qt
from PySide6.QtWidgets import (
    QGroupBox,
    QSizePolicy,
    QTextEdit,
    QVBoxLayout,
)

from excel_workflow.ui import linear_feature_docs as docs

try:
    import pandas as pd
except ImportError:
    pd = None  # type: ignore


class DataPreviewPanel(QGroupBox):
    """可折叠式数据预览（前 100 行）。"""

    def __init__(self, parent=None):
        super().__init__("数据预览（点击标题展开/折叠）", parent)
        self.setCheckable(True)
        self.setChecked(False)
        self.setToolTip(docs.tooltip("preview_data"))
        self.toggled.connect(self._on_toggle)
        lay = QVBoxLayout(self)
        self._text = QTextEdit()
        self._text.setObjectName("DataPreviewEdit")
        self._text.setReadOnly(True)
        self._text.setMinimumHeight(120)
        self._text.setMaximumHeight(420)
        self._text.setLineWrapMode(QTextEdit.LineWrapMode.NoWrap)
        self._text.setVerticalScrollBarPolicy(
            Qt.ScrollBarPolicy.ScrollBarAsNeeded
        )
        self._text.setHorizontalScrollBarPolicy(
            Qt.ScrollBarPolicy.ScrollBarAsNeeded
        )
        # 纵向由外层「右侧整栏」滚动承担，避免预览区把整栏最小高度撑满视口导致不出滚动条
        self._text.setSizePolicy(
            QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Preferred
        )
        self._text.hide()
        lay.addWidget(self._text)
        self._df: Any = None

    def _on_toggle(self, on: bool) -> None:
        self._text.setVisible(on)
        if on:
            self._refresh()

    def set_dataframe(self, df: Any) -> None:
        self._df = df
        if self.isChecked():
            self._refresh()

    def _refresh(self) -> None:
        df = self._df
        if df is None or pd is None:
            self._text.setPlainText("（无表格）")
            return
        try:
            s = df.head(100).to_string()
            self._text.setPlainText(s)
        except Exception as e:
            self._text.setPlainText(str(e))
