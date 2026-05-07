from __future__ import annotations

from PySide6.QtCore import Qt
from PySide6.QtWidgets import QHBoxLayout, QLabel, QWidget

from excel_workflow.core import theme
from excel_workflow.ui import linear_feature_docs as docs


class FlowPills(QWidget):
    """顶部三步 pill：导入 › 操作 › 导出（仅展示，不可点击跳步）。"""

    def __init__(self, parent=None):
        super().__init__(parent)
        lay = QHBoxLayout(self)
        lay.setContentsMargins(0, 0, 0, 0)
        lay.setSpacing(8)
        self._labels: list[QLabel] = []
        titles = ["导入", "操作", "导出"]
        for i, t in enumerate(titles):
            lab = QLabel(f"  {i + 1}  {t}  ")
            lab.setAlignment(Qt.AlignCenter)
            lab.setMinimumWidth(96)
            self._labels.append(lab)
            lay.addWidget(lab)
            if i < len(titles) - 1:
                sep = QLabel("›")
                sep.setStyleSheet(f"color: {theme.MUTED};")
                lay.addWidget(sep)
        lay.addStretch(1)
        self.setToolTip(docs.tooltip("pills_phases"))
        self.set_step(0)

    def set_step(self, step: int) -> None:
        step = max(0, min(2, int(step)))
        active = f"""
            QLabel {{
                background-color: {theme.PRIMARY};
                color: #ffffff;
                border-radius: 16px;
                padding: 6px 12px;
                font-weight: 600;
            }}
        """
        done = f"""
            QLabel {{
                background-color: {theme.SURFACE};
                color: {theme.TEXT};
                border: 1px solid {theme.OUTLINE};
                border-radius: 16px;
                padding: 6px 12px;
            }}
        """
        pending = f"""
            QLabel {{
                background-color: {theme.SURFACE_VARIANT};
                color: {theme.MUTED};
                border-radius: 16px;
                padding: 6px 12px;
            }}
        """
        for i, lab in enumerate(self._labels):
            if i < step:
                lab.setStyleSheet(done)
            elif i == step:
                lab.setStyleSheet(active)
            else:
                lab.setStyleSheet(pending)
