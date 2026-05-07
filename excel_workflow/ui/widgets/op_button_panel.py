from __future__ import annotations

from typing import Callable

from PySide6.QtCore import Qt
from PySide6.QtWidgets import (
    QFrame,
    QGroupBox,
    QHBoxLayout,
    QPushButton,
    QScrollArea,
    QSizePolicy,
    QVBoxLayout,
    QWidget,
)

from excel_workflow.ui import linear_feature_docs as docs
from excel_workflow.ui.linear_manual_dialog import help_tool_button

OpClicked = Callable[[str, str], None]  # op_key, label


class OpButtonPanel(QWidget):
    """
    左侧分组块状按钮（与主界面 LinearContentHost 内 QSS 统一为节点图时代的卡片风格）。
    """

    def __init__(self, on_op: OpClicked, parent=None):
        super().__init__(parent)
        self._on_op = on_op
        self.setMinimumWidth(220)
        self.setMaximumWidth(300)
        self.setSizePolicy(
            QSizePolicy.Policy.Fixed, QSizePolicy.Policy.Expanding
        )

        outer = QVBoxLayout(self)
        outer.setContentsMargins(0, 0, 0, 0)
        outer.setSpacing(0)

        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        scroll.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAlwaysOff)
        scroll.setFrameShape(QFrame.Shape.NoFrame)
        scroll.setSizePolicy(
            QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Expanding
        )

        inner = QWidget()
        root = QVBoxLayout(inner)
        root.setSpacing(10)
        inner.setMinimumWidth(200)

        root.addWidget(self._section("数据操作", [
            ("format_std", "格式标准化"),
            ("dedup", "去重"),
            ("fill_empty", "空值填充"),
            ("filter_rows", "筛选行"),
            ("sort", "排序"),
            ("column_pick", "选列 / 删列"),
            ("replace", "查找替换"),
            ("rename_cols", "列重命名"),
        ]))
        root.addWidget(self._section("结构操作", [
            ("merge_vertical", "合并文件"),
            ("template_fill", "填充模板"),
        ]))
        root.addWidget(self._section("格式与输出", [
            ("write_xlsx", "写出 xlsx"),
            ("zip_pack", "打包 ZIP"),
        ]))
        root.addStretch(1)

        scroll.setWidget(inner)
        outer.addWidget(scroll)

    def _op_row(self, key: str, label: str) -> QWidget:
        row = QWidget()
        hl = QHBoxLayout(row)
        hl.setContentsMargins(0, 2, 0, 2)
        hl.setSpacing(4)
        b = QPushButton(label)
        tip = docs.tooltip(key)
        if tip:
            b.setToolTip(tip + "\n\n点右侧「?」查看完整说明。")
        hb = help_tool_button(self, key)
        b.clicked.connect(lambda _=False, k=key, lb=label: self._on_op(k, lb))
        hl.addWidget(b, 1)
        hl.addWidget(hb, 0)
        return row

    def _section(self, title: str, items: list[tuple[str, str]]) -> QGroupBox:
        gb = QGroupBox(title)
        gb.setObjectName("LinearCard")
        gb.setCheckable(False)
        gl = QVBoxLayout(gb)
        for key, lab in items:
            gl.addWidget(self._op_row(key, lab))
        return gb
