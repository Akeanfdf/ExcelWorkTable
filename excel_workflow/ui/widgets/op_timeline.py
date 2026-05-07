from __future__ import annotations

from typing import Callable, List, Optional

from PySide6.QtCore import Qt
from PySide6.QtWidgets import (
    QFrame,
    QHBoxLayout,
    QLabel,
    QPushButton,
    QVBoxLayout,
    QWidget,
)

from excel_workflow.session.state import AppliedOp
from excel_workflow.ui import linear_feature_docs as docs


class OpTimeline(QWidget):
    """操作序列时间线：编号、名称、摘要、完成标记、删除。"""

    def __init__(
        self,
        on_delete_index: Callable[[int], None],
        parent=None,
    ):
        super().__init__(parent)
        self._vl = QVBoxLayout(self)
        self._vl.setContentsMargins(4, 4, 4, 4)
        self._vl.setSpacing(6)
        self._on_delete = on_delete_index
        self.setToolTip(docs.tooltip("timeline_delete"))

    def clear(self) -> None:
        while self._vl.count():
            item = self._vl.takeAt(0)
            w = item.widget()
            if w:
                w.deleteLater()

    def set_ops(self, ops: List[AppliedOp], pending_label: Optional[str] = None) -> None:
        self.clear()
        if not ops and not pending_label:
            hint = QLabel(
                "暂无步骤。请在左侧点击一项操作，在下方填写参数后点「应用」"
                "（完成后此处会显示 ✓ 与摘要）。"
            )
            hint.setWordWrap(True)
            hint.setStyleSheet("color: #5F6B7A; padding: 10px; background: #F3F4F6; border-radius: 8px;")
            self._vl.addWidget(hint)
        for i, op in enumerate(ops):
            row = QFrame()
            row.setFrameShape(QFrame.StyledPanel)
            hl = QHBoxLayout(row)
            mark = "✓" if op.result_summary else "···"
            color = "#43A047" if op.result_summary else "#3F51B5"
            title = self._title(op.op_type)
            lab = QLabel(f'<span style="color:{color};font-weight:600">{mark}</span>  {i + 1}. {title}')
            lab.setTextFormat(Qt.RichText)
            sumy = op.result_summary or "（未完成）"
            sub = QLabel(sumy)
            sub.setStyleSheet("color: #5F6B7A; font-size: 11px;")
            sub.setWordWrap(True)
            vb = QVBoxLayout()
            vb.addWidget(lab)
            vb.addWidget(sub)
            hl.addLayout(vb, 1)
            del_btn = QPushButton("删除")
            del_btn.setFixedWidth(52)
            del_btn.setToolTip(docs.tooltip("timeline_delete"))
            del_btn.clicked.connect(lambda _=False, idx=i: self._on_delete(idx))
            hl.addWidget(del_btn, 0, Qt.AlignTop)
            self._vl.addWidget(row)
        if pending_label:
            pend = QLabel(f"◆ {pending_label}")
            pend.setStyleSheet("color: #3F51B5; font-weight: 600;")
            self._vl.addWidget(pend)

    @staticmethod
    def _title(op_type: str) -> str:
        m = {
            "format_std": "格式标准化",
            "dedup": "去重",
            "fill_empty": "空值填充",
            "rename_cols": "列重命名",
            "filter_rows": "筛选行",
            "merge_vertical": "合并文件",
            "template_fill": "填充模板",
            "write_xlsx": "写出 xlsx",
            "zip_pack": "打包 ZIP",
            "sort": "排序",
            "column_pick": "选列 / 删列",
            "replace": "查找替换",
        }
        return m.get(op_type, op_type)
