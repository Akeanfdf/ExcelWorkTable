from __future__ import annotations

from typing import Optional

from PySide6.QtCore import Qt
from PySide6.QtWidgets import (
    QDialog,
    QDialogButtonBox,
    QTextBrowser,
    QToolButton,
    QVBoxLayout,
    QWidget,
)

from excel_workflow.ui import linear_feature_docs as docs


def _show_html_dialog(parent: Optional[QWidget], title: str, html: str) -> None:
    d = QDialog(parent)
    d.setWindowTitle(title)
    d.resize(640, 480)
    lay = QVBoxLayout(d)
    br = QTextBrowser()
    br.setHtml(html)
    lay.addWidget(br)
    bb = QDialogButtonBox(QDialogButtonBox.StandardButton.Close)
    bb.rejected.connect(d.reject)
    bb.accepted.connect(d.accept)
    lay.addWidget(bb)
    d.exec()


def show_full_manual(parent: Optional[QWidget]) -> None:
    _show_html_dialog(parent, "功能说明书 — 全文", docs.all_manual_html())


def show_import_help(parent: Optional[QWidget]) -> None:
    _show_html_dialog(parent, "导入阶段 — 说明", docs.import_section_html())


def show_export_help(parent: Optional[QWidget]) -> None:
    _show_html_dialog(parent, "导出阶段 — 说明", docs.export_section_html())


def show_op_manual(parent: Optional[QWidget], op_key: str) -> None:
    """单个操作的说明（「?」按钮）。"""
    _show_html_dialog(parent, f"说明 — {op_key}", docs.body_html(op_key))


def help_tool_button(parent: QWidget, op_key: str) -> QToolButton:
    """带「?」的小按钮，点击打开该功能说明。"""
    hb = QToolButton(parent)
    hb.setText("?")
    hb.setFixedSize(28, 28)
    hb.setToolTip("查看此项说明")
    hb.setAutoRaise(True)
    hb.clicked.connect(lambda: show_op_manual(parent.window() if parent else None, op_key))
    return hb
