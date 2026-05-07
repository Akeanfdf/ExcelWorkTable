"""列名选择：下拉 + 可编辑联想，多列勾选列表。"""

from __future__ import annotations

from typing import Callable, List, Tuple

from PySide6.QtCore import Qt
from PySide6.QtWidgets import (
    QAbstractItemView,
    QComboBox,
    QCompleter,
    QHBoxLayout,
    QLabel,
    QLineEdit,
    QListWidget,
    QListWidgetItem,
    QSizePolicy,
    QVBoxLayout,
    QWidget,
)


def column_options_from_dataframe(df) -> List[str]:
    if df is None:
        return []
    return sorted(df.columns.astype(str).tolist(), key=str.lower)


def make_column_combo(names: List[str], current: str = "") -> QComboBox:
    cb = QComboBox()
    cb.setEditable(True)
    cb.setInsertPolicy(QComboBox.InsertPolicy.NoInsert)
    cb.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Fixed)
    for n in names:
        cb.addItem(n)
    if current and cb.findText(current) < 0:
        cb.addItem(current)
    idx = cb.findText(current)
    if idx >= 0:
        cb.setCurrentIndex(idx)
    elif current:
        cb.setEditText(current)
    else:
        cb.setEditText("")
    comp = QCompleter(names)
    comp.setCaseSensitivity(Qt.CaseSensitivity.CaseInsensitive)
    comp.setFilterMode(Qt.MatchFlag.MatchContains)
    cb.setCompleter(comp)
    return cb


def combo_column_text(cb: QComboBox) -> str:
    return cb.currentText().strip()


def attach_column_completer(line_edit: QLineEdit, names: List[str]) -> None:
    comp = QCompleter(names)
    comp.setCaseSensitivity(Qt.CaseSensitivity.CaseInsensitive)
    comp.setFilterMode(Qt.MatchFlag.MatchContains)
    line_edit.setCompleter(comp)


def make_column_list_picker(
    names: List[str], prefill_csv: str = "", max_height: int = 0
) -> Tuple[QListWidget, Callable[[], str]]:
    lw = QListWidget()
    lw.setSelectionMode(QAbstractItemView.SelectionMode.MultiSelection)
    lw.setMinimumHeight(min(120, 28 + min(len(names), 6) * 22))
    if max_height > 0:
        lw.setMaximumHeight(max_height)
    for n in names:
        lw.addItem(QListWidgetItem(n))
    want = {x.strip() for x in prefill_csv.split(",") if x.strip()}
    for i in range(lw.count()):
        it = lw.item(i)
        if it.text() in want:
            it.setSelected(True)

    def collect() -> str:
        return ",".join(it.text() for it in lw.selectedItems())

    return lw, collect


def column_hint_label(names: List[str]) -> QLabel:
    if not names:
        return QLabel("（当前无主表列名，请先完成导入）")
    preview = "、".join(names[:12])
    if len(names) > 12:
        preview += "…"
    lab = QLabel(f"主表列名参考：{preview}")
    lab.setWordWrap(True)
    lab.setStyleSheet("color:#5F6B7A;font-size:11px;")
    return lab
