"""线性流程 — 导出阶段：选目录、格式、列表项并写出或复制。"""

from __future__ import annotations

import os
import shutil
import zipfile
from io import StringIO
from typing import Callable, List, Optional

from PySide6.QtCore import Signal, Qt
from PySide6.QtWidgets import (
    QComboBox,
    QFileDialog,
    QFrame,
    QHBoxLayout,
    QLabel,
    QLineEdit,
    QListWidget,
    QListWidgetItem,
    QMessageBox,
    QPushButton,
    QVBoxLayout,
    QWidget,
)

from excel_workflow.session.state import LinearSession
from excel_workflow.ui import linear_feature_docs as docs
from excel_workflow.ui.linear_manual_dialog import show_export_help

try:
    import pandas as pd
except ImportError:
    pd = None  # type: ignore


class ExportPage(QWidget):
    """导出页：勾选项、目录、格式、单项或批量导出。"""

    export_done = Signal(str)
    back_requested = Signal()

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

        root = QVBoxLayout(self)
        root.setContentsMargins(8, 8, 8, 8)
        root.setSpacing(10)

        card = QFrame()
        card.setObjectName("LinearCard")
        inner = QVBoxLayout(card)
        inner.setContentsMargins(16, 16, 16, 16)
        inner.setSpacing(10)

        top = QHBoxLayout()
        top.addWidget(QLabel("导出列表（勾选要导出的项）"))
        top.addStretch(1)
        b_help = QPushButton("本页说明…")
        b_help.clicked.connect(lambda: show_export_help(self.window()))
        top.addWidget(b_help)
        inner.addLayout(top)

        self._list = QListWidget()
        self._list.setMinimumHeight(160)
        inner.addWidget(self._list, 1)

        row_dir = QHBoxLayout()
        row_dir.addWidget(QLabel("输出目录"))
        self._dir_edit = QLineEdit()
        self._dir_edit.setPlaceholderText("选择文件夹…")
        self._dir_edit.setToolTip(docs.tooltip("export_dir"))
        b_dir = QPushButton("浏览…")
        b_dir.clicked.connect(self._pick_dir)
        row_dir.addWidget(self._dir_edit, 1)
        row_dir.addWidget(b_dir)
        inner.addLayout(row_dir)

        row_fmt = QHBoxLayout()
        row_fmt.addWidget(QLabel("主表写出格式"))
        self._fmt = QComboBox()
        self._fmt.addItems(["xlsx", "csv", "zip", "pdf"])
        self._fmt.setToolTip(docs.tooltip("export_format"))
        # pdf 项单独提示（与历史 spec 中「pdf 待接入」一致）
        self._fmt.setItemData(3, docs.tooltip("export_pdf_stub"), Qt.ItemDataRole.ToolTipRole)
        row_fmt.addWidget(self._fmt)
        row_fmt.addStretch(1)
        inner.addLayout(row_fmt)

        btn_row = QHBoxLayout()
        self._btn_one = QPushButton("导出选中项")
        self._btn_one.setToolTip(docs.tooltip("export_selected_one"))
        self._btn_one.clicked.connect(self._export_selected)
        self._btn_all = QPushButton("导出全部（已勾选）")
        self._btn_all.setToolTip(docs.tooltip("export_batch_all"))
        self._btn_all.clicked.connect(self._export_all_checked)
        self._btn_back = QPushButton("返回操作")
        self._btn_back.setToolTip(docs.tooltip("export_back"))
        self._btn_back.clicked.connect(self.back_requested.emit)
        btn_row.addWidget(self._btn_one)
        btn_row.addWidget(self._btn_all)
        btn_row.addStretch(1)
        if on_back_to_import:
            b_imp = QPushButton("返回导入")
            b_imp.setToolTip(docs.tooltip("back_to_import_step"))
            b_imp.clicked.connect(on_back_to_import)
            btn_row.addWidget(b_imp)
        btn_row.addWidget(self._btn_back)
        inner.addLayout(btn_row)

        root.addWidget(card, 1)

    def prepare(self, session: LinearSession, select_all: bool = False) -> None:
        self._session = session
        self._list.clear()
        items: List[str] = [str(p) for p in session.source_files]
        items.extend(str(a) for a in session.artifacts)
        for i, p in enumerate(items):
            it = QListWidgetItem(f"{os.path.basename(p)}  ·  {p}")
            it.setFlags(it.flags() | Qt.ItemFlag.ItemIsUserCheckable)
            it.setData(Qt.UserRole, p)
            if select_all:
                it.setCheckState(Qt.CheckState.Checked)
            else:
                it.setCheckState(
                    Qt.CheckState.Checked if i == 0 else Qt.CheckState.Unchecked
                )
            self._list.addItem(it)

    def _pick_dir(self) -> None:
        d = QFileDialog.getExistingDirectory(self, "选择输出目录", self._dir_edit.text())
        if d:
            self._dir_edit.setText(os.path.abspath(d))

    def _primary_path(self) -> Optional[str]:
        if not self._session or not self._session.source_files:
            return None
        return str(self._session.source_files[0].resolve())

    def _checked_paths(self) -> List[str]:
        out: List[str] = []
        for i in range(self._list.count()):
            it = self._list.item(i)
            if it.checkState() == Qt.CheckState.Checked:
                s = it.data(Qt.UserRole)
                if s:
                    out.append(str(s))
        return out

    def _export_selected(self) -> None:
        paths = []
        for it in self._list.selectedItems():
            s = it.data(Qt.UserRole)
            if s:
                paths.append(str(s))
        if not paths:
            QMessageBox.information(self, "提示", "请在列表中单击选中一项后再导出。")
            return
        self._do_export(paths)

    def _export_all_checked(self) -> None:
        paths = self._checked_paths()
        if not paths:
            QMessageBox.information(self, "提示", "请至少勾选一项。")
            return
        self._do_export(paths)

    def _do_export(self, paths: List[str]) -> None:
        if not self._session:
            return
        out_dir = self._dir_edit.text().strip()
        if not out_dir:
            QMessageBox.warning(self, "提示", "请先选择输出目录。")
            return
        out_dir = os.path.abspath(out_dir)
        os.makedirs(out_dir, exist_ok=True)
        fmt = self._fmt.currentText().strip().lower()
        primary = self._primary_path()
        df = self._session.dataframe

        if fmt == "pdf":
            QMessageBox.information(
                self,
                "PDF 导出",
                docs.tooltip("export_pdf_stub"),
            )
            return

        try:
            for p in paths:
                p = os.path.abspath(p)
                base = os.path.basename(p)
                stem, _ext = os.path.splitext(base)
                is_primary = bool(primary) and os.path.normcase(
                    p
                ) == os.path.normcase(os.path.abspath(primary))

                if is_primary and df is not None and pd is not None:
                    dest = self._dest_path_for_table(out_dir, stem, fmt)
                    self._write_dataframe(df, dest, fmt)
                    self._append_log(f"[导出] {dest}")
                else:
                    if not os.path.isfile(p):
                        raise FileNotFoundError(f"不是有效文件: {p}")
                    dest = os.path.join(out_dir, base)
                    dest = self._unique_path(dest)
                    shutil.copy2(p, dest)
                    self._append_log(f"[导出] 复制 {dest}")
        except Exception as e:
            QMessageBox.critical(self, "导出失败", str(e))
            return

        QMessageBox.information(self, "完成", f"已导出到:\n{out_dir}")
        self.export_done.emit(out_dir)

    @staticmethod
    def _unique_path(path: str) -> str:
        if not os.path.exists(path):
            return path
        d, base = os.path.split(path)
        stem, ext = os.path.splitext(base)
        n = 1
        while True:
            cand = os.path.join(d, f"{stem}_{n}{ext}")
            if not os.path.exists(cand):
                return cand
            n += 1

    def _dest_path_for_table(self, out_dir: str, stem: str, fmt: str) -> str:
        if fmt == "csv":
            name = f"{stem}.csv"
        elif fmt == "zip":
            name = f"{stem}.zip"
        else:
            name = f"{stem}.xlsx"
        return self._unique_path(os.path.join(out_dir, name))

    def _write_dataframe(self, df, dest: str, fmt: str) -> None:
        if pd is None:
            raise RuntimeError("未安装 pandas")
        if fmt == "csv":
            df.to_csv(dest, index=False, encoding="utf-8-sig")
        elif fmt == "zip":
            csv_name = os.path.splitext(os.path.basename(dest))[0] + ".csv"
            with zipfile.ZipFile(dest, "w", zipfile.ZIP_DEFLATED) as zf:
                buf = StringIO()
                df.to_csv(buf, index=False, encoding="utf-8-sig")
                zf.writestr(csv_name, buf.getvalue().encode("utf-8-sig"))
        else:
            df.to_excel(dest, index=False)
