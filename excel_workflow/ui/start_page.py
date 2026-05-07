"""Welcome / template cards overlay (Material cards)."""

from __future__ import annotations

from typing import Callable, List, Optional

from PySide6.QtCore import Qt
from PySide6.QtWidgets import (
    QFrame,
    QGridLayout,
    QHBoxLayout,
    QLabel,
    QPushButton,
    QScrollArea,
    QVBoxLayout,
    QWidget,
)

from excel_workflow.core import theme as t
from excel_workflow.ui import recent_workflows


class _Card(QFrame):
    def __init__(
        self,
        accent_rgb: tuple,
        title: str,
        desc: str,
        button_text: str,
        on_action: Callable[[], None],
        parent=None,
    ):
        super().__init__(parent)
        self.setObjectName("StartCard")
        r = t.RADIUS_LG
        self.setStyleSheet(
            f"#StartCard {{ background: {t.SURFACE}; border: 1px solid {t.OUTLINE}; "
            f"border-radius: {r}px; }}"
        )
        lay = QHBoxLayout(self)
        lay.setContentsMargins(0, 0, 16, 0)
        lay.setSpacing(0)

        stripe = QFrame()
        stripe.setFixedWidth(6)
        stripe.setStyleSheet(
            f"background: rgb({accent_rgb[0]},{accent_rgb[1]},{accent_rgb[2]}); "
            f"border-top-left-radius: {r}px; border-bottom-left-radius: {r}px;"
        )
        lay.addWidget(stripe)

        inner = QVBoxLayout()
        inner.setContentsMargins(16, 16, 8, 16)
        inner.setSpacing(8)
        t1 = QLabel(title)
        t1.setStyleSheet(f"font-size: 16px; font-weight: 700; color: {t.TEXT};")
        d1 = QLabel(desc)
        d1.setWordWrap(True)
        d1.setStyleSheet(f"color: {t.MUTED}; font-size: 12px;")
        inner.addWidget(t1)
        inner.addWidget(d1)
        btn = QPushButton(button_text)
        btn.setCursor(Qt.PointingHandCursor)
        btn.setStyleSheet(
            f"QPushButton {{ background: {t.PRIMARY}; color: {t.ON_PRIMARY}; border: none; "
            f"border-radius: {t.RADIUS_SM}px; padding: 8px 16px; font-weight: 600; max-width: 200px; }}"
            f"QPushButton:hover {{ background: {t.PRIMARY_DARK}; }}"
        )
        btn.clicked.connect(on_action)
        inner.addWidget(btn, 0, Qt.AlignLeft)
        lay.addLayout(inner, 1)


class StartPage(QWidget):
    def __init__(
        self,
        on_close: Callable[[], None],
        on_open: Callable[[], None],
        on_new: Callable[[], None],
        on_sales: Callable[[], None],
        on_merge: Callable[[], None],
        on_mapping: Callable[[], None],
        parent=None,
    ):
        super().__init__(parent)
        self._on_open = on_open
        self._on_reload_recent: Optional[Callable[[], None]] = None

        self.setAutoFillBackground(True)
        pal = self.palette()
        pal.setColor(self.backgroundRole(), Qt.transparent)
        self.setPalette(pal)

        root = QVBoxLayout(self)
        root.setContentsMargins(32, 32, 32, 32)
        root.setSpacing(20)

        top = QHBoxLayout()
        title = QLabel("欢迎使用 Excel 可视化工作流")
        title.setStyleSheet(f"font-size: 22px; font-weight: 700; color: {t.TEXT};")
        top.addWidget(title)
        top.addStretch(1)
        close_btn = QPushButton("关闭启动页")
        close_btn.setCursor(Qt.PointingHandCursor)
        close_btn.setStyleSheet(
            f"QPushButton {{ background: transparent; color: {t.PRIMARY}; border: 1px solid {t.OUTLINE}; "
            f"border-radius: {t.RADIUS_SM}px; padding: 6px 14px; }}"
            f"QPushButton:hover {{ background: rgba(63,81,181,0.08); }}"
        )
        close_btn.clicked.connect(on_close)
        top.addWidget(close_btn)
        root.addLayout(top)

        sub = QLabel("从模板开始，或打开已有工作流 JSON；在画布上连接节点后点击顶部「运行」。")
        sub.setWordWrap(True)
        sub.setStyleSheet(f"color: {t.MUTED}; font-size: 13px;")
        root.addWidget(sub)

        grid = QGridLayout()
        grid.setSpacing(16)
        cards = [
            (
                (63, 81, 181),
                "销售月报 · 模板填充",
                "读取 Excel → 格式标准化 → 按行写入 xlsx 模板。请补全各节点路径与映射。",
                "加载此模板",
                on_sales,
            ),
            (
                (0, 137, 123),
                "双表纵向合并",
                "两个「Excel 读取」→「纵向合并」→「写出 Excel」。适合对账或拼接两张表。",
                "加载此模板",
                on_merge,
            ),
            (
                (251, 140, 0),
                "旧版 mapping_config",
                "若仍使用旧 JSON 配置，可一键生成「读取 + 模板填充」节点，再检查路径。",
                "选择 mapping_config.json",
                on_mapping,
            ),
        ]
        for i, c in enumerate(cards):
            grid.addWidget(_Card(c[0], c[1], c[2], c[3], c[4]), i // 2, i % 2)
        root.addLayout(grid)

        row2 = QHBoxLayout()
        row2.setSpacing(16)
        row2.addWidget(
            _Card(
                (33, 150, 243),
                "打开本地工作流",
                "选择已保存的 workflow.json。",
                "浏览…",
                on_open,
            ),
            1,
        )
        row2.addWidget(
            _Card(
                (103, 58, 183),
                "新建空白",
                "清空画布，从左侧节点库点击添加节点。",
                "新建",
                on_new,
            ),
            1,
        )
        root.addLayout(row2)

        self._recent_frame = QFrame()
        self._recent_frame.setStyleSheet(
            f"QFrame {{ background: {t.SURFACE}; border: 1px solid {t.OUTLINE}; "
            f"border-radius: {t.RADIUS_MD}px; }}"
        )
        rfl = QVBoxLayout(self._recent_frame)
        rfl.setContentsMargins(16, 12, 16, 12)
        rfl.addWidget(QLabel("最近打开"))
        self._recent_inner = QVBoxLayout()
        self._recent_inner.setSpacing(4)
        rfl.addLayout(self._recent_inner)
        root.addWidget(self._recent_frame)

        root.addStretch(1)

        self._fill_recent()

    def _fill_recent(self):
        while self._recent_inner.count():
            item = self._recent_inner.takeAt(0)
            if item.widget():
                item.widget().deleteLater()
        paths = recent_workflows.load_recent()
        if not paths:
            lab = QLabel("暂无记录；打开工作流后会出现在这里。")
            lab.setStyleSheet(f"color: {t.MUTED};")
            self._recent_inner.addWidget(lab)
            return
        for p in paths:
            b = QPushButton(p)
            b.setFlat(True)
            b.setCursor(Qt.PointingHandCursor)
            b.setStyleSheet(
                f"QPushButton {{ text-align: left; color: {t.PRIMARY}; padding: 4px; }}"
                f"QPushButton:hover {{ text-decoration: underline; }}"
            )

            def _go(path=p):
                from excel_workflow.core.workflow_io import load_workflow

                mw = self.window()
                graph = getattr(mw, "_graph", None)
                if graph is None:
                    return
                try:
                    load_workflow(graph, path)
                    if hasattr(mw, "_append_log"):
                        mw._append_log(f"已加载: {path}")
                    recent_workflows.push_recent(path)
                    if hasattr(mw, "_hide_start_page"):
                        mw._hide_start_page()
                except Exception as e:
                    from PySide6.QtWidgets import QMessageBox

                    QMessageBox.critical(mw, "错误", str(e))

            b.clicked.connect(_go)
            self._recent_inner.addWidget(b)

    def refresh_recent(self):
        self._fill_recent()
