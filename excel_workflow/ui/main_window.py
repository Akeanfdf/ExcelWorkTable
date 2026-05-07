from __future__ import annotations

from PySide6.QtCore import QSize, Qt, Slot
from PySide6.QtWidgets import (
    QFrame,
    QLabel,
    QMainWindow,
    QProgressBar,
    QSplitter,
    QStyle,
    QTextEdit,
    QToolBar,
    QToolButton,
    QVBoxLayout,
    QWidget,
)

import excel_workflow.nodes  # noqa: F401  — 加载节点模块以注册所有 runner

from excel_workflow.core import theme
from excel_workflow.ui import material_style
from excel_workflow.ui.linear_host import LinearFlowHost
from excel_workflow.ui.linear_manual_dialog import show_full_manual


class MainWindow(QMainWindow):
    """单窗口：上方线性流程 + 底部日志区（可拖动分隔条调整高度）。"""

    _MAIN_STRETCH = 22
    _LOG_STRETCH = 7

    def __init__(self):
        super().__init__()
        self.setWindowTitle("Excel 表单处理器")
        self.resize(1280, 800)

        self._progress = QProgressBar()
        self._progress.setRange(0, 100)
        self._progress.setValue(0)
        self._progress.setTextVisible(False)
        self._progress.setFixedHeight(4)
        self._progress.hide()

        self._build_ui()
        self._build_app_bar()
        self.setStyleSheet(material_style.build_app_stylesheet())

    def _build_ui(self) -> None:
        central = QWidget()
        root = QVBoxLayout(central)
        root.setContentsMargins(0, 0, 0, 0)
        root.setSpacing(0)

        root.addWidget(self._progress)

        inner = QWidget()
        inner.setObjectName("LinearMainInner")
        lay = QVBoxLayout(inner)
        lay.setContentsMargins(8, 8, 8, 0)
        lay.setSpacing(8)

        self._linear_host = LinearFlowHost(self._append_log)
        self._linear_host.setObjectName("LinearContentHost")
        lay.addWidget(self._linear_host, 1)

        log_wrap = QFrame()
        log_wrap.setObjectName("LogCard")
        log_wrap.setMinimumHeight(96)
        log_lay = QVBoxLayout(log_wrap)
        log_lay.setContentsMargins(8, 6, 8, 8)
        log_lay.setSpacing(4)
        log_lay.addWidget(QLabel("日志"))
        self._console = QTextEdit()
        self._console.setObjectName("LogConsole")
        self._console.setReadOnly(True)
        self._console.setVerticalScrollBarPolicy(
            Qt.ScrollBarPolicy.ScrollBarAsNeeded
        )
        self._console.setHorizontalScrollBarPolicy(
            Qt.ScrollBarPolicy.ScrollBarAsNeeded
        )
        log_lay.addWidget(self._console, 1)

        split = QSplitter(Qt.Orientation.Vertical)
        split.setObjectName("MainLogSplitter")
        split.setHandleWidth(6)
        split.setChildrenCollapsible(False)
        split.addWidget(inner)
        split.addWidget(log_wrap)
        split.setStretchFactor(0, self._MAIN_STRETCH)
        split.setStretchFactor(1, self._LOG_STRETCH)
        h = max(self.height() - 40, 400)
        split.setSizes([h * self._MAIN_STRETCH // (self._MAIN_STRETCH + self._LOG_STRETCH), h * self._LOG_STRETCH // (self._MAIN_STRETCH + self._LOG_STRETCH)])
        root.addWidget(split, 1)

        self.setCentralWidget(central)

        self._status = QLabel("就绪")
        self.statusBar().addWidget(self._status)

    def _build_app_bar(self) -> None:
        tb = QToolBar("AppBar")
        tb.setObjectName("AppBar")
        tb.setMovable(False)
        tb.setIconSize(QSize(22, 22))
        self.addToolBar(tb)

        hlp = QToolButton()
        hlp.setIcon(self.style().standardIcon(QStyle.StandardPixmap.SP_DialogHelpButton))
        hlp.setToolTip("打开《功能说明书》全文")
        hlp.clicked.connect(lambda: show_full_manual(self))
        tb.addWidget(hlp)

        lab = QLabel("  功能说明书（全文）")
        lab.setStyleSheet(f"color: {theme.MUTED};")
        tb.addWidget(lab)

    @Slot(str)
    def _append_log(self, msg: str) -> None:
        c = getattr(self, "_console", None)
        if c is not None:
            c.append(msg)
