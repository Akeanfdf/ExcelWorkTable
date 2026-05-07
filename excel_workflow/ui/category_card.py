"""Material-style collapsible node library (left dock)."""

from __future__ import annotations

from collections import OrderedDict
from typing import Callable, Dict, List, Tuple

from PySide6.QtCore import Qt
from PySide6.QtWidgets import (
    QFrame,
    QHBoxLayout,
    QLabel,
    QPushButton,
    QScrollArea,
    QSizePolicy,
    QToolButton,
    QVBoxLayout,
    QWidget,
)

from excel_workflow.core import theme as t
from excel_workflow.core.registry import NODE_CLASSES


def _grouped_nodes() -> "OrderedDict[str, List[Tuple[str, str]]]":
    groups: "OrderedDict[str, List[Tuple[str, str]]]" = OrderedDict()
    for cls in NODE_CLASSES:
        ident = getattr(cls, "__identifier__", "")
        if not ident:
            continue
        display = getattr(cls, "NODE_NAME", cls.__name__)
        tid = f"{ident}.{cls.__name__}"
        groups.setdefault(ident, []).append((tid, display))
    return groups


class _CategoryBlock(QFrame):
    def __init__(
        self,
        _category_id: str,
        title: str,
        dot_rgb: Tuple[int, int, int],
        entries: List[Tuple[str, str]],
        on_pick: Callable[[str], None],
        parent=None,
    ):
        super().__init__(parent)
        self.setObjectName("CategoryBlock")
        self.setStyleSheet(
            f"#CategoryBlock {{ background: {t.SURFACE}; border: 1px solid {t.OUTLINE}; "
            f"border-radius: {t.RADIUS_MD}px; }}"
        )
        root = QVBoxLayout(self)
        root.setContentsMargins(0, 0, 0, 6)
        root.setSpacing(0)

        head = QHBoxLayout()
        head.setContentsMargins(10, 8, 8, 8)
        self._toggle = QToolButton()
        self._toggle.setText("▼")
        self._toggle.setAutoRaise(True)
        self._toggle.setToolTip("展开 / 折叠")
        self._toggle.clicked.connect(self._on_toggle)
        self._expanded = True

        dot = QLabel()
        dot.setFixedSize(10, 10)
        dot.setStyleSheet(
            f"background: rgb({dot_rgb[0]},{dot_rgb[1]},{dot_rgb[2]}); "
            f"border-radius: 5px;"
        )
        lab = QLabel(title)
        lab.setStyleSheet(f"font-weight: 600; color: {t.TEXT};")

        head.addWidget(self._toggle)
        head.addWidget(dot)
        head.addWidget(lab, 1)
        root.addLayout(head)

        self._body = QWidget()
        bl = QVBoxLayout(self._body)
        bl.setContentsMargins(4, 0, 4, 4)
        bl.setSpacing(2)
        for tid, name in entries:
            row = QPushButton(f"  {name}")
            row.setFlat(True)
            row.setCursor(Qt.PointingHandCursor)
            row.setStyleSheet(
                f"QPushButton {{ text-align: left; padding: 6px 10px; border: none; "
                f"border-radius: {t.RADIUS_SM}px; color: {t.TEXT}; }}"
                f"QPushButton:hover {{ background: {t.SURFACE_VARIANT}; }}"
            )
            row.clicked.connect(lambda _=False, x=tid: on_pick(x))
            bl.addWidget(row)
        root.addWidget(self._body)

    def _on_toggle(self):
        self._expanded = not self._expanded
        self._body.setVisible(self._expanded)
        self._toggle.setText("▼" if self._expanded else "▶")


class NodeLibraryPanel(QWidget):
    """Scrollable list of category cards; click row to create node on graph."""

    def __init__(self, graph, category_labels: Dict[str, str], on_create: Callable[[str], None]):
        super().__init__()
        self._graph = graph
        self._on_create = on_create
        self.setMinimumWidth(200)
        self.setSizePolicy(QSizePolicy.Preferred, QSizePolicy.Expanding)

        outer = QVBoxLayout(self)
        outer.setContentsMargins(0, 0, 0, 0)
        outer.setSpacing(8)

        title = QLabel("节点库")
        title.setStyleSheet(f"font-size: 15px; font-weight: 700; color: {t.TEXT}; padding: 4px 6px;")
        outer.addWidget(title)

        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        scroll.setFrameShape(QFrame.NoFrame)
        scroll.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        body = QWidget()
        lay = QVBoxLayout(body)
        lay.setSpacing(10)
        lay.setContentsMargins(0, 0, 4, 0)

        grouped = _grouped_nodes()
        for ident, entries in grouped.items():
            label = category_labels.get(ident, ident.split(".")[-1])
            rgb = t.CATEGORY_RGB.get(ident, (99, 99, 99))
            block = _CategoryBlock(ident, label, rgb, entries, self._pick)
            lay.addWidget(block)
        lay.addStretch(1)
        scroll.setWidget(body)
        outer.addWidget(scroll, 1)

    def _pick(self, type_id: str):
        self._on_create(type_id)
