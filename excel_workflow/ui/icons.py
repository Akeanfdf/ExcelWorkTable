"""Small SVG icons as QIcon (no external assets required)."""

from PySide6.QtCore import QByteArray
from PySide6.QtGui import QIcon, QImage, QPainter, QPixmap
from PySide6.QtWidgets import QStyle, QWidget

try:
    from PySide6.QtSvg import QSvgRenderer
except ImportError:  # pragma: no cover
    QSvgRenderer = None  # type: ignore


def _icon_from_svg(svg_xml: str, size: int = 22) -> QIcon:
    if QSvgRenderer is None:
        return QIcon()
    data = QByteArray(svg_xml.encode("utf-8"))
    renderer = QSvgRenderer(data)
    img = QImage(size, size, QImage.Format_ARGB32_Premultiplied)
    img.fill(0)
    painter = QPainter(img)
    renderer.render(painter)
    painter.end()
    # PySide6: QIcon() may not accept QImage directly; use QPixmap.
    return QIcon(QPixmap.fromImage(img))


def icon_open_folder() -> QIcon:
    return _icon_from_svg(
        """<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24" fill="#5F6B7A">
        <path d="M10 4H4c-1.1 0-1.99.9-1.99 2L2 18c0 1.1.9 2 2 2h16c1.1 0 2-.9 2-2V8c0-1.1-.9-2-2-2h-8l-2-2z"/>
        </svg>"""
    )


def icon_save() -> QIcon:
    return _icon_from_svg(
        """<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24" fill="#5F6B7A">
        <path d="M17 3H5c-1.11 0-2 .9-2 2v14c0 1.1.89 2 2 2h14c1.1 0 2-.9 2-2V7l-4-4zm-5 16c-1.66 0-3-1.34-3-3s1.34-3 3-3 3 1.34 3 3-1.34 3-3 3zm3-10H5V5h10v4z"/>
        </svg>"""
    )


def icon_import() -> QIcon:
    return _icon_from_svg(
        """<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24" fill="#5F6B7A">
        <path d="M19 9h-4V3H9v6H5l7 7 7-7zM5 18v2h14v-2H5z"/>
        </svg>"""
    )


def icon_play() -> QIcon:
    return _icon_from_svg(
        """<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24" fill="#FFFFFF">
        <path d="M8 5v14l11-7z"/>
        </svg>"""
    )


def icon_stop() -> QIcon:
    return _icon_from_svg(
        """<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24" fill="#5F6B7A">
        <path d="M6 6h12v12H6z"/>
        </svg>"""
    )


def icon_watch() -> QIcon:
    return _icon_from_svg(
        """<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24" fill="#5F6B7A">
        <path d="M12 4.5C7 4.5 2.73 7.61 1 12c1.73 4.39 6 7.5 11 7.5s9.27-3.11 11-7.5c-1.73-4.39-6-7.5-11-7.5zM12 17c-2.76 0-5-2.24-5-5s2.24-5 5-5 5 2.24 5 5-2.24 5-5 5zm0-8c-1.66 0-3 1.34-3 3s1.34 3 3 3 3-1.34 3-3-1.34-3-3-3z"/>
        </svg>"""
    )


def icon_schedule() -> QIcon:
    return _icon_from_svg(
        """<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24" fill="#5F6B7A">
        <path d="M11.99 2C6.47 2 2 6.48 2 12s4.47 10 9.99 10C17.52 22 22 17.52 22 12S17.52 2 11.99 2zM12 20c-4.42 0-8-3.58-8-8s3.58-8 8-8 8 3.58 8 8-3.58 8-8 8zm.5-13H11v6l5.25 3.15.75-1.23-4.5-2.67z"/>
        </svg>"""
    )


def fallback_style_icon(widget: QWidget, standard: QStyle.StandardPixmap) -> QIcon:
    return widget.style().standardIcon(standard)
