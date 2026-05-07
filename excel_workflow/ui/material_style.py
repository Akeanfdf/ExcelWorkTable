"""Global QSS for Material-style shell (PySide6)."""

from excel_workflow.core import theme as t


def build_app_stylesheet() -> str:
    r = t.RADIUS_SM
    return f"""
    QWidget {{
      font-family: "Segoe UI Variable", "Microsoft YaHei UI", "Roboto", sans-serif;
      font-size: 13px;
      color: {t.TEXT};
      background: {t.BG};
    }}
    QMainWindow {{ background: {t.BG}; }}
    QMenuBar {{
      background: {t.SURFACE};
      color: {t.TEXT};
      border-bottom: 1px solid {t.OUTLINE};
      padding: 2px 6px;
    }}
    QMenuBar::item:selected {{ background: {t.SURFACE_VARIANT}; }}
    QToolBar {{
      background: {t.SURFACE};
      border: none;
      border-bottom: 1px solid {t.OUTLINE};
      spacing: 8px;
      padding: 6px 10px;
    }}
    QToolButton {{
      background: transparent;
      border: 1px solid {t.OUTLINE};
      border-radius: {r}px;
      padding: 6px 10px;
      color: {t.TEXT};
    }}
    QToolButton:hover {{
      background: rgba(63, 81, 181, 0.08);
    }}
    QToolButton:pressed {{
      background: rgba(63, 81, 181, 0.14);
    }}
    QToolButton#PrimaryRun {{
      background: {t.PRIMARY};
      color: {t.ON_PRIMARY};
      border: none;
      font-weight: 600;
      padding: 8px 20px;
      border-radius: {r}px;
    }}
    QToolButton#PrimaryRun:hover {{
      background: {t.PRIMARY_DARK};
    }}
    QStatusBar {{
      background: {t.SURFACE};
      border-top: 1px solid {t.OUTLINE};
      color: {t.MUTED};
    }}
    QSplitter::handle {{
      background: {t.OUTLINE};
      width: 1px;
    }}
    QTabWidget::pane {{
      border: 1px solid {t.OUTLINE};
      border-radius: {t.RADIUS_MD}px;
      top: -1px;
      background: {t.SURFACE};
    }}
    QTabBar::tab {{
      padding: 8px 18px;
      margin-right: 2px;
      border-top-left-radius: {r}px;
      border-top-right-radius: {r}px;
      background: {t.SURFACE_VARIANT};
      color: {t.TEXT};
    }}
    QTabBar::tab:selected {{
      background: {t.SURFACE};
      border: 1px solid {t.OUTLINE};
      border-bottom-color: {t.SURFACE};
      font-weight: 600;
    }}
    QScrollBar:vertical {{
      width: 8px;
      background: transparent;
      margin: 2px;
    }}
    QScrollBar::handle:vertical {{
      background: {t.OUTLINE};
      min-height: 24px;
      border-radius: 4px;
    }}
    QScrollBar:horizontal {{
      height: 8px;
      background: transparent;
    }}
    QScrollBar::handle:horizontal {{
      background: {t.OUTLINE};
      min-width: 24px;
      border-radius: 4px;
    }}
    QFrame#PropsCard, QFrame#LogCard {{
      background: {t.SURFACE};
      border: 1px solid {t.OUTLINE};
      border-radius: {t.RADIUS_MD}px;
    }}
    QFrame#PropsCard QLineEdit, QFrame#PropsCard QSpinBox, QFrame#PropsCard QComboBox,
    QFrame#PropsCard QTextEdit {{
      border: 1px solid {t.OUTLINE};
      border-radius: {r}px;
      padding: 4px 8px;
      background: {t.SURFACE};
    }}
    QFrame#PropsCard QTabWidget::pane {{
      border: 1px solid {t.OUTLINE};
      border-radius: {r}px;
      top: -1px;
      background: {t.SURFACE};
    }}
    QFrame#PropsCard QTabBar::tab {{
      padding: 6px 12px;
      margin-right: 2px;
      border-top-left-radius: {r}px;
      border-top-right-radius: {r}px;
      background: {t.SURFACE_VARIANT};
    }}
    QFrame#PropsCard QTabBar::tab:selected {{
      background: {t.SURFACE};
      border: 1px solid {t.OUTLINE};
      border-bottom: none;
    }}
    QProgressBar {{
      border: none;
      background: {t.SURFACE_VARIANT};
      height: 4px;
      border-radius: 2px;
    }}
    QProgressBar::chunk {{
      background: {t.PRIMARY};
      border-radius: 2px;
    }}
    QTextEdit#LogConsole {{
      background: #0F172A;
      color: #94A3B8;
      font-family: Consolas, monospace;
      font-size: 11px;
      border: 1px solid #1E293B;
      border-radius: {t.RADIUS_MD}px;
    }}
    QTextEdit#LogConsole QScrollBar:vertical,
    QTextEdit#LogConsole QScrollBar:horizontal {{
      width: 12px;
      height: 12px;
      min-width: 12px;
      min-height: 12px;
    }}
    QTextEdit#DataPreviewEdit {{
      font-family: Consolas, "Cascadia Mono", "Microsoft YaHei UI", monospace;
      font-size: 11px;
    }}
    QTextEdit#DataPreviewEdit QScrollBar:vertical,
    QTextEdit#DataPreviewEdit QScrollBar:horizontal {{
      width: 12px;
      height: 12px;
      min-width: 12px;
      min-height: 12px;
    }}
    QWidget#LinearMainInner {{
      background: {t.BG};
    }}
    QWidget#LinearContentHost {{
      background: {t.SURFACE};
      border: 1px solid {t.OUTLINE};
      border-radius: {t.RADIUS_MD}px;
    }}
    QFrame#LinearHostTop, QFrame#LinearHostBottom {{
      background: {t.SURFACE};
      border: 1px solid {t.OUTLINE};
      border-radius: {t.RADIUS_MD}px;
    }}
    QStackedWidget#LinearStack {{
      background: transparent;
      border: none;
    }}
    QFrame#LinearCard {{
      background: {t.SURFACE};
      border: 1px solid {t.OUTLINE};
      border-radius: {t.RADIUS_MD}px;
    }}
    QGroupBox#LinearCard {{
      font-weight: 600;
      background: {t.SURFACE};
      border: 1px solid {t.OUTLINE};
      border-radius: {t.RADIUS_MD}px;
      margin-top: 10px;
      padding-top: 8px;
    }}
    QGroupBox#LinearCard::title {{
      subcontrol-origin: margin;
      left: 10px;
      padding: 0 6px;
      color: {t.TEXT};
    }}
    QWidget#LinearContentHost QPushButton {{
      background: {t.SURFACE_VARIANT};
      border: 1px solid {t.OUTLINE};
      border-radius: {r}px;
      padding: 6px 12px;
      color: {t.TEXT};
    }}
    QWidget#LinearContentHost QPushButton:hover {{
      background: rgba(63, 81, 181, 0.08);
      border-color: {t.PRIMARY};
    }}
    QWidget#LinearContentHost QPushButton:pressed {{
      background: rgba(63, 81, 181, 0.14);
    }}
    QWidget#LinearContentHost QLineEdit, QWidget#LinearContentHost QTextEdit,
    QWidget#LinearContentHost QComboBox {{
      border: 1px solid {t.OUTLINE};
      border-radius: {r}px;
      padding: 4px 8px;
      background: {t.SURFACE};
    }}
    QWidget#LinearContentHost QListWidget {{
      border: 1px solid {t.OUTLINE};
      border-radius: {r}px;
      background: {t.SURFACE};
    }}
    QWidget#LinearContentHost QScrollArea {{
      background: transparent;
    }}
    QSplitter#MainLogSplitter::handle {{
      background: {t.OUTLINE};
    }}
    """
