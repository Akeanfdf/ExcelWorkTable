"""
Excel 可视化工作流 — 应用入口。

运行（在项目根目录 excel表格）::

    pip install -r requirements-workflow.txt
    python -m excel_workflow.app
"""

from __future__ import annotations

import os
import sys

_MIN_PY = (3, 10)


def main():
    if sys.version_info < _MIN_PY:
        ver = ".".join(map(str, sys.version_info[:3]))
        print(
            f"当前 Python 为 {ver}。本程序需要 Python {_MIN_PY[0]}.{_MIN_PY[1]}+（PySide6 / pandas 2 不支持 3.7）。\n"
            "请安装 Python 3.10+ 后执行，例如：\n"
            "  py -3.11 -m pip install -r requirements-workflow.txt\n"
            "  py -3.11 -m excel_workflow.app",
            file=sys.stderr,
        )
        sys.exit(1)

    os.environ.setdefault("QT_API", "pyside6")

    try:
        from PySide6.QtWidgets import QApplication
    except ModuleNotFoundError:
        print(
            "未找到 PySide6。请先在同一 Python 下安装依赖：\n"
            "  py -3.11 -m pip install -r requirements-workflow.txt",
            file=sys.stderr,
        )
        sys.exit(1)

    from excel_workflow.ui.main_window import MainWindow

    app = QApplication(sys.argv)
    win = MainWindow()
    win.show()
    sys.exit(app.exec())


if __name__ == "__main__":
    main()
