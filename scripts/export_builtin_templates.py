"""One-shot: write examples/templates/sales_report.json and merge_folder.json (needs display/Qt)."""

from __future__ import annotations

import sys
from pathlib import Path

_ROOT = Path(__file__).resolve().parents[1]
if str(_ROOT) not in sys.path:
    sys.path.insert(0, str(_ROOT))


def main() -> None:
    from PySide6.QtWidgets import QApplication

    from NodeGraphQt import NodeGraph

    from excel_workflow.nodes import register_all_nodes
    from excel_workflow.ui import template_presets

    app = QApplication(sys.argv)
    out = _ROOT / "examples" / "templates"
    out.mkdir(parents=True, exist_ok=True)

    g = NodeGraph()
    register_all_nodes(g)
    template_presets.apply_sales_report_preset(g)
    g.save_session(str(out / "sales_report.json"))

    g.clear_session()
    template_presets.apply_merge_folder_preset(g)
    g.save_session(str(out / "merge_folder.json"))
    print("Wrote:", out / "sales_report.json")
    print("Wrote:", out / "merge_folder.json")


if __name__ == "__main__":
    main()
