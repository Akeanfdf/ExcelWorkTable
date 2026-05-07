"""Built-in workflow presets: load examples/templates/*.json when present, else build in code."""

from __future__ import annotations

from pathlib import Path

from excel_workflow.core.workflow_io import load_workflow


def templates_dir() -> Path:
    return Path(__file__).resolve().parents[2] / "examples" / "templates"


def apply_sales_report_preset(graph) -> None:
    p = templates_dir() / "sales_report.json"
    if p.is_file():
        load_workflow(graph, str(p))
        return
    graph.clear_session()
    r = graph.create_node("excel_workflow.source.ExcelReader", pos=(0, 0))
    c = graph.create_node("excel_workflow.clean.FormatStandardize", pos=(260, 0))
    t = graph.create_node("excel_workflow.template.TemplateMapper", pos=(520, 0))
    r.output(0).connect_to(c.input(0))
    c.output(0).connect_to(t.input(0))
    r.set_property("file_path", "")
    r.set_property("sheet_name", "")
    c.set_property("columns", "")
    c.set_property("strip", True)
    t.set_property("template_path", "")
    t.set_property("output_dir", "")
    t.set_property("name_column", "")
    t.set_property("sheet_name", "Sheet1")
    t.set_property("mapping_json", '[["列名","A1"]]')


def apply_merge_folder_preset(graph) -> None:
    p = templates_dir() / "merge_folder.json"
    if p.is_file():
        load_workflow(graph, str(p))
        return
    graph.clear_session()
    a = graph.create_node("excel_workflow.source.ExcelReader", pos=(0, -120))
    b = graph.create_node("excel_workflow.source.ExcelReader", pos=(0, 120))
    m = graph.create_node("excel_workflow.merge.VerticalMerge", pos=(300, 0))
    w = graph.create_node("excel_workflow.export.ExcelWriter", pos=(560, 0))
    a.output(0).connect_to(m.input(0))
    b.output(0).connect_to(m.input(1))
    m.output(0).connect_to(w.input(0))
    a.set_property("file_path", "")
    a.set_property("sheet_name", "")
    b.set_property("file_path", "")
    b.set_property("sheet_name", "")
    m.set_property("ignore_index", True)
    w.set_property("output_path", "")
    w.set_property("sheet_name", "Sheet1")
