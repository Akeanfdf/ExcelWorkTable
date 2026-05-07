"""注册所有节点类与执行器（导入副作用）。"""

from excel_workflow.nodes import ai  # noqa: F401
from excel_workflow.nodes import clean  # noqa: F401
from excel_workflow.nodes import control  # noqa: F401
from excel_workflow.nodes import export  # noqa: F401
from excel_workflow.nodes import fileops  # noqa: F401
from excel_workflow.nodes import filter_split  # noqa: F401
from excel_workflow.nodes import merge  # noqa: F401
from excel_workflow.nodes import source  # noqa: F401
from excel_workflow.nodes import template  # noqa: F401
from excel_workflow.nodes import transform  # noqa: F401


def register_all_nodes(graph):
    from excel_workflow.core.node_theme import apply_node_category_color
    from excel_workflow.core.registry import NODE_CLASSES

    graph.register_nodes(list(NODE_CLASSES))
    graph.node_created.connect(apply_node_category_color)
