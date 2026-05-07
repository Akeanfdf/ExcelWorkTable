"""Apply Material category colors to graph nodes."""

from __future__ import annotations

from typing import TYPE_CHECKING

from NodeGraphQt.nodes.backdrop_node import BackdropNode
from NodeGraphQt.nodes.base_node import BaseNode
from NodeGraphQt.nodes.port_node import PortInputNode, PortOutputNode

from excel_workflow.core import theme
from excel_workflow.core.registry import node_type_id

if TYPE_CHECKING:
    pass


def _is_executable_node(n) -> bool:
    if not isinstance(n, BaseNode):
        return False
    if isinstance(n, (BackdropNode, PortInputNode, PortOutputNode)):
        return False
    return True


def apply_node_category_color(node) -> None:
    if not _is_executable_node(node):
        return
    tid = node_type_id(node)
    if not tid or "." not in tid:
        return
    ident = tid.rsplit(".", 1)[0]
    rgb = theme.CATEGORY_RGB.get(ident)
    if rgb:
        node.set_color(rgb[0], rgb[1], rgb[2])


def apply_all_node_colors(graph) -> None:
    for n in graph.all_nodes():
        apply_node_category_color(n)
