from __future__ import annotations

from typing import Any, Callable, Dict, List, Optional, Type

# NodeGraphQt BaseNode 子类在启动时注册
NODE_CLASSES: List[Type] = []

# type_id -> runner(node, inputs, ctx) -> {port_name: value}
RUNNERS: Dict[str, Callable[..., Dict[str, Any]]] = {}


def register_node(node_cls: Type) -> Type:
    """装饰器：注册 UI 节点类。"""
    NODE_CLASSES.append(node_cls)
    return node_cls


def register_runner(type_id: str):
    """装饰器：注册可执行 runner。"""

    def deco(fn: Callable[..., Dict[str, Any]]):
        RUNNERS[type_id] = fn
        return fn

    return deco


def get_runner(type_id: str) -> Optional[Callable[..., Dict[str, Any]]]:
    return RUNNERS.get(type_id)


def node_type_id(node) -> str:
    """返回 NodeGraphQt 注册的节点类型字符串，如 excel_workflow.source.ExcelReader。"""
    t = getattr(node, "type_", None)
    if callable(t):
        try:
            t = t()
        except TypeError:
            pass
    if t:
        return str(t)
    cls = node.__class__
    ident = getattr(cls, "__identifier__", "")
    name = cls.__name__
    if ident:
        return f"{ident}.{name}"
    return name
