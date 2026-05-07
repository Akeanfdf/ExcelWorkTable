from __future__ import annotations

from collections import deque
from typing import Any, Callable, Dict, List, Optional, Set, Tuple

from NodeGraphQt.nodes.backdrop_node import BackdropNode
from NodeGraphQt.nodes.base_node import BaseNode
from NodeGraphQt.nodes.port_node import PortInputNode, PortOutputNode

from excel_workflow.core.context import ExecContext
from excel_workflow.core.registry import get_runner, node_type_id


def _is_executable(n) -> bool:
    if not isinstance(n, BaseNode):
        return False
    if isinstance(n, (BackdropNode, PortInputNode, PortOutputNode)):
        return False
    return True


def _build_pred_succ(graph) -> Tuple[List, Dict[Any, Set[Any]], Dict[Any, Set[Any]]]:
    nodes = [n for n in graph.all_nodes() if _is_executable(n)]
    node_set = set(nodes)
    preds: Dict[Any, Set[Any]] = {n: set() for n in nodes}
    succ: Dict[Any, Set[Any]] = {n: set() for n in nodes}
    for n in nodes:
        for in_port in n.input_ports():
            for out_port in in_port.connected_ports():
                up = out_port.node()
                if up in node_set:
                    preds[n].add(up)
                    succ[up].add(n)
    return list(nodes), preds, succ


def topological_order(nodes: List, preds: Dict[Any, Set[Any]], succ: Dict[Any, Set[Any]]) -> List:
    pred = {n: set(preds[n]) for n in nodes}
    q = deque([n for n in nodes if not pred[n]])
    order: List = []
    while q:
        n = q.popleft()
        order.append(n)
        for s in succ[n]:
            pred[s].discard(n)
            if not pred[s]:
                q.append(s)
    if len(order) != len(nodes):
        raise RuntimeError("工作流存在环路或存在未连接的节点链")
    return order


def _collect_inputs(node, results: Dict[str, Dict[str, Any]]) -> Dict[str, Any]:
    inputs: Dict[str, Any] = {}
    for in_port in node.input_ports():
        name = in_port.name()
        cps = in_port.connected_ports()
        if not cps:
            inputs[name] = None
            continue
        out_port = cps[0]
        up = out_port.node()
        uid = str(up.id)
        if uid not in results:
            raise RuntimeError(f"上游节点 {up.name()} 无输出，请先连接或检查顺序")
        port_key = out_port.name()
        pack = results[uid]
        if port_key not in pack:
            raise RuntimeError(f"上游节点 {up.name()} 缺少输出端口 {port_key}")
        inputs[name] = pack[port_key]
    return inputs


def run_graph(
    graph,
    ctx: ExecContext,
    on_node_begin: Optional[Callable[[Any], None]] = None,
    on_node_end: Optional[Callable[[Any, Optional[Exception]], None]] = None,
) -> Dict[str, Dict[str, Any]]:
    """
    按 DAG 拓扑执行图中所有可执行节点。
    返回 node_id -> {port_name: value}
    """
    nodes, preds, succ = _build_pred_succ(graph)
    if not nodes:
        ctx.log("[引擎] 无可执行节点")
        return {}
    order = topological_order(nodes, preds, succ)
    results: Dict[str, Dict[str, Any]] = {}
    total = len(order)

    for i, node in enumerate(order):
        if ctx.cancelled():
            ctx.log("[引擎] 已取消")
            break
        uid = str(node.id)
        tid = node_type_id(node)
        ctx.progress(i / max(total, 1), f"运行 {node.name()} ({tid})")
        if on_node_begin:
            on_node_begin(node)
        runner = get_runner(tid)
        err: Optional[Exception] = None
        try:
            if not runner:
                ctx.log(f"[跳过] 无执行器: {tid}")
                results[uid] = {}
            else:
                inputs = _collect_inputs(node, results)
                out = runner(node, inputs, ctx) or {}
                results[uid] = out
        except Exception as e:
            err = e
            ctx.log(f"[错误] {node.name()}: {e}")
            results[uid] = {}
            if on_node_end:
                on_node_end(node, err)
            raise
        if on_node_end:
            on_node_end(node, None)
    ctx.progress(1.0, "完成")
    return results
