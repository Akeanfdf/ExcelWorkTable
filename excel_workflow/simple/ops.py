from __future__ import annotations

from typing import Any, Dict

from excel_workflow.core.context import ExecContext
from excel_workflow.core.registry import get_runner
from excel_workflow.simple.mock_node import PropertyBagNode


def run_registered(type_id: str, props: dict, inputs: Dict[str, Any], ctx: ExecContext) -> Dict[str, Any]:
    runner = get_runner(type_id)
    if not runner:
        raise ValueError(f"未注册执行器: {type_id}")
    node = PropertyBagNode(props)
    return runner(node, inputs, ctx) or {}
