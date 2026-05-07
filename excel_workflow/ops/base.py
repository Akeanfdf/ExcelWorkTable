from __future__ import annotations

from abc import ABC, abstractmethod
from typing import TYPE_CHECKING, Any, Dict

if TYPE_CHECKING:
    from excel_workflow.core.context import ExecContext
    from excel_workflow.session.state import LinearSession


class BaseOp(ABC):
    """操作抽象：子类实现 op_key 与 apply。"""

    op_key: str = ""

    @abstractmethod
    def apply(
        self, session: "LinearSession", params: Dict[str, Any], ctx: "ExecContext"
    ) -> str:
        """执行并返回时间线摘要。"""
