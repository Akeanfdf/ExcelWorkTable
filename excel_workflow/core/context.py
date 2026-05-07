from __future__ import annotations

import threading
import time
from dataclasses import dataclass, field
from typing import Any, Callable, Dict, List, Optional


LogFn = Callable[[str], None]
ProgressFn = Callable[[float, str], None]


@dataclass
class ExecContext:
    """节点执行上下文：日志、进度、取消。"""

    log: LogFn
    progress: ProgressFn
    cancel_event: threading.Event = field(default_factory=threading.Event)
    node_outputs: Dict[str, Dict[str, Any]] = field(default_factory=dict)

    def cancelled(self) -> bool:
        return self.cancel_event.is_set()

    def sleep_step(self, seconds: float = 0) -> None:
        if seconds:
            time.sleep(seconds)
