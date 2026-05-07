from __future__ import annotations

import uuid
import os
import shutil
from dataclasses import dataclass, field
from datetime import datetime, timezone
from pathlib import Path
from typing import Any, List, Optional

from excel_workflow.utils.cache_dir import staging_sessions_dir


def _utc_now() -> str:
    return datetime.now(timezone.utc).replace(microsecond=0).isoformat()


@dataclass
class AppliedOp:
    op_type: str
    params: dict
    timestamp: str = field(default_factory=_utc_now)
    result_summary: str = ""


@dataclass
class LinearSession:
    """三段式线性流程会话（导入 › 操作 › 导出）。"""

    phase: int = 0  # 0 导入 1 操作 2 导出
    source_files: List[Path] = field(default_factory=list)
    staging_dir: Optional[Path] = None
    ops_history: List[AppliedOp] = field(default_factory=list)
    current_files: List[Path] = field(default_factory=list)
    dataframe: Any = None
    artifacts: List[str] = field(default_factory=list)
    is_dirty: bool = False
    legacy_mapping_config: Optional[dict] = None

    def staging_dir_str(self) -> str:
        return str(self.staging_dir) if self.staging_dir else ""

    def add_artifact(self, path: str) -> None:
        path = os.path.abspath(path)
        if path not in self.artifacts:
            self.artifacts.append(path)

    def new_staging(self) -> None:
        sid = uuid.uuid4().hex
        base = staging_sessions_dir()
        self.staging_dir = base / sid
        self.staging_dir.mkdir(parents=True, exist_ok=True)

    def reset_for_new_import(self) -> None:
        self._clear_staging_disk()
        self.phase = 0
        self.source_files.clear()
        self.ops_history.clear()
        self.current_files.clear()
        self.dataframe = None
        self.artifacts.clear()
        self.is_dirty = False
        self.legacy_mapping_config = None
        self.staging_dir = None

    def _clear_staging_disk(self) -> None:
        d = self.staging_dir
        if d and d.is_dir():
            try:
                shutil.rmtree(d, ignore_errors=True)
            except OSError:
                pass

    def reset_staging_keep_sources(self) -> None:
        """撤销/重放前清空产物与缓存目录，保留 source_files。"""
        self._clear_staging_disk()
        self.new_staging()
        self.ops_history.clear()
        self.current_files = list(self.source_files)
        self.dataframe = None
        self.artifacts.clear()
        self.is_dirty = True

    def to_serializable_ops(self) -> List[dict]:
        return [
            {"op_type": o.op_type, "params": dict(o.params)} for o in self.ops_history
        ]

    @staticmethod
    def from_ops_list(ops: List[dict]) -> List[AppliedOp]:
        out = []
        for o in ops:
            out.append(
                AppliedOp(
                    op_type=str(o.get("op_type", "")),
                    params=dict(o.get("params") or {}),
                    timestamp=str(o.get("timestamp") or _utc_now()),
                    result_summary=str(o.get("result_summary") or ""),
                )
            )
        return out
