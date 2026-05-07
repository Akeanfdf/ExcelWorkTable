from __future__ import annotations

import os
import shutil
import uuid
from dataclasses import dataclass, field
from typing import Any, List

from excel_workflow.simple.paths import staging_sessions_dir


def _default_staging() -> str:
    return str(staging_sessions_dir() / uuid.uuid4().hex)


@dataclass
class CacheSession:
    """In-memory table + files staged under staging_dir until user exports."""

    staging_dir: str = field(default_factory=_default_staging)
    dataframe: Any | None = None
    artifacts: List[str] = field(default_factory=list)
    imported_files: List[str] = field(default_factory=list)
    op_history: List[dict] = field(default_factory=list)

    def reset(self) -> None:
        old = self.staging_dir
        try:
            if old and os.path.isdir(old):
                shutil.rmtree(old, ignore_errors=True)
        except OSError:
            pass
        self.staging_dir = _default_staging()
        self.dataframe = None
        self.artifacts.clear()
        self.imported_files.clear()
        self.op_history.clear()

    def clear_history_only(self) -> None:
        self.op_history.clear()

    def remember_op(self, type_id: str, props: dict) -> None:
        self.op_history.append({"type_id": type_id, "props": dict(props)})

    def add_artifact(self, path: str) -> None:
        path = os.path.abspath(path)
        if path not in self.artifacts:
            self.artifacts.append(path)
