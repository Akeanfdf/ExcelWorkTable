"""workflow.json v2 读写（与 workflow_v2 同义，便于按目录结构导入）。"""

from excel_workflow.session.workflow_v2 import (
    load_workflow_v2,
    mapping_config_to_template_ops,
    save_workflow_v2,
)

__all__ = [
    "save_workflow_v2",
    "load_workflow_v2",
    "mapping_config_to_template_ops",
]
