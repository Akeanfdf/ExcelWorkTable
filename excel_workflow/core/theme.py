# Material Design 3 inspired tokens (light) + legacy aliases for existing imports

# Surfaces
BG = "#F7F8FA"
SURFACE = "#FFFFFF"
SURFACE_VARIANT = "#EEF1F6"
SIDEBAR = SURFACE_VARIANT  # legacy
CARD = SURFACE
WHITE = "#FFFFFF"

# Brand
PRIMARY = "#3F51B5"
PRIMARY_DARK = "#303F9F"
ON_PRIMARY = "#FFFFFF"
SECONDARY = "#00897B"
ACCENT = PRIMARY  # legacy
ACCENT2 = SECONDARY  # legacy

# Semantic
WARNING = "#FB8C00"
DANGER = "#E53935"
SUCCESS = "#43A047"

# Text / stroke
TEXT = "#1F2937"
MUTED = "#5F6B7A"
OUTLINE = "#D7DCE3"
BORDER = OUTLINE  # legacy

# Shape (px)
RADIUS_SM = 6
RADIUS_MD = 12
RADIUS_LG = 16

# Shadow blur offsets (for QGraphicsDropShadowEffect if used later)
ELEV_1_BLUR = 8
ELEV_1_Y = 2
ELEV_2_BLUR = 16
ELEV_2_Y = 4

# Node category -> RGB (for BaseNode.set_color); matches CATEGORY_LABELS order in main_window
CATEGORY_RGB = {
    "excel_workflow.source": (63, 81, 181),  # Indigo
    "excel_workflow.clean": (0, 137, 123),  # Teal
    "excel_workflow.filter": (251, 140, 0),  # Amber
    "excel_workflow.merge": (233, 30, 99),  # Pink
    "excel_workflow.transform": (103, 58, 183),  # Deep purple
    "excel_workflow.template": (33, 150, 243),  # Blue
    "excel_workflow.export": (67, 160, 71),  # Green
    "excel_workflow.fileops": (84, 110, 122),  # Blue grey
    "excel_workflow.ai": (0, 172, 193),  # Cyan
    "excel_workflow.control": (121, 85, 72),  # Brown
}

RUNNING_NODE_RGB = (251, 192, 45)  # amber highlight while executing
FAILED_NODE_RGB = (229, 57, 53)
DONE_NODE_RGB = (67, 160, 71)
