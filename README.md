# Excel 工作流与批量表单工具

## 功能简介

- **可视化工作流**：对 xlsx / xls / csv 做「导入 → 多步操作 → 导出」；支持多文件、步骤链、保存/载入 `workflow.json`、多表套用同一套步骤、内置说明书与帮助。
- **线性界面常用能力**：数据整理（标准化、去重、空值、筛选、排序、选列、替换、重命名）、合并文件、填充 Excel 模板、写出 xlsx、ZIP 打包；引擎侧还可扩展合并/透视/PDF/打印/文件批处理及可选 AI、PDF 表格提取等（见 `excel_workflow/nodes`）。
- **批量表单生成器（旧版）**：按数据源**每一行**填入模板单元格，批量生成多份 Excel，配置存 `mapping_config.json`。

| 程序 | 典型用途 |
|------|----------|
| 工作流主程序 | 规则复杂、多步、需复用流程或批量处理多张表 |
| Tk 生成器 | 规则固定、逐行套打模板即可 |

---

本仓库包含两套可独立使用的桌面程序：

1. **Excel 可视化工作流**（主程序）：基于 **PySide6** 的「导入 → 操作 → 导出」线性流程，对表格做多步处理、保存 `workflow.json`（v2）、批量对多表重放步骤并导出。窗口标题为「Excel 表单处理器」。
2. **Excel 批量表单生成器**（旧版）：基于 **Tkinter** 的单工具，按数据源行填充模板单元格，配置保存在 `mapping_config.json`。

远程仓库：[Akeanfdf/ExcelWorkTable](https://github.com/Akeanfdf/ExcelWorkTable)

---

## 环境要求

- **Python 3.10+**（工作流依赖 PySide6、pandas 2 等，不支持 3.7）
- Windows 上推荐用启动器指定版本，例如：`py -3.11`

---

## 一、Excel 可视化工作流（推荐）

### 安装依赖

在项目根目录执行（与运行时使用**同一**解释器）：

```bash
py -3.11 -m pip install -r requirements-workflow.txt
```

依赖见 `requirements-workflow.txt`（含 PySide6、QtPy、NodeGraphQt、pandas、openpyxl、pdfplumber、watchdog、APScheduler、pyzipper、openai、pytesseract、Pillow；Windows 另含 pywin32）。

### 启动方式（任选其一）

```bash
py -3.11 -m excel_workflow.app
```

```bash
py -3.11 run_workflow.py
```

入口模块：`excel_workflow/app.py`；打包入口脚本：`run_workflow.py`。

### 界面与流程概要

| 区域 | 说明 |
|------|------|
| 启动页 | 空白流程、打开最近、示例卡片等（Material 风格） |
| 线性主流程 | **导入**：拖放/多选 xlsx、xls、csv，或扫描文件夹勾选；可读旧版 `mapping_config.json` 预填；可载入 `workflow.json` v2。**操作**：左侧功能块追加步骤，右侧时间线填参数并「应用」；支持追加导入、多表批量重放等。**导出**：勾选项、输出目录与格式。 |
| 顶栏 | 「导入 / 操作 / 导出」为**阶段提示**（不可点击跳步），须按顺序完成。 |
| 底栏 | **保存流程**（导出 `workflow.json`）、**导出文件**、**批量导出**。 |
| 日志 | 主窗口下方可调整高度的日志区。 |
| 帮助 | 工具栏 **?** 打开《功能说明书》全文；各处的 **?** 与内嵌说明来自 `excel_workflow/ui/linear_feature_docs.py`。 |

业务逻辑由 `excel_workflow/nodes` 注册的运行器驱动，与界面步骤一致。

### 数据与缓存目录

默认在非 C 盘首个可用盘符下使用 `.excel_workflow`（如 `D:\.excel_workflow\staging` 等）。若需固定位置，可设置环境变量 **`EXCEL_WORKFLOW_DATA_ROOT`**（程序会在其下创建 `staging`、`simple_templates`、`recent.json` 等）。

### AI 功能（可选）

将 `examples/secrets.example.json` 复制到用户目录下的 `secrets.json`（见示例文件内说明），或设置环境变量 **`OPENAI_API_KEY`**。

### 打包为 exe（工作流）

```bash
pyinstaller Excel工作流.spec
```

生成物在 `dist/`（该目录已在 `.gitignore` 中忽略）。

---

## 二、Excel 批量表单生成器（Tk 旧版）

仅需较轻依赖：

```bash
pip install openpyxl pandas
python main.py
```

### 界面页签

| 页面 | 功能 |
|------|------|
| 文件配置 | 数据源、模板、输出目录、命名列与工作表 |
| 字段映射 | 数据列 → 模板单元格 |
| 开始生成 | 摘要、进度、日志 |
| 使用说明 | 内置操作指引 |

配置自动保存为项目目录下的 **`mapping_config.json`**。

### 打包为 exe（旧版）

```bash
pip install pyinstaller openpyxl pandas
pyinstaller Excel批量表单生成器.spec
```

或等价的一行命令（与 spec 中 `main.py` 一致）：

```bash
pyinstaller --onefile --windowed --name "Excel批量表单生成器" main.py
```

---

## 仓库结构（节选）

```
excel表格/
├── excel_workflow/          # 可视化工作流包（app、ui、nodes、core、session、ops…）
├── main.py                  # Tk 批量表单生成器入口
├── run_workflow.py          # 工作流启动脚本（便于双击 / PyInstaller）
├── requirements-workflow.txt
├── Excel工作流.spec         # PyInstaller：工作流
├── Excel批量表单生成器.spec # PyInstaller：Tk 工具
├── mapping_config.json      # Tk 工具配置（运行后生成，可不提交）
├── examples/                # 运行说明、模板 JSON、密钥示例、排障文档
│   ├── 运行说明.txt         # 团队快速开始与常见 Python 环境问题
│   └── templates/           # 内置 workflow 等参考
└── scripts/                 # 如 export_builtin_templates.py
```

---

## 更多说明与排障

- 安装失败、多版本 Python、`encodings` 报错、Anaconda 路径等问题：优先阅读 **`examples/运行说明.txt`**。
- 旧版映射迁移说明：`examples/templates/mapping_migrate.md`。
