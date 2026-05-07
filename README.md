# Excel 批量表单生成器

## 功能概述
将「数据源 Excel」每一行的数据，自动填入「模板 Excel」的指定单元格，批量生成独立的 Excel 表单文件。

---

## 快速开始

### 方案 A：直接运行 Python（开发用）
```bash
pip install openpyxl pandas
python main.py
```

### 方案 B：打包为 .exe（发给他人使用）
```bash
# 安装打包工具
pip install pyinstaller openpyxl pandas

# 一键打包（Windows）
pyinstaller --onefile --windowed --name "Excel批量表单生成器" main.py

# 生成文件在 dist/ 目录下
```

---

## 界面说明

| 页面 | 功能 |
|------|------|
| 📂 文件配置 | 选择数据源、模板、输出目录，设置命名列 |
| 🔗 字段映射 | 配置「数据列 → 模板单元格」对应关系 |
| ▶  开始生成 | 查看配置摘要、实时进度、运行日志 |
| 📋 使用说明 | 操作指引 |

---

## 文件结构
```
excel_generator/
├── main.py          ← 主程序
├── build_exe.py     ← 打包指南
├── mapping_config.json  ← 自动保存的配置（运行后生成）
└── README.md
```
