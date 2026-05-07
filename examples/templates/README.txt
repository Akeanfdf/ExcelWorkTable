启动页使用的内置模板文件：
  sales_report.json   — Excel 读取 → 格式标准化 → 模板填充
  merge_folder.json   — 两个 Excel 读取 → 纵向合并 → 写出 Excel

修改了 template_presets 中的逻辑后，可在仓库根目录执行以重新导出：
  python scripts/export_builtin_templates.py

若缺少上述 JSON，程序会按 excel_workflow/ui/template_presets.py 中的代码现场搭图。
