"""
线性流程各功能的说明书（短提示 + 长文 HTML）。
全文说明书由 all_manual_html() 组装；各功能的 tooltip / 单项「?」仍读 DOCS。
"""

from __future__ import annotations

from typing import Dict, Tuple

DOCS: Dict[str, Tuple[str, str]] = {
    "overview_flow": (
        "导入 → 操作 → 导出 三步",
        "<p>本程序按<strong>固定顺序</strong>完成三件事：<strong>导入</strong>把表格读入内存；<strong>操作</strong>在左侧依次点功能、填参数、点「应用」，步骤会出现在「步骤流程」里；<strong>导出</strong>把结果写到您选的文件夹。</p>"
        "<p>列表里<strong>第一个文件</strong>始终是<strong>主表</strong>：预览、列名联想、筛选条件默认值等都来自主表。多文件时可用于批量对每张表执行同一套步骤。</p>",
    ),
    "pills_phases": (
        "顶部三步仅作进度展示",
        "<p>界面上方的「导入 / 操作 / 导出」用于<strong>提示当前阶段</strong>，不能点击跳转。必须按顺序完成：先导入成功进入操作，再点底栏「导出文件」进入导出。</p>",
    ),
    "import_drop": (
        "拖放 xlsx / xls / csv",
        "<p>在导入页虚线框内<strong>拖入</strong>一个或多个表格文件即可开始载入。支持扩展名：xlsx、xls、csv。</p>",
    ),
    "import_select_files": (
        "多选本地表格",
        "<p>点「选择文件」可多选。确认后<strong>第一个被选中的文件</strong>会成为主表；顺序与资源管理器中的选择顺序有关，若主表不对请重新选择或拖放调整顺序（需重新导入）。</p>",
    ),
    "import_folder": (
        "扫描文件夹",
        "<p>点「选择文件夹」会扫描该目录下的表格文件，在弹出列表中<strong>勾选</strong>要载入的项再确定。适合一次加入多个分散文件。</p>",
    ),
    "import_mapping": (
        "旧版 mapping_config.json",
        "<p>若您曾使用旧版「映射配置文件」，可在此读取。导入后请先按正常流程<strong>再导入数据表</strong>进入操作页；进入后「填充模板」等参数可按映射<strong>预填</strong>，请核对路径与列名。</p>",
    ),
    "append_import": (
        "追加导入",
        "<p>在操作页已载入文件卡片中点「追加导入」，可在<strong>不改动已应用步骤</strong>的前提下向列表末尾增加文件。若追加后需把新文件当主表，请先删除原主表或返回导入重新选顺序。</p>",
    ),
    "remove_source_files": (
        "删除所选已载入文件",
        "<p>在文件列表中<strong>按住 Ctrl 多选</strong>后点「删除所选」，可从会话中移除（<strong>不会删除磁盘上的文件</strong>）。</p>"
        "<p>若移除的是<strong>当前主表</strong>，程序会用新的第一个文件作主表并尽量<strong>重放</strong>已有步骤；若条件在新主表上不成立可能报错。若删到一张不剩，可提示返回导入界面。</p>",
    ),
    "back_to_import_step": (
        "返回导入",
        "<p>操作页或导出页的「返回导入」会回到第一步。若有已载入文件或已执行步骤，会先<strong>二次确认</strong>；确认后<strong>清空</strong>当前会话（文件列表与步骤），与重新开始一次导入等价。</p>",
    ),
    "batch_all_sources": (
        "对全部表格执行当前步骤",
        "<p>至少载入<strong>两个</strong>表格且<strong>已有至少一步</strong>已应用步骤时可用。程序会对<strong>每一张</strong>已载入表依次重放整条步骤链。</p>"
        "<p>因此每张表都必须具备步骤里用到的<strong>全部列名</strong>，且「筛选行」里的条件式在每张表上都能计算通过；否则会报错并<strong>中止批量</strong>，错误信息中会写明哪张表、哪一步、何种原因。</p>",
    ),
    "import_workflow_json": (
        "导入 workflow.json v2",
        "<p>读取本程序保存的 <code>workflow.json</code>（版本 2），用其中的步骤列表<strong>替换</strong>当前步骤流程并从主表<strong>重放</strong>。用于复用他人或自己保存的流程模板。载入失败时请检查文件编码与步骤参数是否与当前数据匹配。</p>",
    ),
    "bottom_save": (
        "保存流程",
        "<p>将当前「步骤流程」中的全部已应用步骤导出为 <code>workflow.json</code>，便于下次「导入 workflow」或备份。不含原始 Excel 文件本身，仅含操作类型与参数。</p>",
    ),
    "bottom_export": (
        "导出文件",
        "<p>进入<strong>导出</strong>阶段：勾选要导出的项（主表、缓存生成文件等），选输出目录与格式后执行写出或复制。</p>",
    ),
    "bottom_batch_export": (
        "批量导出",
        "<p>与「导出文件」相同，但会预先<strong>勾选全部</strong>可导出项，适合一次全部落盘。</p>",
    ),
    "export_selected_one": (
        "导出选中项",
        "<p>在导出列表中<strong>单击选中一行</strong>（与左侧勾选无关），再点「导出选中项」。主表按当前内存结果按所选格式写出；其它已在磁盘上的文件一般按<strong>复制</strong>处理。</p>",
    ),
    "export_batch_all": (
        "导出全部已勾选",
        "<p>导出所有在列表中<strong>已打勾</strong>的项到目标目录。若文件名冲突会自动加序号避免覆盖。</p>",
    ),
    "format_std": (
        "格式标准化",
        "<p>对指定列做文本层面的整理（如首尾空白）。在表单中用<strong>多选列表</strong>勾选列；<strong>不选</strong>表示对<strong>全部列</strong>处理。</p>",
    ),
    "dedup": (
        "去重",
        "<p>按指定列组合或整表行去重。用多选列表勾选参与判断的列；<strong>不选</strong>表示按<strong>整行完全相同</strong>去重。可勾选是否保留首次出现行。</p>",
    ),
    "fill_empty": (
        "空值填充",
        "<p>将指定列中的空单元格填成同一固定值。先填「填充值」，再用多选列表勾选列；<strong>不选列</strong>表示对<strong>全部列</strong>填充。</p>",
    ),
    "filter_rows": (
        "筛选行：按条件保留满足条件的行",
        "<p><b>作用</b>：根据你写的<strong>一条条件式</strong>，只保留当前表里「条件成立」的那些行，相当于对内存中的表做一次筛选（不修改磁盘上的原文件）。</p>"
        "<p><b>怎么写</b>：条件式里可以直接用<strong>列名</strong>当作变量。程序对<strong>每一行</strong>代入各列的值判断，结果为真的行会留下。</p>"
        "<p><b>列名</b>：纯英文数字列名通常可直接写。列名含中文、空格或与关键字冲突时，用<strong>英文反引号</strong>把列名包起来，例如 <code>`销售金额` &gt; 1000</code>。</p>"
        "<p><b>运算与逻辑</b>：比较用 <code>&gt;</code> <code>&lt;</code> <code>&gt;=</code> <code>&lt;=</code> <code>==</code> <code>!=</code>；多条件用 <code>&amp;</code>（并且）与 <code>|</code>（或者）；文字用英文双引号。空值可用 <code>`列名`.isna()</code> / <code>.notna()</code>。</p>"
        "<p><b>技巧</b>：输入框有列名联想；「从主表插入列名」可插入当前主表列名，减少手误。</p>",
    ),
    "sort": (
        "排序",
        "<p>按<strong>一个</strong>列升序或降序重排全表。请从下拉或联想中选择列名。</p>",
    ),
    "column_pick": (
        "选列",
        "<p>只保留勾选的列，其余列从当前表中移除（内存中）。请至少勾选一列再应用。</p>",
    ),
    "replace": (
        "查找替换",
        "<p>在<strong>单列</strong>内按字面文本查找并替换（非正则）。先选列，再填查找与替换为的内容。</p>",
    ),
    "rename_cols": (
        "列重命名",
        "<p>用 JSON 对象描述「旧列名 → 新列名」。也可用「从主表列构造」一行行加入。键名须与当前表头一致。</p>",
    ),
    "merge_vertical": (
        "纵向合并",
        "<p>把<strong>第二个</strong>表格的所有行接到当前主表下方。第二个表可从已载入列表选，也可「其他文件」浏览。列名对齐方式与底层引擎一致，列不全时可能出现空列。</p>",
    ),
    "template_fill": (
        "填充模板",
        "<p>按主表<strong>每一行</strong>生成一份填好的模板表文件。需指定模板路径、输出在缓存目录内、用于文件命名的列、工作表名及列到单元格的映射（可为 JSON 数组）。映射写法见节点说明或示例 JSON。</p>",
    ),
    "write_xlsx": (
        "写出 xlsx",
        "<p>把当前内存表写出到<strong>会话缓存目录</strong>下的 xlsx 文件，便于后续打包或导出阶段勾出。文件名可留空由程序自动生成。</p>",
    ),
    "zip_pack": (
        "打包 ZIP",
        "<p>把当前缓存里已生成的文件打成一个压缩包，同样写在缓存目录。需先有「写出」等步骤产生的文件。</p>",
    ),
    "timeline_delete": (
        "删除步骤",
        "<p>在步骤流程中点某一行的「删除」，会<strong>去掉该步</strong>并用剩余步骤从主表<strong>从头重放</strong>。缓存目录与中间文件会重建，请确认无未导出重要结果再删。</p>",
    ),
    "param_apply": (
        "应用",
        "<p>用当前表单参数<strong>执行一次</strong>该功能，并把结果记为步骤流程中的一步。若执行报错，该步不会加入历史。</p>",
    ),
    "param_cancel": (
        "取消草稿",
        "<p>放弃左侧当前选中的功能表单，不清空已应用步骤。标题恢复为「请从左侧选择一项操作」。</p>",
    ),
    "preview_data": (
        "数据预览",
        "<p>勾选「数据预览」标题左侧的展开框后，可查看当前内存表前约 100 行文本预览，便于核对筛选、排序等结果。</p>",
    ),
    "export_dir": (
        "输出目录",
        "<p>导出前请选择目标文件夹。若目录不存在，程序会尝试创建。</p>",
    ),
    "export_format": (
        "输出格式",
        "<p><strong>主表</strong>按所选格式从<strong>内存</strong>写出：<b>xlsx</b> 为表格文件；<b>csv</b> 为带 BOM 的 UTF-8 文本，便于 Excel 打开；<b>zip</b> 为内含一份 csv 的压缩包。</p>"
        "<p><strong>其它项</strong>一般为磁盘上已有路径，多按<strong>复制</strong>到目标目录处理。</p>",
    ),
    "export_pdf_stub": (
        "PDF 占位",
        "<p>格式列表中的「pdf」为预留项，当前版本<strong>尚未实现</strong>写出。请选择 xlsx、csv 或 zip。</p>",
    ),
    "export_back": (
        "返回操作",
        "<p>从导出阶段回到<strong>操作</strong>阶段，不自动清空已勾选导出项；主表与步骤仍保留。</p>",
    ),
}


def tooltip(key: str) -> str:
    t = DOCS.get(key)
    return t[0] if t else ""


def body_html(key: str) -> str:
    t = DOCS.get(key)
    return t[1] if t else "<p>暂无说明。</p>"


def import_section_html() -> str:
    """导入页「本页功能说明」对话框。"""
    css = "font-family:Microsoft YaHei UI,sans-serif;font-size:13px;line-height:1.65;"
    parts = [
        f"<html><body style='{css}'>",
        "<h2>导入阶段</h2>",
        "<p>完成本页操作后，会自动进入「操作」页。主表为列表中的<strong>第一个</strong>文件。</p>",
        "<h3>拖入文件</h3>",
        body_html("import_drop"),
        "<h3>选择文件与文件夹</h3>",
        body_html("import_select_files"),
        body_html("import_folder"),
        "<h3>旧版映射配置</h3>",
        body_html("import_mapping"),
        "</body></html>",
    ]
    return "\n".join(parts)


def export_section_html() -> str:
    """导出页「本页说明」对话框。"""
    css = "font-family:Microsoft YaHei UI,sans-serif;font-size:13px;line-height:1.65;"
    parts = [
        f"<html><body style='{css}'>",
        "<h2>导出阶段</h2>",
        "<p>勾选要导出的条目，选择目录与格式后执行。主表始终反映<strong>当前内存中</strong>经过各步骤后的结果。</p>",
        "<h3>目录与格式</h3>",
        body_html("export_dir"),
        body_html("export_format"),
        body_html("export_pdf_stub"),
        "<h3>导出方式</h3>",
        body_html("export_selected_one"),
        body_html("export_batch_all"),
        "<h3>返回操作</h3>",
        body_html("export_back"),
        "</body></html>",
    ]
    return "\n".join(parts)


def all_manual_html() -> str:
    """工具栏《功能说明书》全文：章节清晰、中文为主。"""
    css = (
        "font-family:Microsoft YaHei UI,Segoe UI,sans-serif;font-size:13px;"
        "line-height:1.65;color:#1e293b;"
    )
    h2 = "margin-top:22px;margin-bottom:8px;font-size:15px;color:#0f172a;border-left:4px solid #3f51b5;padding-left:10px;"
    h3 = "margin-top:14px;margin-bottom:6px;font-size:13px;color:#334155;font-weight:600;"
    note = "margin:8px 0;padding:10px 12px;background:#f1f5f9;border-radius:8px;font-size:12px;color:#475569;"

    def p(html: str) -> str:
        return f"<div style='margin:6px 0 10px 0;'>{html}</div>"

    parts: list[str] = [
        f"<html><head><meta charset='utf-8'/></head><body style='{css}'>",
        "<h1 style='font-size:18px;color:#0f172a;margin-top:0;'>Excel 表单工作流 · 使用说明书</h1>",
        f"<p style='{note}'>说明对象为主窗口内的<strong>线性流程</strong>（导入 → 操作 → 导出）。"
        "与旧版「节点连线图」若并存于其它入口，以本窗口为准。</p>",

        f"<h2 style='{h2}'>一、整体概念</h2>",
        p(body_html("overview_flow")),
        p(body_html("pills_phases")),

        f"<h2 style='{h2}'>二、导入阶段</h2>",
        "<p>首次使用或点击「返回导入」后进入。完成载入后自动进入操作页。</p>",
        f"<h3 style='{h3}'>载入方式</h3>",
        p(body_html("import_drop")),
        p(body_html("import_select_files")),
        p(body_html("import_folder")),
        f"<h3 style='{h3}'>旧版映射文件</h3>",
        p(body_html("import_mapping")),

        f"<h2 style='{h2}'>三、操作阶段 · 文件与流程</h2>",
        f"<h3 style='{h3}'>已载入文件列表</h3>",
        p(body_html("append_import")),
        p(body_html("remove_source_files")),
        p(body_html("back_to_import_step")),
        f"<h3 style='{h3}'>步骤流程与时间线</h3>",
        p(body_html("timeline_delete")),
        f"<h3 style='{h3}'>批量与流程文件</h3>",
        p(body_html("batch_all_sources")),
        p(body_html("import_workflow_json")),

        f"<h2 style='{h2}'>四、操作阶段 · 数据处理</h2>",
        "<p>左侧点选功能后，在右侧填写参数，点<strong>应用</strong>写入一步。下列各节对应左侧面板中的名称。</p>",
        f"<h3 style='{h3}'>格式标准化</h3>",
        p(body_html("format_std")),
        f"<h3 style='{h3}'>去重</h3>",
        p(body_html("dedup")),
        f"<h3 style='{h3}'>空值填充</h3>",
        p(body_html("fill_empty")),
        f"<h3 style='{h3}'>筛选行</h3>",
        p(body_html("filter_rows")),
        f"<h3 style='{h3}'>排序</h3>",
        p(body_html("sort")),
        f"<h3 style='{h3}'>选列</h3>",
        p(body_html("column_pick")),
        f"<h3 style='{h3}'>查找替换</h3>",
        p(body_html("replace")),
        f"<h3 style='{h3}'>列重命名</h3>",
        p(body_html("rename_cols")),

        f"<h2 style='{h2}'>五、操作阶段 · 结构与产出</h2>",
        f"<h3 style='{h3}'>纵向合并</h3>",
        p(body_html("merge_vertical")),
        f"<h3 style='{h3}'>填充模板</h3>",
        p(body_html("template_fill")),
        f"<h3 style='{h3}'>写出到缓存</h3>",
        p(body_html("write_xlsx")),
        f"<h3 style='{h3}'>打包 ZIP</h3>",
        p(body_html("zip_pack")),

        f"<h2 style='{h2}'>六、应用、取消与预览</h2>",
        p(body_html("param_apply")),
        p(body_html("param_cancel")),
        p(body_html("preview_data")),

        f"<h2 style='{h2}'>七、底栏按钮</h2>",
        p(body_html("bottom_save")),
        p(body_html("bottom_export")),
        p(body_html("bottom_batch_export")),

        f"<h2 style='{h2}'>八、导出阶段</h2>",
        p(body_html("export_dir")),
        p(body_html("export_format")),
        p(body_html("export_pdf_stub")),
        p(body_html("export_selected_one")),
        p(body_html("export_batch_all")),
        p(body_html("export_back")),

        "<p style='margin-top:24px;color:#94a3b8;font-size:12px;'>— 文档随软件更新，若与界面不一致以当前界面为准 —</p>",
        "</body></html>",
    ]
    return "\n".join(parts)
