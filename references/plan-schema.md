# 计划字段与布局

`render_scut.mjs` 读取 UTF-8 JSON，画布为 1280 × 720（16:9）。每页使用一个布局；同一种布局可重复使用，页数由任务决定。

## 顶层字段

| 字段 | 类型 | 用途 |
|---|---|---|
| `title` | string | 汇报名称 |
| `reference` | string | 全局完整参考文献信息，写入每页讲稿 |
| `doi` | string | DOI 或来源链接，写入讲稿；无 DOI 时可省略 |
| `assets.directory` | string | 资产目录；相对路径以计划文件所在目录为基准 |
| `assets.logo` | string | PNG 校名/校徽文件名，默认 `scut-logo-blue.png` |
| `assets.cover` | string | JPEG 封面图片名，默认 `scut-campus.jpg` |
| `slides` | array | 按放映顺序排列的页面 |

生成器默认资产目录为计划文件同级的 `assets`。使用技能内素材时，将 `assets.directory` 设为该技能 `assets` 的绝对路径。

## 每页通用字段

| 字段 | 要求与用途 |
|---|---|
| `layout` | 必填，取下表中的布局名 |
| `title` | 必填，标题不可为空且全稿不能重复 |
| `source` | 必填，具体页码、章节、图表编号或可核验位置 |
| `subtitle` | 正文页可选；封面必填 |
| `citation` | 可选，页脚短引文；缺省使用 `source` |
| `notes` | 讲稿；说明核心论证和必要限定，生成器会追加来源 |

仅在原文存在时使用页码和 DOI。汇报者的推论应在页面或讲稿中标为解读，同时指向支撑材料。

## 布局字段

`items`、`leftItems`、`rightItems` 的元素均为 `{"label":"短标题","body":"说明文字"}`。`values` 为包含表头的二维字符串数组，行的列数一致，单元格不能为空。

| layout | 所需专用字段 | 当前适配的内容量 |
|---|---|---|
| `cover` | `subtitle`、`authorLine`、`journalLine` | 大标题、简短副标题、已知作者/汇报信息及期刊信息 |
| `intro` | `statement`、`items` | 左侧一句主张，右侧 3 项支持信息 |
| `comparison` | `values`、`takeaway` | 4 列表格，通常表头加 3 行；底部一句结论 |
| `decision` | `items`、`takeaway` | 3 组标签与说明，适合选型或条件判断 |
| `process` | `items`、`takeaway` | 4 个横向步骤，步骤编号自动生成 |
| `screening` | `leftTitle`、`rightTitle`、`leftItems`、`rightItems` | 左右两组内容，各 2 项 |
| `synthesis` | `values`、`takeaway` | 3 列表格，通常表头加 3 行 |
| `quality` | `values` | 3 列表格，通常表头加 4 行 |
| `contribution` | `items`、`takeaway` | 3 组标签与说明 |
| `summary` | `statement`、`items` | 一句主张及 3 项简洁结论 |

这些名称表示版式，内容主题可按论文自由组织。例如 `comparison` 可用于实验设置对比，`process` 可用于研究流程。所列数量是当前页面空间的适配范围；任何新增字段或更大容量，需同步扩展生成器并验证。

中文表格单元格通常以 15–25 字以内起步，结合实际列宽调整。长标题、长英文术语、手动换行都占用空间。字符串中的 `\n` 表示换行；来源和讲稿也应保持准确易读。

## 一个正文页示例

以下仅展示页面对象，引用位置需填写实际来源：

```json
{
  "layout": "decision",
  "title": "研究方案如何匹配问题",
  "subtitle": "按证据需求组织比较。",
  "items": [
    {"label": "研究问题", "body": "说明需要回答的具体问题。"},
    {"label": "证据需求", "body": "说明支持判断所需的材料。"},
    {"label": "分析路径", "body": "说明材料如何转化为结论。"}
  ],
  "takeaway": "方法选择应能支持预期研究结论。",
  "source": "实际原文章节与页码",
  "notes": "讲解各项选择之间的联系，并交代原文给出的适用条件。"
}
```

## 生成接口

`render_scut.mjs --plan PLAN.json --output CANDIDATE.pptx [--previews DIR]`

`--plan`、`--output` 必填；预览默认写入输出文件同级的 `previews`。生成器同时输出逐页 PNG、布局 JSON 和文字布局审计信息。使用捆绑运行时，脚本同目录的 `style_native_tables.py` 负责原生表格样式。

`finalize_scut.mjs --candidate CANDIDATE.pptx --output FINAL.pptx --workspace WORKSPACE --skill-dir PRESENTATIONS_SKILL_DIR --plan PLAN.json`

最终校验的 `--skill-dir` 是 presentations 技能的绝对目录。`--plan` 用于核对总页数和计划要求的原生表格。校验之后对 FINAL.pptx 调用：

`render_wps.py FINAL.pptx WPS_PREVIEW_DIR`

WPS 导出的逐页图片用于实际显示效果检查。最终交付文件有修改时，检查与预览应对应修改后的版本。
