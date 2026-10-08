# 模板计划字段

模板计划把内容准确映射到原模板页面与文字形状。生成器保留所选原页的图形、图像、裁切、遮罩、渐变、装饰及母版，按计划修改指定文字。文件使用 UTF-8 JSON。

## 顶层字段

| 字段 | 类型 | 含义 |
|---|---|---|
| `title` | string | 汇报标题，写入文件信息 |
| `reference` | string | 完整参考文献，写入各页讲稿 |
| `doi` | string，可选 | DOI 或来源链接 |
| `template` | string | 源 PPTX 路径；相对路径以计划文件所在目录为基准 |
| `slides` | array | 按最终放映顺序排列的页面计划 |

默认模板为技能目录下的 `assets/scut-blue-template.pptx`。位于 `examples/` 的计划使用 `"template": "../assets/scut-blue-template.pptx"`。其他位置的计划应调整路径或使用绝对路径。 四套开题模板按 [模板库](defense-template-library.md) 的副本编辑流程使用；目录中的 `page` 是放映页序。本文字段和下方专用脚本用于原有华工蓝流程，其他模板须先验证页部件映射、对象定位及画布兼容性。

## 每页字段

| 字段 | 类型 | 含义 |
|---|---|---|
| `templateSlide` | integer | 选取的源模板页码，从 1 开始 |
| `title` | string | 本页内容标题，用于计划和校验；页面上的标题仍须通过 `edits` 指定 |
| `source` | string | 原文页码、章节、图表编号等具体位置 |
| `notes` | string | 讲稿、必要限定及完整原文出处 |
| `edits` | array | 所选原页的文字形状修改清单 |

`templateSlide` 始终是源模板中的页码，不是输出中的序号。编写内容计划时可重复选择同一源页。重复使用版式的计划先交给 `prepare_template.py`，由其调用 `pptx` 技能的页面复制脚本，生成模板副本和独立页码计划；渲染器读取准备后的计划。源模板保持原样。

## 每个文字修改对象

| 字段 | 类型 | 含义 |
|---|---|---|
| `name` | string | 原始文字 shape 的精确名称，作为定位键 |
| `text` | string | 对应新文本；用 `\n` 表示换行 |
| `fontSize` | number | 字号，单位 px |
| `bold` | boolean | 是否加粗 |
| `color` | string | 文字颜色，使用脚本接受的色值 |
| `position` | object，可选 | 局部调整文本框，含 `left`、`top`、`width`、`height`，单位均为 px |
| `height` | number，可选 | 仅调整文本框高度，单位 px |
| `align` | string，可选 | 水平对齐，使用生成器支持的值 |
| `vertical` | string，可选 | 垂直对齐，使用生成器支持的值 |

从原页的形状清单读取名称及几何信息。每个原本非空的文字 shape 都须有且仅有一项修改，名称必须精确匹配；不要把显示文本当作形状名称。文字形状既包括标题和正文，也包括标签、页码或其他非空文字。

保留原框位置和尺寸作为起点。确需适配内容时，只在计划中调整相关文本框。使用 `position.height` 或 `height` 表达高度调整，避免同时提供相互冲突的值。逐页检查修改后的文字与校园图片、遮罩及装饰之间的关系。

## 内容编排

根据论文和汇报目的选择原页，页面顺序和数量服从讲述逻辑。每页突出明确论点，将推导、限定条件和完整来源写入 `notes`。资料不足时注明范围；汇报者解读应有明确标记和支撑出处。

内容超过所选页容量时，先改写、拆页或改选原页。新增研究图表应服务于内容，并与原模板视觉协调。完整字段示例见 [template-review.json](../examples/template-review.json)，其中内容及形状映射必须按本次原文与原页核验。

## 生成与检查

```text
render_template.mjs --plan PLAN.json --output CANDIDATE.pptx
finalize_scut.mjs --candidate CANDIDATE.pptx --output FINAL.pptx --workspace WORKSPACE --skill-dir PRESENTATIONS_SKILL_DIR --plan PLAN.json
render_wps.py FINAL.pptx WPS_PREVIEW_DIR
```

生成器可选 `--replace-draft` 仅用于覆盖本次自己的草稿。源模板不可覆盖，最终文件使用新的输出路径。`--skill-dir` 是 presentations 技能的绝对目录。

完成后核对输出页数与计划一致、所选原页及文字映射准确、讲稿出处完整。再对 WPS 导出的最终页面逐页目视，并与对应原页对照。文件校验和视觉检查应针对同一个最终版本。
