---
name: scut-paper-ppt
description: "制作可编辑的华工学术论文汇报与开题答辩 PPT，支持华工蓝通用模板及模板一至四，保留原版配图与构图，按内容关系排版并提供本机预览。仅在用户选择华工、华南理工/SCUT 风格、这些模板或此技能时使用。"
---

# 华工蓝论文PPT

把论文、开题报告或研究计划组织为适合讲述的汇报，沿用所选模板的配色、校园配图、校名校徽和特色构图。页数、章节与页面结构由材料内容和汇报目标决定。

## 模板选择

原有华工蓝通用模板位于 `assets/scut-blue-template.pptx`。已收录的开题答辩模板库包含模板一 13 页、模板二 15 页、模板三 16 页、模板四 15 页，文件与预览均存放在 `assets/defense-templates/`。

用户指定四套模板、提到这些开题模板，或需要浏览模板时，读取 [开题答辩模板库](references/defense-template-library.md)，按其索引选页与定位文件。使用用户指定的模板；一般华工蓝论文汇报沿用原有默认模板。 四套开题模板按该参考中的副本编辑流程生成，并按各自实际尺寸校验；下方专用模板脚本沿用于原有华工蓝模板。

开题任务以开题报告、研究计划或用户说明为内容依据，按需要补充相关论文。明确区分研究基础、拟实施的方法和预期结果；每页出处与讲稿按本次交付要求安排。

## 内容与模板

论文汇报先阅读全文，包括方法、结果、限制和图表说明，再确定每页论点、证据与出处。正式论文汇报在正文中展开研究问题、关键方法、依据或案例、结果或论证、贡献与局限；讲稿补充推导、解释和衔接。根据原文信息量安排页数，保留有证据支持的实质细节。准确保留数值、单位、条件和不确定性，并区分作者结论与汇报者解读。原文缺失的部分应明确可核验范围。

默认模板为 `assets/scut-blue-template.pptx`。先查看原模板全部页面，识别其色彩、字体层级、校园配图与特色裁切，再按每页的阅读任务组织内容。比较关系用表格或同维度对照，步骤用流程，条件选择用分支，编码与提取用信息表，质量评价用检查清单。布局选择与整稿检查参见 [内容与布局](references/content-layout.md)。

原页构图适合当前内容时，沿用其图形、图像、裁切、遮罩、渐变和母版，并局部调整文字。正文需要新的信息结构时，使用 `presentations:Presentations` 的原生文字、表格、形状和连接线重新排版，保持内容可编辑；按页面作用安排校园图文和原模板特色页。相同的信息关系可使用一致表达，整稿应检查连续页面的阅读动作、信息重心与内容密度。

关键方法、结果和依据须在正文中可见；讲稿承接推导、细节和衔接。正文参考 23–28 px（约 17–21 pt），结合页面尺度与投影可读性调整。容量不足时优先整理层级、调整结构或拆页。每页保留来源，必要解释写入讲稿。

## 默认流程

1. 核对论文或开题材料、所选模板和汇报要求，读取可用的 `presentations:Presentations` 技能；调用 `load_workspace_dependencies` 定位捆绑 Node.js、Python 及库。WPS 导出需本机 WPS 演示与可用的 `win32com.client`。
2. 阅读材料与所选模板，形成逐页计划：论点、证据、阅读任务、信息关系、拟用布局、来源与讲稿。先确定内容结构，再选择适合的原页或独立设计正文。
3. 按各页需要生成可编辑内容；整页沿用模板和原生正文可在同一汇报中组合。输出使用新文件，源模板不可覆盖。已有 `scripts/assemble_layout.py` 时，可先查看其用法，再用于合并模板页与独立正文；也可使用当前演示工具支持的页面组装方式。
4. 按所用生成方式完成结构校验，并核对页序、文字、原生对象、来源和讲稿。适合模板计划的候选文件可用 `finalize_scut.mjs` 生成最终文件；`--skill-dir` 指向 **presentations 技能目录**，`--plan` 使用对应模板计划。
5. 用 WPS 导出最终 PPTX 的逐页 PNG，逐页检查文字清晰度、换行、遮挡、图表关系、配图裁切和来源，再用整稿总览检查连续页面的重心、密度与节奏。沿用原页的页面同时与原模板对照。程序校验通过不等于视觉通过；修改后重新导出受影响页面并复查整稿。
6. 交付可编辑 PPTX、实际预览及必要的计划/讲稿。若 WPS 不可用，继续完成可行的生成与校验，准确说明尚未完成的视觉检查。

## 整页沿用模板的工具

`render_template.mjs` 及 `preserve_template` 工具适用于需要整体保留原页构图的页面。使用前读取 [模板计划字段](references/template-plan-schema.md)，将原页每个非空文字形状的原始 `name` 与新内容精确对应，并记录位置与尺寸。`examples/template-review.json` 是字段示例，内容按实际论文填写。

PowerShell 调用形式如下，变量先设为本次使用的绝对路径：

```powershell
& $node "$scutSkill/scripts/render_template.mjs" --plan $plan --output $candidate
& $node "$scutSkill/scripts/finalize_scut.mjs" --candidate $candidate --output $final --workspace $workspace --skill-dir $presentationsSkill --plan $plan
& $python "$scutSkill/scripts/render_wps.py" $final $wpsPreview
```

同一原页需要出现多次时，先用 `prepare_template.py` 创建模板副本及独立页码的计划；`--pptx-skill` 指向已安装 `pptx` 技能目录。使用新的准备目录执行：

```powershell
& $python "$scutSkill/scripts/prepare_template.py" --plan $plan --output-dir $preparedDir --pptx-skill $pptxSkill
```

后续模板生成和校验使用 `$preparedDir/plan.json`；模板副本也保存在该目录。提交给模板渲染器的计划中，每个准备后的原页只出现一次。仅覆盖本次自己的草稿时可传 `--replace-draft`。