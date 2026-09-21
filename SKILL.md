---
name: artifact-template-mygene-ppt
description: "Create or revise editable PowerPoint decks with the retained 迈景（Mygene）公司级PPT模板. Use when the user names Mygene/迈景, requests the company PPT template, or invokes $artifact-template-mygene-ppt. Preserves native masters, layouts, grouped artwork, linked title placeholders, directory pages, section dividers, typography, and brand colors."
---

# 迈景（Mygene）公司级PPT模板

使用保留的公司模板制作或修改可编辑 PowerPoint。必须保持母版联动、组合图形结构、品牌色和公司标识；不要把模板当成普通图片背景，也不要修改原始参考文件。

## 必读资源

1. 读取 artifact-template.json，并相对于本技能目录解析 reference 和 preview。
2. 加载 [@presentations](plugin://presentations@openai-primary-runtime)，遵循其创建、渲染和验证流程。
3. 始终读取 references/ooxml-safety.md。
4. 需要选择目录页、章节页或模板图形资源时，读取 references/template-map.md。

## 内容与设计原则

- 先确定叙事结构，再选模板视觉资源；图形必须与内容语义匹配。
- 允许新建可编辑的表格、流程图、关系图和数据图表，但要沿用公司蓝、青绿、红色强调色、微软雅黑字体和页眉标识。
- 没有比例或循环关系时不要使用环形图；没有真实定量数据时不要使用柱状图、折线图或饼图。
- 没有合适的模板图形时，使用简洁的自建版式，不为“套模板”而套模板。
- 普通内容页优先使用原模板的“内容”版式；标题必须写入真实标题占位符，以便修改母版标题样式后自动联动。
- 不要在标题占位符上覆盖一个外观相似的新文本框。
- 短篇演示稿通常不强制增加目录和章节页。用户明确要求，或内容具有三个及以上自然章节时，可以使用目录和章节页。
- 目录项目、章节页标题和实际章节顺序必须完全一致。

## 推荐工作流

1. 根据用户内容拟定逐页大纲，先控制信息密度和章节层级。
2. 对照模板图谱，选择真正合适的封面、目录、章节页和图形示例。
3. 以 assets/reference.pptx 的原生 OOXML 包为基础生成新文件，不改写参考文件。
4. 只修改必要的文本和页面级内容；不得整体重新序列化母版、版式、主题或复杂组合图形。
5. 封面、目录和章节页尽量复制原生页面后只替换文字。
6. 新建内容页时关联原始内容版式，并使用原生标题占位符。
7. 完成后运行结构验证脚本，再用 PowerPoint 原生导出逐页检查。

## 验收标准

- 输出文件中的 slideMasters、slideLayouts、theme 和模板已有 media 与参考文件字节一致。
- 封面、目录、章节页的 Logo、装饰线、章节数字和组合图形没有错位。
- 普通内容页标题为 type="title" 的占位符，且没有页面级坐标或列表样式覆盖母版。
- 表格、图表和流程图保持可编辑；数据图表不得虚构数值。
- PowerPoint 原生渲染中没有文本溢出、遮挡、残留占位符或异常换行。

## 验证命令

结构验证：

    pwsh.exe -NoLogo -NoProfile -File scripts/Test-MygenePptx.ps1 -PptxPath "输出文件.pptx" -ExpectedSlideCount 15

PowerPoint 原生导出：

    pwsh.exe -NoLogo -NoProfile -File scripts/Export-MygenePowerPointPreview.ps1 -PptxPath "输出文件.pptx" -OutputDirectory "预览目录"

结构检查不能替代视觉检查。复杂组合图形在第三方渲染器中可能显示异常，最终以 Microsoft PowerPoint 原生渲染为准。
