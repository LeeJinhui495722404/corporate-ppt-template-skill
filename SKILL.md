---
name: artifact-template-mygene-ppt
description: "Create or revise editable PowerPoint decks with the retained 迈景（Mygene）公司级PPT模板. Use when the user names Mygene/迈景, requests the company PPT template, or invokes $artifact-template-mygene-ppt. Preserves native masters, layouts, grouped artwork, linked title placeholders, directory pages, section dividers, typography, and brand colors; selects integrated diagrams, charts, timelines and icons by content semantics."
---

# 迈景（Mygene）公司级PPT模板

使用保留的公司模板制作或修改可编辑 PowerPoint。保持母版联动、组合结构、品牌色和公司标识；不要把模板当成图片背景。常规制作不得修改参考文件；只有用户明确要求更新模板时，才编辑临时副本、验证后原子更新同一个技能。

## 必读资源

1. 读取 `artifact-template.json`，相对于本技能目录解析 reference 和 preview。
2. 加载 [@presentations](plugin://presentations@openai-primary-runtime)，遵循其创建、渲染和验证流程。
3. 始终读取 `references/ooxml-safety.md`。
4. 选择封面、目录、章节或原有图形时，读取 `references/template-map.md`。
5. 选择展示方式、图表或图标时，完整读取 `references/visual-selection.md`，再按需要检索 `references/visual-resources.json` 并查看候选预览。不要只凭资源名称直接套用。

## 单一参考文件与资源位置

唯一参考文件仍是 `assets/reference.pptx`，共109页：

- 第1–56页：原公司模板，原顺序与原生结构保持不变。
- 第57–104页：48种新增展示方式，编号R001–R048。
- 第105–109页：5张新增图标页，432个对象（336个原生形状、96个图片对象，含重复和品牌素材）。
- 第56页仍是结束页，不能用最后一页代替。新增资源的来源、准确页码、图标形状ID、适用条件见索引；预览位于 `assets/resource-previews/`。
- 没有独立展示方式库/图标库PPT作为第二参考；新增资源只能补充正文表达，不能取代原封面、目录、章节、母版规则。
- 109页是参考库规模，不是成品页数；选取需要的页面，不把素材页自动加入汇报。

## 内容与设计原则

- 先确定叙事结构和信息关系，再选视觉资源。根据节点数量、读序、正文密度与横纵方向选择，不为填满图形虚构内容。
- 允许自建可编辑表格、流程图、关系图和原生数据图表；沿用公司蓝、青绿、红色强调色、微软雅黑和页眉标识。
- 没有比例或反馈关系不用环形结构；没有真实定量数据不用柱状、折线或饼图。图标不能替代事实和检测证据。
- 模板不合适时创建简洁的新可编辑图形，不强求使用素材。品牌图标默认排除；同页图标风格、大小和视觉重量一致。
- 普通内容页关联原模板“内容”版式，标题写入真实标题占位符。不得叠加仿制标题框，也不得页面级覆盖坐标、字体、字号或颜色而破坏母版联动。
- 短篇演示不强制目录和章节；用户明确要求或有三个及以上自然章节时使用。目录条目、章节标题、实际顺序必须一致。
- 示例业务、年份、图例、数字和引用不是用户事实。引用后全部替换/删除，并同步图表的工作簿、缓存、公式范围和注释。
- 本技能优化展示形式、图表和图标选择，不包含医学突变位点筛选规则；用户若要求后者，先确认具体含义和证据要求。

## 推荐工作流

1. 根据内容拟定逐页大纲，控制信息密度和章节层级。
2. 对照原模板图谱选择结构页；按视觉选择规则检索2–3个正文候选，查看预览后选用。
3. 以参考文件的原生OOXML包为基础输出新文件，不改写参考文件。可用 `scripts/Export-MygeneReferencePages.ps1` 选择唯一页码生成草稿；重复页另行克隆部件和关系ID。
4. 只修改必要的文本和页面级内容，不整体重新序列化母版、版式、主题或复杂组合。
5. 封面、目录、章节页复制原生页面，只替换文字；删减目录时删除整个多余组合，章节编号与正文保持一致。
6. 复制新增展示页的原生结构后替换示例内容；图标按索引的 `NativeShapeId` 复制并保留依赖。普通标题不添加局部样式。
7. 新建内容页关联原始内容版式；不合适的模板图形可重建正文可编辑对象，但保留公司视觉系统。
8. 运行结构、图表数据及Presentations技能要求的检查，再用PowerPoint原生导出逐页复核并修复。

## 验收标准

- 输出中的slideMasters、slideLayouts、theme和模板已有media与参考文件字节一致。
- 封面、目录、章节的Logo、装饰条、章节数字和组合图形无错位；原生变换保留。
- 普通标题为 `p:ph type="title"`，`p:spPr` 与 `a:lstStyle` 为空，段落/文字运行不覆盖母版字体、字号、颜色等样式。
- 表格、数据图表和流程图保持可编辑；图片图标不宣称路径可编辑。数值、单位、比例、时间跨度与正文一致。
- 无不相关图标、示例数字或引用残留、文本溢出、遮挡、异常换行；结构检查不能替代视觉检查。
- 原参考第46/48/49页的历史图表示例有缺省类别或稀疏缓存，严格图表检查可能不支持。参考库保留原生结构；制作成品时用真实数据重建或完整更新并验证，不虚构零值、不跳过成品图表检查。新增第99–104页图表已通过结构和嵌入工作簿校验，仍须在替换数据后重新验证。

## 工具调用

以下命令在技能目录中执行；也可以使用脚本绝对路径。中文文本显式UTF-8读取。

```powershell
pwsh.exe -NoLogo -NoProfile -File scripts/Get-MygeneVisualResource.ps1 -Kind method -Query 时间轴
pwsh.exe -NoLogo -NoProfile -File scripts/Get-MygeneVisualResource.ps1 -Kind icon -Query 日历
& ./scripts/Export-MygeneReferencePages.ps1 -Page 2,4,5,63,56 -OutputPath 'D:\工作目录\草稿.pptx'
pwsh.exe -NoLogo -NoProfile -File scripts/Test-MygenePptx.ps1 -PptxPath '输出文件.pptx' -ExpectedSlideCount 15
pwsh.exe -NoLogo -NoProfile -File scripts/Export-MygenePowerPointPreview.ps1 -PptxPath '输出文件.pptx' -OutputDirectory '预览目录'
```

选择脚本生成的是未填充草稿，不是成品。标题检查遍历活动页关系，支持新增资源的非标准slide部件名。复杂组合在第三方渲染器中可能不同，视觉效果最终以Microsoft PowerPoint原生渲染为准。
