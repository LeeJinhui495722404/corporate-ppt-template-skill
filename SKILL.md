---
name: artifact-template-mygene-ppt
description: "Create or iteratively revise editable 迈景（Mygene）company PowerPoint decks. Use when the user requests the Mygene company template, company PPT production, or $artifact-template-mygene-ppt. Combines retained native template and visual resources, PPT-Master collaboration when available, screenshot teaching layouts, and source-preserving revision and validation."
---

# 迈景（Mygene）公司级PPT制作

使用保留的公司模板制作或修改可编辑 PowerPoint。保持母版联动、组合结构、品牌色和公司标识；不要把模板当成图片背景。常规制作不得修改参考文件；只有用户明确要求更新模板时，才编辑临时副本、验证后原子更新同一个技能。

## 综合能力与职责

- **公司模板层**：本仓库提供原生母版、版式、视觉资源、字体与品牌规则，决定成稿外观及结构保护要求。
- **内容与视觉制作层**：吸收 PPT-Master 的内容梳理、叙事设计、图解制作、备注及质量检查方法；适合时调用已安装的 `ppt-master`。它是独立维护的可选协作技能，本仓库没有打包其完整源码或取代它的上游工作流。
- **迭代交付层**：优先沿用用户确认的设计，基于最新稿做局部修改，保留用户手工编辑、动画和未涉及页面，交付可继续修改的PPTX。

只调用本技能即可进入公司PPT工作流。PPT-Master可用时按需协作；缺少该依赖时使用本仓库的原生编辑路线，不自动安装或修改其他技能。职责与路由详见 [references/integrated-workflow.md](references/integrated-workflow.md)。

## 必读资源

1. 读取 `artifact-template.json`，相对于本技能目录解析 reference 和 preview。
2. 加载 [@presentations](plugin://presentations@openai-primary-runtime)，遵循其创建、渲染和验证流程。
3. 始终读取 `references/ooxml-safety.md`。
4. 选择封面、目录、章节或原有图形时，读取 `references/template-map.md`。
5. 选择展示方式、图表或图标时，完整读取 `references/visual-selection.md`，再按需要检索 `references/visual-resources.json` 并查看候选预览。不要只凭资源名称直接套用。
6. 连续修改已有稿件、增删页或处理动画时，读取 [references/revision-and-review.md](references/revision-and-review.md)。使用用户截图讲解软件时，读取 [references/screenshot-teaching.md](references/screenshot-teaching.md)。

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
- 沿用已经确认的风格和展示方向，不在每轮小改时重复发起风格确认。用户只要求大纲、部分章节或空白占位页时，按该范围完成，不自动填满整套内容。
- 概念、能力演进和工具使用优先用关系图、流程、对话、时间轴、真实截图或前后对照；产品多维比较与真实数据仍适合原生表格。避免每页都用表格，也不为追求形式强行图解数据。
- 允许自建可编辑表格、流程图、关系图和原生数据图表；沿用公司蓝、青绿、红色强调色、微软雅黑和页眉标识。
- 没有比例或反馈关系不用环形结构；没有真实定量数据不用柱状、折线或饼图。图标不能替代事实和检测证据。
- 模板不合适时创建简洁的新可编辑图形，不强求使用素材。品牌图标默认排除；同页图标风格、大小和视觉重量一致。
- 普通内容页关联原模板“内容”版式，标题写入真实标题占位符。不得叠加仿制标题框，也不得页面级覆盖坐标、字体、字号或颜色而破坏母版联动。
- 短篇演示不强制目录和章节；用户明确要求或有三个及以上自然章节时使用。目录条目、章节标题、实际顺序必须一致。
- 示例业务、年份、图例、数字和引用不是用户事实。引用后全部替换/删除，并同步图表的工作簿、缓存、公式范围和注释。
- 本技能优化展示形式、图表和图标选择，不包含医学突变位点筛选规则；用户若要求后者，先确认具体含义和证据要求。

## 推荐工作流

1. **确定起点与范围**：新建稿以公司参考包为基础；已有稿以最新交付或用户更新的PPTX为起点。先核对文件、活动页顺序和修改范围，不从陈旧SVG整体重建已编辑稿。
2. **组织内容**：结合受众、用途和时长拟定大纲；根据语义选择图解、截图或表格。已有设计共识可直接继续。新增设计与可选PPT-Master协作按 `integrated-workflow.md` 执行。
3. **制作可编辑正文**：封面、目录、章节优先复制原生结构后换字；新增普通页关联原内容版式。图解可自建，标题继续写入真实占位符。正文统一使用模板的安全区与对齐基线，避免内容长期过低；不无差别移动Logo、页脚和背景。
4. **局部修改与保留**：只改必要的文字、对象、页面和依赖；增删页时同步目录、章节、页码、备注和包元数据。原有母版、媒体和无关页面保持不变，保留未要求修改的动画。具体清单见 `revision-and-review.md`。
5. **验证并交付**：先结构与字体/几何检查，再PowerPoint原生渲染。局部修改检查改动页并确认其他页无意外变化；全新或全面重排稿逐页检查。延续用户命名，固定主题前缀加递增版本，不反复更换文件名后缀描述。

## 验收标准

- 输出中的slideMasters、slideLayouts、theme和模板已有media与参考文件字节一致。
- 封面、目录、章节的Logo、装饰条、章节数字和组合图形无错位；原生变换保留。
- 普通标题为 `p:ph type="title"`，`p:spPr` 与 `a:lstStyle` 为空，段落/文字运行不覆盖母版字体、字号、颜色等样式。
- 表格、数据图表和流程图保持可编辑；图片图标不宣称路径可编辑。数值、单位、比例、时间跨度与正文一致。
- 无不相关图标、示例数字或引用残留、文本溢出、遮挡、异常换行；结构检查不能替代视觉检查。
- 原生表格文字拥挤或自动检查提示溢出时，可用 `scripts/Test-MygeneNativeTableText.ps1` 测量实际单元格文字高度，并结合原生预览处理；不得仅为消除警告篡改验证器或伪造通过。
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
