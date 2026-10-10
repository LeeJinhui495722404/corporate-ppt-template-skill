# OOXML 与母版安全规则

## 为什么会错位

迈景模板包含嵌套组合图形，例如 p:grpSp，以及 a:chOff、a:chExt、flipV 等坐标变换。部分 PPT 库在导入、复制或重新保存时会展开组合、重算坐标或丢失变换，因此会出现 Logo、装饰条、章节数字和封面元素错位。

## 安全生成方式

- 把 assets/reference.pptx 作为 ZIP/OOXML 基础包，输出到新文件；绝不覆盖参考文件。
- 完整保留 ppt/slideMasters、ppt/slideLayouts、ppt/theme 和模板已有 ppt/media。
- 复制原生页面时同时复制 slide XML 和对应 rels；仅在必要时去除备注关系，保留版式和图片关系。
- 封面、目录和章节页优先只替换 a:t 文本，不重建复杂图形。
- 新内容页应关联原始内容版式。当前模板通常对应 slideLayout2.xml，但每次仍应从页面关系中确认。
- 新增或删除页面时同步更新 ppt/presentation.xml、ppt/_rels/presentation.xml.rels、[Content_Types].xml 和 docProps/app.xml 的页数信息。

## 标题占位符联动

普通内容页标题必须满足：

- 使用 p:ph type="title"。
- 标题形状的 p:spPr 为空，不写页面级 x、y、cx、cy。
- a:lstStyle 为空，不在页面上覆盖字号、字体、缩进或项目符号。
- 段落和文字运行的pPr/rPr/defRPr/endParaRPr也不得覆盖母版的字体、字号、颜色等；不能仅清空坐标就认定母版联动正常。
- 不创建覆盖在占位符上的独立标题文本框。

目录和章节页可保留各自专用版式中的原生几何；复制原页后只替换文字。

## 视觉资源选择

- 流程图、关系图、卡片、层级图必须与内容结构一致。
- 环形图只用于真实比例、组成或循环。
- 数据图表必须有真实定量数据；图标和装饰不能替代证据。
- 模板没有合适资源时，创建克制、可编辑的新页面。
- 选择方法和图标时读取visual-selection.md，按visual-resources.json定位。普通内容标题继续使用原版式，不采用外部模板标题。
- 新增资源使用vr_business_/vr_icons_部件名，遍历活动页应通过presentation的关系与sldId顺序，不按slide数字.xml推断页码。
- 复制图表/SmartArt/图标时保留charts、embeddings、diagrams及相关rels。保留组合坐标、旋转、层级；不能整体缩放原母版来适配正文。

## 最终验证

1. 运行 scripts/Test-MygenePptx.ps1。
2. 运行 Presentations 技能要求的最终检查。
3. 用 scripts/Export-MygenePowerPointPreview.ps1 调用 PowerPoint 原生导出。
4. 至少检查封面、目录、每张章节页、普通标题、Logo、组合图形、表格和页脚。

复杂组合图形在 LibreOffice 或其他第三方渲染器中可能与 PowerPoint 不一致。若结构检查通过，以 Microsoft PowerPoint 原生渲染作为视觉验收依据。
