# 公司级 PPT 制作技能

用于基于随附的公司 PowerPoint 模板创建或修改可编辑演示稿。技能保留原生母版、版式、组合图形、公司主题色和标题占位符联动，适用于内部汇报、技术交流和方案演示。

当前是三个部分组成的综合入口：公司模板与视觉资源、可选的[PPT-Master](https://github.com/hugohe3/ppt-master)内容/视觉制作协作、以及连续修改形成的交付习惯。PPT-Master作为独立技能维护，本仓库不内嵌其全部源码；没有该依赖也可按本技能的原生路线制作和修改。模板及素材的原有授权保持不变。

> 本仓库包含模板资源，仅限授权的内部用途；请保持仓库为私有，未经授权请勿分发。

## 当前参考资源

全部资源整合在同一份 `assets/reference.pptx`，共109页，不需要额外的图标库或展示方式库PPT。

- 第1–56页：原公司模板；基础版式、目录和章节规则不变，第56页仍为结束页。
- 第57–104页：已选的48种展示方式，统一为公司配色。
- 第105–109页：5张图标素材页，含432个对象（336个原生形状、96个图片对象，包含重复和品牌图标）。

只有需要的页面才进入成品。详细的选择优化、适用条件与调用注意事项见 `references/visual-selection.md`，准确页码和图标对象ID见 `references/visual-resources.json`。

## 调用

在 Codex 中选择或调用已安装的此技能，然后描述希望制作的演示稿主题、受众、页数和已有内容。

例如：

    使用公司PPT模板制作一份技术汇报，保留母版联动并加入目录和章节页。

## 关键规则

- 普通内容页标题写入原始标题占位符，而不是新建仿制标题框；修改母版标题样式后，已生成页面标题可以随之联动。
- 保留模板中的母版、版式、主题和原有媒体文件，避免复杂组合图形因重新保存导致 Logo、装饰线或章节数字错位。
- 根据内容再选择模板中的目录页、章节页、关系图、表格或图表资源；没有合适资源时，使用简洁的自建可编辑版式。
- 环形图仅用于真实的组成或循环关系；数据图表仅用于真实的定量数据。
- 目录、章节页与内容结构相匹配。短篇汇报不强制增加目录；有三个及以上自然章节或用户明确要求时，可加入目录和章节页。
- 完成后运行结构检查，并用 Microsoft PowerPoint 原生渲染复核，避免只依赖第三方渲染器。
- 连续修改基于最新PPTX，保留用户手改、原有动画和未涉及页面，不从旧SVG重建整稿；沿用已确认风格和固定文件名前缀。
- 概念、流程和操作优先用合适的图解与截图，多产品多维比较仍用原生表格；避免整套页面形式单一。
- 截图教学可居中框选功能区并就近标注，也可对照展示输入、操作和结果。保留原图，使用原生裁切，新增说明保持可编辑。

## 文件结构

    SKILL.md                         技能主说明与工作流
    artifact-template.json           模板注册信息
    assets/reference.pptx            保留的公司模板
    assets/preview.png               模板预览图
    references/ooxml-safety.md       母版、组合图形与标题联动规则
    references/template-map.md       目录、章节页和常用视觉资源图谱
    references/visual-selection.md  展示方式、图表和图标选择优化
    references/visual-resources.json 精确页码、对象ID与来源索引
    references/integrated-workflow.md 公司模板与PPT-Master协作、任务路由
    references/revision-and-review.md 最新稿局部修改、动画、版本与验收
    references/screenshot-teaching.md 软件截图、就近标注与操作示例
    assets/resource-previews/        单一参考中的资源页预览
    scripts/                         检索、选页、结构检查和原生PNG导出

## 检索与选页

在技能目录运行：

```powershell
pwsh.exe -NoLogo -NoProfile -File scripts/Get-MygeneVisualResource.ps1 -Kind method -Query 时间轴
pwsh.exe -NoLogo -NoProfile -File scripts/Get-MygeneVisualResource.ps1 -Kind icon -Query 日历
& ./scripts/Export-MygeneReferencePages.ps1 -Page 2,4,5,63,56 -OutputPath 'D:\工作目录\草稿.pptx'
```

选页只生成草稿，需要继续替换全部示例文字、数字、图表数据及引用。品牌图标默认不参与检索；无合适资源时可以自行绘制公司色的可编辑图形。图表不能只改文字或缓存而不更新工作簿。

选择优化指PPT视觉资源，不包含医学突变位点筛选规则。原模板部分历史图表有缺省类别/稀疏缓存，参考库保留原生结构；调用为成品时用真实数据完整更新或重建并验证，不能为了通过检查虚构零值。

## 安装到 Codex

将仓库克隆或复制到本机 Codex 技能目录。目录名可以自定义，但应保留仓库内的相对文件结构。

    C:\Users\Administrator\.codex\skills\<技能目录>

重启或刷新 Codex 的技能列表后，即可调用该技能。

## 验证成稿

运行 scripts 目录中的结构检查脚本，确认参考模板中的母版、版式、主题与现有媒体未被改写，并检查页数及标题占位符。

再运行原生导出脚本，使用本机 Microsoft PowerPoint 导出 PNG 预览，以 PowerPoint 的渲染结果作为最终视觉验收依据。

若第三方渲染器与 PowerPoint 的效果不一致，以 PowerPoint 原生渲染为准。

原生表格有文字高度/溢出疑问时，可只读测量指定页面，不保存源稿：

```powershell
pwsh.exe -NoLogo -NoProfile -File scripts/Test-MygeneNativeTableText.ps1 -PptxPath '输出文件.pptx' -Page 5 -ReportPath '构建目录/表格检查.json'
```

脚本返回实际单元格文字高度，不输出业务正文；没有表格、无法测量或高度溢出时返回失败。测量结果仍须结合原生预览判断宽度、图片、遮挡和可读性。
