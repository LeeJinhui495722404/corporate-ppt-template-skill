# 迈景（Mygene）公司级 PPT 模板技能

用于基于迈景（Mygene）公司级 PowerPoint 模板创建或修改可编辑演示稿。技能保留原生母版、版式、组合图形、公司色和标题占位符联动，适用于公司内部汇报、技术交流和方案演示。

> 本仓库包含公司模板资源，仅限授权的内部用途；请勿公开分发或移除私有仓库保护。

## 调用

在 Codex 中输入：

    使用 $artifact-template-mygene-ppt 制作一份迈景公司PPT，主题是……

也可以直接描述“使用迈景/ Mygene 公司 PPT 模板制作演示稿”，系统会自动匹配该技能。

## 技能保证的关键事项

- 普通内容页的标题写入原始标题占位符，而不是新建仿制标题框；修改母版标题样式后，已生成页面的标题可随之联动。
- 保留模板中的母版、版式、主题和原有媒体文件，避免复杂组合图形因重新保存而导致 Logo、装饰线或章节数字错位。
- 根据内容再选择模板中的目录页、章节页、关系图、表格或图表资源；没有合适资源时，使用简洁的自建可编辑版式。
- 环形图仅用于真实的组成或循环关系；数据图表仅用于真实的定量数据。
- 目录、章节页与内容结构相匹配。短篇汇报不强制增加目录；有三个及以上自然章节或用户明确要求时，可加入目录和章节页。
- 完成后运行结构检查，并用 Microsoft PowerPoint 原生渲染复核，避免只依赖第三方渲染器。

## 文件结构

    SKILL.md                         技能主说明与工作流
    artifact-template.json           模板注册信息
    assets/reference.pptx            保留的迈景公司模板
    assets/preview.png               模板预览图
    references/ooxml-safety.md       母版、组合图形与标题联动规则
    references/template-map.md       目录、章节页和常用视觉资源图谱
    scripts/Test-MygenePptx.ps1      结构一致性检查脚本
    scripts/Export-MygenePowerPointPreview.ps1
                                     PowerPoint 原生 PNG 导出脚本

## 安装到 Codex

将仓库克隆或复制到本机 Codex 技能目录：

    C:\Users\Administrator\.codex\skills\artifact-template-mygene-ppt

重启或刷新 Codex 的技能列表后，即可通过 $artifact-template-mygene-ppt 调用。

## 验证成稿

结构检查：确认参考模板中的母版、版式、主题与现有媒体未被改写，并检查页数及标题占位符。

    pwsh.exe -NoLogo -NoProfile -File scripts/Test-MygenePptx.ps1 -PptxPath "输出文件.pptx" -ExpectedSlideCount 15

原生导出：使用本机 Microsoft PowerPoint 导出 PNG 预览，以 PowerPoint 的渲染结果作为最终视觉验收依据。

    pwsh.exe -NoLogo -NoProfile -File scripts/Export-MygenePowerPointPreview.ps1 -PptxPath "输出文件.pptx" -OutputDirectory "预览目录"

若第三方渲染器与 PowerPoint 的效果不一致，以 PowerPoint 原生渲染为准。
