---
name: dongpeng-ppt-editor
description: Create, revise, replicate, inspect, render, and validate fully editable Dongpeng corporate PowerPoint decks with strict fidelity to the official red-and-white template. Use for 东鹏 PPT、集团标准模板、企业汇报、演示文稿复刻、PPTX 编辑、根据大纲生成可编辑 PPT、按批注修改页面、增删或排序页面、替换图片、调整文字与几何对象、导出 PDF/PNG，或要求其他 AI 软件也能按同一规则稳定生成东鹏风格文档的任务。
---

# Dongpeng PPT Editor

把官方模板视为唯一视觉真源。所有输出必须可编辑、可复现、可校验；不要把“看起来像”当作完成。

## 环境与资产

- 在装有桌面版 Microsoft PowerPoint 的 Windows 上运行。
- 使用 `assets/东鹏集团PPT标准模板.pptx`，不得覆盖母模板。
- 使用 `scripts/*.ps1` 执行检查、生成、导出和校验；环境要求授权时先申请 PowerPoint/GUI 权限。
- 画布固定为 960 × 540 pt、16:9。

## 先读取正确的资料

按任务选择，不要一次加载所有参考文件：

| 任务 | 必读 |
|---|---|
| 新建或复刻整套 PPT | [design-rules.md](references/design-rules.md)、[layout-catalog.md](references/layout-catalog.md)、[replication-workflow.md](references/replication-workflow.md) |
| 按批注修改现有 PPT | [common-errors.md](references/common-errors.md)、[operations-schema.md](references/operations-schema.md) |
| 其他 AI 软件接入 | [ai-integration.md](references/ai-integration.md)、[operations-schema.md](references/operations-schema.md) |
| 最终验收 | [quality-rubric.md](references/quality-rubric.md)、[validation-checklist.md](references/validation-checklist.md) |
| 参考已验证案例 | [case-studies.md](references/case-studies.md) |

## 不可跳过的工作流

1. **理解内容形状**：先整理受众、时长、页数、叙事主线和每页唯一结论。不要直接把大纲逐条贴进模板。
2. **选择版式**：为每页从 `assets/layout-map.json` 选择注册的 `layout_id`。版式容量不匹配时换版式或拆页，不缩小正文硬塞。
3. **建立候选规范**：在项目目录创建 `operations.json`，输出到候选文件名，`overwrite:false`。每页写演讲者备注。
4. **预检规范**：

   ```powershell
   powershell -NoProfile -ExecutionPolicy Bypass -File "<SKILL_ROOT>\scripts\validate-spec.ps1" -SpecPath "operations.json"
   ```

5. **生成候选版**：

   ```powershell
   powershell -NoProfile -ExecutionPolicy Bypass -File "<SKILL_ROOT>\scripts\edit-ppt.ps1" -SpecPath "operations.json"
   ```

6. **逐页看 PNG**：必须目视检查每一页；重点查标题换行、留白、装饰线与文字、模板残留对象、图片裁切、页脚安全区。不能只看 JSON 或自动校验。
7. **按候选迭代**：每次只修明确问题，重新生成候选版并检查受影响页及相邻页。不要直接覆盖正式版。
8. **生成正式版**：候选版通过后，把输出路径改为正式文件并仅在用户已授权时设置 `overwrite:true`。
9. **最终校验**：

   ```powershell
   powershell -NoProfile -ExecutionPolicy Bypass -File "<SKILL_ROOT>\scripts\validate-ppt.ps1" -InputPath "output.pptx" -OutputPath "validation.json"
   ```

10. **交付**：只有在 `warning_count = 0`、逐页视觉检查通过、PPTX/PDF 页数一致时交付。

## 三种执行模式

### 新建

- 使用 `new_deck:true`。
- 通过 `insert_template_slide` 添加页面，不从空白页绘制整套设计。
- 保持红色仪式页与白色信息页节奏；五页以上至少各两页。

### 修改

- 先运行 `inspect-ppt.ps1`，获取对象路径、名称和几何信息。
- 优先使用语义 `role`；未注册对象使用 `selector.path` 或带坐标容差的选择器。
- 修改装饰线时同时检查其相邻标题、正文和文本框内边距。

### 复刻

- 先建立“页码 → 内容目标 → 注册版式 → 必要定制 → 风险点”表。
- 复制模板结构，只替换内容；需要定制时保留品牌页眉、Logo、页脚、背景和几何语言。
- 用候选版与参考页并排比较，按 [quality-rubric.md](references/quality-rubric.md) 达到全部硬门槛。

## 核心设计约束

- 中文使用微软雅黑；英文和数字使用 Arial 或微软雅黑。
- 主强调色仅使用 `#E21413`，其他颜色只用登记的红色阶、黑白和中性灰。
- 普通正文不低于 12 pt；页眉标签 9 pt；版权 8.4 pt。
- 直角、平面、红白体系；不使用 SaaS 卡片、霓虹渐变、Emoji、混杂图标或图片拉伸。
- 场景照片使用 `cover`；产品图、截图、图表和信息图使用 `contain`。
- 底部 60 pt 为品牌安全区，正文、图注和图表不得进入。
- 文本框默认左右内边距 7.2 pt。线条与“可见文字”对齐时，要补偿内边距，不能只对齐文本框边界。

## 高频失败门槛

出现以下任一问题不得交付：

- 标题出现孤字换行、字号不一致或与红线重叠。
- 目录套错容量，内容集中在角落或遗留空章节占位。
- 多出或缺少箭头、短线、模板装饰对象。
- 分栏竖线未覆盖对应文字高度，或不同竖线不等高。
- 十字线穿过标题或正文。
- 内容整体重心过低/过高，大片留白没有叙事作用。
- Logo、页脚、背景、字体或颜色偏离模板。
- 只通过自动校验但没有逐页看预览图。

## 脚本职责

- `inspect-ppt.ps1`：读取页面、标签、对象路径、文字、类型和几何。
- `validate-spec.ps1`：在启动 PowerPoint 前检查操作规范、版式、字体、颜色、图片策略和导出要求。
- `edit-ppt.ps1`：执行页面、文字、图片、形状、表格、图表、备注、保存和导出操作。
- `validate-ppt.ps1`：检查最终画布、字体、颜色、溢出和品牌风险。

## 交付契约

交付可编辑 `.pptx`、同步 `.pdf`、`validation.json`；必要时保留 PNG 预览。简要说明修改内容和校验结果，不交付候选版或失败草稿作为正式文件。
