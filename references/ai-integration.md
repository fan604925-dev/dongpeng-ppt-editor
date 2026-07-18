# 其他 AI 软件接入规范

## 核心原则

其他 AI 不直接操作 PPT 视觉界面，而是生成统一的 `operations.json`，再调用本 Skill 的 PowerShell 脚本。这样可避免不同模型、不同编辑器产生不可控的样式差异。

## 必需输入

- 本 Skill 目录路径。
- 用户的大纲、文档或现有 PPTX。
- 项目输出目录。
- 可选图片/数据资产。
- 是否允许覆盖正式文件。

## 必需输出

- `operations.json`。
- 可编辑 `.pptx`。
- 同步 `.pdf`。
- PNG 预览目录。
- `validation.json`。

## 可复制给其他 AI 的任务提示

```text
使用 dongpeng-ppt-editor Skill 制作或修改东鹏集团标准 PPT。
必须：
1. 读取 SKILL.md 及任务对应的 references；
2. 只使用 assets/layout-map.json 中注册的 layout_id；
3. 先生成不覆盖正式文件的候选版 operations.json；
4. 运行 validate-spec.ps1、edit-ppt.ps1；
5. 逐页检查 PNG，修复标题换行、留白、线条/箭头、图片裁切和品牌元素；
6. 候选版通过后再生成正式 PPTX/PDF；
7. 运行 validate-ppt.ps1，必须 0 warning；
8. 交付可编辑 PPTX、PDF 和校验结果。
不得从空白页自由设计，不得覆盖母模板，不得使用未登记字体和颜色。
```

## 调用顺序

```powershell
# 1. 如为修改任务，先检查原文件
powershell -NoProfile -ExecutionPolicy Bypass -File "<SKILL_ROOT>\scripts\inspect-ppt.ps1" -InputPath "input.pptx" -OutputPath "inspection.json"

# 2. 检查 AI 生成的操作规范
powershell -NoProfile -ExecutionPolicy Bypass -File "<SKILL_ROOT>\scripts\validate-spec.ps1" -SpecPath "operations.json" -OutputPath "spec-validation.json"

# 3. 生成 PPTX/PDF/PNG
powershell -NoProfile -ExecutionPolicy Bypass -File "<SKILL_ROOT>\scripts\edit-ppt.ps1" -SpecPath "operations.json"

# 4. 检查最终 PPT
powershell -NoProfile -ExecutionPolicy Bypass -File "<SKILL_ROOT>\scripts\validate-ppt.ps1" -InputPath "output.pptx" -OutputPath "validation.json"
```

## 机器可判定状态

- `spec-validation.json.valid = true`：规范可以执行。
- `validation.json.valid = true` 且 `warning_count = 0`：结构校验通过。
- 最终完成状态还必须包含人工/视觉模型逐页确认；不能仅依赖机器状态。

## 可移植性

- 相对路径以 `operations.json` 所在目录为基准。
- 跨机器共享时优先使用相对路径，不写用户目录和临时目录。
- 必须在 Windows + Microsoft PowerPoint 环境执行最终生成；其他系统可负责生成 JSON、内容和审阅意见。

