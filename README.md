# 东鹏 PPT 100% 复刻编辑器

基于东鹏集团标准模板创建、修改、复刻并校验 **完全可编辑的 PowerPoint 演示文稿**。

它不是一套只能替换文字的静态模板，而是一套面向 Codex、WorkBuddy 及其他支持 `SKILL.md` 的 AI 软件的完整制作规范，覆盖页面管理、文字、图片、形状、备注、导出和质量检查。

> 目标：在 Windows、桌面版 Microsoft PowerPoint、相同模板和字体环境下，稳定生成符合东鹏红白设计体系的 PPTX、PDF 和预览图。

## 能做什么

| 能力 | 说明 |
|---|---|
| 新建演示文稿 | 根据大纲自动选择模板版式并生成完整 PPT |
| 编辑现有 PPT | 修改文字、字体、字号、颜色、位置和页面结构 |
| 页面管理 | 新增、删除、复制、移动页面并调整页面顺序 |
| 图片处理 | 插入、替换、裁切、等比缩放和调整图片焦点 |
| 对象编辑 | 修改线条、箭头、色块、表格、图表和组合对象 |
| 演讲者备注 | 为每页写入演讲稿或讲解提示 |
| 多格式导出 | 保存可编辑 PPTX，并同步导出 PDF 和 PNG 预览 |
| 自动检查 | 检查版式、字体、品牌色、溢出、画布和导出要求 |
| 视觉验收 | 按 100 分质量规则逐页检查常见排版问题 |

## 工作流程

```mermaid
flowchart LR
    A[内容大纲或现有 PPT] --> B[选择东鹏注册版式]
    B --> C[生成 operations.json]
    C --> D[制作前预检]
    D --> E[生成可编辑 PPTX]
    E --> F[导出 PDF 和 PNG]
    F --> G[逐页视觉检查]
    G --> H[零警告正式交付]
```

## 环境要求

- Windows 10/11。
- 桌面版 Microsoft PowerPoint。
- PowerShell 5.1 或更高版本。
- 中文字体使用微软雅黑；英文和数字使用 Arial 或微软雅黑。
- Codex、WorkBuddy，或其他能够读取 `SKILL.md` 并执行 PowerShell 的 AI Agent。

WPS 可以打开最终 PPTX，但不建议用于执行生成脚本或最终视觉验收。

## 安装到 Codex

### 方法一：让 Codex 自动安装

在 Codex 中输入：

```text
请使用 skill-installer，从 https://github.com/fan604925-dev/dongpeng-ppt-editor 安装 Skill。
Skill 位于仓库根目录，安装名称为 dongpeng-ppt-editor。
```

安装完成后新建一个任务，让 Codex 重新加载技能。

### 方法二：手动安装

1. 点击仓库右上方的 `Code`，选择 `Download ZIP`。
2. 解压下载文件。
3. 将文件夹重命名为 `dongpeng-ppt-editor`。
4. 把整个文件夹放入：

```text
C:\Users\<用户名>\.codex\skills\dongpeng-ppt-editor
```

5. 重新打开 Codex 或新建任务。

## 安装到 WorkBuddy

如果当前版本支持从 GitHub 安装，在 WorkBuddy 中输入：

```text
请从 https://github.com/fan604925-dev/dongpeng-ppt-editor 安装根目录下的 dongpeng-ppt-editor Skill。
```

也可以下载 ZIP，解压并重命名后放入：

```text
C:\Users\<用户名>\.workbuddy\skills\dongpeng-ppt-editor
```

随后重新启动 WorkBuddy。

## 快速使用

### 根据大纲生成新 PPT

```text
使用 dongpeng-ppt-editor，根据下面的大纲制作一份 10 页东鹏标准 PPT。
先生成候选版，逐页检查排版，再交付可编辑 PPTX、PDF、PNG 预览和校验结果。

受众：集团设计与产品团队
演讲时长：15 分钟
内容大纲：……
```

### 根据批注修改现有 PPT

```text
使用 dongpeng-ppt-editor 修改这份 PPT。
保持模板、Logo、页脚和品牌色不变，逐条处理我的页面批注；
完成后重新导出 PPTX 和 PDF，并检查受影响页面及相邻页面。
```

### 复刻参考演示文稿

```text
使用 dongpeng-ppt-editor，把这份参考 PPT 的内容重组为东鹏标准风格。
每页只能使用注册版式；正文不得小于 12 pt；完成后按 100 分标准验收。
```

## 输出文件

标准任务应至少交付：

```text
项目名称/
├── 项目名称.pptx       # 可编辑演示文稿
├── 项目名称.pdf        # 同步 PDF
├── preview/             # 每页 PNG 预览
├── operations.json      # 可复现的编辑操作
└── validation.json      # 自动校验结果
```

正式交付要求：

- PPTX 与 PDF 页数一致。
- 自动校验 `warning_count = 0`。
- 每页完成视觉检查。
- 不存在文字溢出、孤字换行、重复装饰对象或错误图片裁切。

## 设计系统摘要

- 画布：16:9，960 × 540 pt。
- 主色：东鹏红 `#E21413`。
- 中文字体：微软雅黑。
- 内容页标题：通常为 34 pt。
- 正文：不低于 12 pt。
- 页面底部约 60 pt 为品牌安全区。
- 使用直角、平面、红白体系，不使用霓虹渐变、Emoji 或混杂图标。
- 新页面必须从注册模板版式复制，不能随意从空白页绘制整套设计。

完整规范见 [`references/design-rules.md`](references/design-rules.md)。

## 项目结构

```text
dongpeng-ppt-editor/
├── SKILL.md                         # AI 技能入口
├── agents/openai.yaml               # Codex 界面信息
├── assets/
│   ├── 东鹏集团PPT标准模板.pptx       # 标准模板
│   └── layout-map.json              # 注册版式及角色坐标
├── scripts/
│   ├── inspect-ppt.ps1              # 检查页面和对象
│   ├── validate-spec.ps1            # 制作前预检
│   ├── edit-ppt.ps1                 # 创建和编辑 PPT
│   └── validate-ppt.ps1             # 最终质量检查
└── references/                      # 设计、操作和验收规则
```

## 重要说明

- 不要覆盖 `assets` 中的母模板。
- PowerPoint、模板和字体环境不同，可能产生换行或几何偏差，必须重新渲染检查。
- 自动校验不能替代逐页视觉验收。
- 仓库公开不代表东鹏商标、Logo、模板或品牌资产已放弃相关权利。
- 使用者应遵守所在组织的品牌资产、信息安全和版权要求。

## 反馈与改进

如果发现版式错误、脚本问题或希望增加新的东鹏页面类型，可以在仓库中提交 Issue。反馈时建议附上：

- 问题页截图。
- 期望效果。
- PowerPoint 和 Windows 版本。
- 可公开的复现步骤或操作规范片段。

---

**Dongpeng PPT Editor** · 可编辑 · 可复现 · 可校验
