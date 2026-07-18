# 东鹏 PPT 100% 复刻工作流

## 目录

1. 复刻的定义
2. 输入标准化
3. 叙事与版式决策
4. 候选版生成
5. 视觉复核
6. 正式版与验收

## 1. 复刻的定义

“100% 复刻”不是让不同 AI 自己画一套相似页面，而是让不同 AI 生成同一份操作规范，再由同一模板和脚本执行。复刻对象包括：

- 960 × 540 pt 画布和注册版式。
- 品牌背景、Logo、页眉、页脚、颜色和字体。
- 标题、正文、线条、图片和图表的几何关系。
- 红底仪式页与白底信息页的节奏。
- 可编辑 PPTX、PDF 和预览图的交付格式。

只要模板、`layout-map.json`、操作 JSON 和 PowerPoint 版本一致，生成结果应保持确定性。

## 2. 输入标准化

开始前把需求整理为以下结构：

```yaml
title: 汇报主题
audience: 受众
duration_minutes: 15
slide_count: 8-12
goal: 听众结束后应记住什么
sections:
  - name: 章节名
    takeaway: 唯一结论
    evidence: 数据/图片/案例
assets:
  - path: 图片路径
    usage: 场景图/产品图/截图
constraints:
  - 必须保留的文案或数据
```

如果只有思维导图或长文，先压缩成“每页一个结论”，再生成 PPT。每页正文建议不超过三组信息；超过时拆页。

## 3. 叙事与版式决策

### 3.1 建立页级表

在写操作 JSON 前完成：

| 页码 | 页面任务 | 唯一结论 | 内容形状 | layout_id | 主题 | 风险 |
|---:|---|---|---|---|---|---|
| 1 | 开场 | 为什么值得听 | 标题+副题 | `cover` | 红 | 标题换行 |
| 2 | 导航 | 全文结构 | 4–6 章节 | `contents` | 白 | 容量不匹配 |
| 3 | 观点 | 一句核心判断 | 标语+3 解释 | `statement` | 白 | 模板短线残留 |

### 3.2 按内容形状选版式

- 单一观点：`statement`、`body-text`。
- 3 个平级要点：`three-columns`、`keyword-three`、`kpi-3`。
- 4 个象限：`swot`、`icon-four`。
- 流程或演进：`bars-timeline`、`stair-*`、`keyword-chain`。
- 图片证据：`image-text`、`image-pair`、`product-hero`。
- 结尾：`closing`。

不要只按“看起来好看”选版式。内容形状不匹配是后续文字挤压、线条错位和留白失衡的主要原因。

### 3.3 规划视觉节奏

- 第 1 页使用红底封面。
- 每 3–4 页至少出现一次红底仪式页或强视觉页。
- 连续三页不得使用同一主体结构。
- 信息密集页后安排观点页、红底页或大字号页缓冲。

## 4. 候选版生成

### 4.1 规范文件

- 在项目目录创建 `operations.json`。
- 候选输出使用 `_candidate`、`_draft2` 等独立文件名。
- `overwrite:false`。
- 每页添加 `set_speaker_notes`。
- 同时配置 `export_pdf` 和 `export_png`。

### 4.2 先预检再启动 PowerPoint

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File "<SKILL_ROOT>\scripts\validate-spec.ps1" -SpecPath "operations.json"
```

修复所有 error；warning 必须逐项判断，不允许无视。

### 4.3 生成候选版

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File "<SKILL_ROOT>\scripts\edit-ppt.ps1" -SpecPath "operations.json"
```

## 5. 视觉复核

### 5.1 必须逐页检查

依次看 PNG，不要只抽查：

1. 先看整体重心：页面是否上空下挤、左右失衡。
2. 再看标题：是否孤字换行、字号不一致、层级过弱。
3. 再看几何：短线、竖线、箭头、十字线是否与文字对齐或重叠。
4. 再看正文：是否低于 12 pt、行距过密、内容超出版式容量。
5. 再看品牌：Logo、页眉、页脚、红晕、安全区是否完整。
6. 最后看相邻页：视觉节奏是否重复或突然跳脱。

### 5.2 修订原则

- 先修版式选择，再修文案，再修字号，最后微调坐标。
- 调整线条时，用可见文字边缘而不是文本框边缘对齐；默认文本框左右内边距约 7.2 pt。
- 删除模板残留对象时使用准确 `selector.path`，不要按模糊坐标删除多个对象。
- 每轮只修明确问题，保留候选文件便于回看。

## 6. 正式版与验收

候选版通过后：

1. 将输出改为正式 `.pptx`、`.pdf` 和正式预览目录。
2. 只有用户授权覆盖时才设置 `overwrite:true`。
3. 重新生成，不要直接复制候选文件改名。
4. 运行 `validate-ppt.ps1`。
5. 对修改页再做一次视觉复核。
6. 按 [quality-rubric.md](quality-rubric.md) 达到 100 分和全部硬门槛后交付。

