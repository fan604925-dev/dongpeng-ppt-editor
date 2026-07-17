# 注册版式目录

新增页面时使用 `insert_template_slide` 和下列 `layout_id`。复制后的页面会写入 `dongpeng_layout` 标签，后续可用语义 `role` 编辑。

| layout_id | 模板页 | 类型 | 典型用途 |
|---|---:|---|---|
| `cover` | 4 | 正式封面 | 报告标题、英文副题、汇报人、日期 |
| `contents` | 5 | 目录 | 4–6 个章节 |
| `section-red` | 6 | 红底章节 | 强章节切换 |
| `statement` | 7 | 中央标语 | 品牌观点、核心结论 |
| `image-text` | 8 | 左文右图 | 空间图、案例图、产品图 |
| `map-chart` | 9 | 标题+图表/地图 | 区域、渠道、地图 |
| `body-text` | 11 | 纯文字 | 单一观点、说明段落 |
| `kpi-3` | 12 | 三栏 KPI | 三组关键指标 |
| `closing` | 13 | 红底收束 | 结束页、过渡页 |
| `section-light` | 16 | 浅底章节 | 类别分隔 |
| `three-columns` | 18 | 三栏文字 | 三个并列要点 |
| `swot` | 19 | 红底四象限 | SWOT、四象限分析 |
| `keyword-chain` | 21 | 关键词链 | 线性关系 |
| `keyword-three` | 22 | 三圆关键词 | 三个平级主题 |
| `keyword-radial` | 23 | 放射关系 | 中心主题与四个分支 |
| `keyword-formula` | 24 | 关键词公式 | 三项组合关系 |
| `image-pair` | 26 | 双图拼贴 | 两张空间图或案例图 |
| `product-hero` | 27 | 场景+产品 | 全幅效果图与单品信息 |
| `product-matrix` | 28 | 产品矩阵 | 系列、规格、纹理、型号 |
| `geometry` | 29 | 几何结构 | 层级、组织、结构关系 |
| `icon-four` | 31 | 四图标能力项 | 四项能力、价值、步骤 |
| `table` | 32 | 表格 | 季度、年度、项目台账 |
| `stair-red` | 33 | 红底阶梯图 | 强调增长或分层 |
| `stair-light` | 34 | 白底阶梯图 | 常规增长或分层 |
| `donut` | 35 | 环形图 | 构成比例 |
| `bars-timeline` | 36 | 条形图+时间线 | 进度与阶段说明 |
| `icon-library-light` | 38 | 红色图标库 | 复制红色图标 |
| `icon-library-red` | 39 | 白色图标库 | 复制白色图标 |

## 选择原则

- 图片占主导时选 `image-text`、`image-pair`、`product-hero` 或 `product-matrix`。
- 数据占主导时选 `kpi-3`、`stair-*`、`donut` 或 `bars-timeline`。
- 关系占主导时选 `swot`、`keyword-*` 或 `geometry`。
- 内容超过当前版式容量时拆页，不缩小正文。
- `closing` 和图标库页通常不需要替换背景资产。

