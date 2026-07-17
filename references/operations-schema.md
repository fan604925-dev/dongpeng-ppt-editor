# 编辑操作规范

## 目录

1. 顶层结构
2. 页面操作
3. 文本操作
4. 图片与对象操作
5. 表格操作
6. 导出操作
7. 选择器
8. 示例

## 1. 顶层结构

```json
{
  "new_deck": true,
  "input": "可选：已有PPT.pptx",
  "template": "可选：覆盖默认模板路径",
  "output": "交付/输出.pptx",
  "overwrite": false,
  "operations": []
}
```

- 新建时设置 `new_deck: true`，不提供 `input`。
- 修改已有文件时提供 `input`，默认写入不同的 `output`。
- 只有用户明确允许覆盖时才设置 `overwrite: true`。
- 相对路径以操作 JSON 所在目录为基准。

## 2. 页面操作

```json
{"op":"insert_template_slide","layout_id":"cover","position":1}
{"op":"duplicate_slide","slide":2,"position":3}
{"op":"delete_slide","slide":4}
{"op":"move_slide","slide":6,"to":2}
{"op":"set_slide_tag","slide":1,"name":"project_section","value":"overview"}
```

页面可用 `slide`、`slide_id` 或 `slide_layout` + `occurrence` 选择。

## 3. 文本操作

```json
{
  "op":"set_text",
  "slide":1,
  "role":"cover_title",
  "text":"2026 年度经营报告",
  "font":{"name":"微软雅黑","size":40,"bold":false,"color":"#FFFFFF"},
  "paragraph":{"alignment":"left"}
}
```

可选字段：`left`、`top`、`width`、`height`、`rotation`、`margins`、`font`、`paragraph`。

新增文本框：

```json
{"op":"add_textbox","slide":2,"left":60,"top":120,"width":320,"height":80,"text":"新文字","font":{"name":"微软雅黑","size":34,"color":"#000000"}}
```

## 4. 图片与对象操作

替换模板图片：

```json
{"op":"replace_image","slide":2,"role":"hero_image","path":"images/room.jpg","fit":"cover","focus_x":0.5,"focus_y":0.35}
```

插入新图片：

```json
{"op":"insert_image","slide":2,"path":"images/product.png","left":600,"top":150,"width":220,"height":270,"fit":"contain"}
```

`fit` 支持 `cover`、`contain`、`stretch`。除非用户明确要求，不使用 `stretch`。

其他对象操作：

```json
{"op":"set_shape","slide":2,"selector":{"name":"Rectangle 1"},"fill":{"color":"#E21413","transparency":0},"line":{"visible":false}}
{"op":"delete_shape","slide":2,"selector":{"name":"TextBox 7"}}
{"op":"duplicate_shape","slide":2,"selector":{"name":"Picture 1"},"left":500,"top":200}
{"op":"add_shape","slide":2,"shape_type":"rectangle","left":60,"top":200,"width":200,"height":80,"fill":{"color":"#E21413"},"line":{"visible":false}}
```

从模板第 38/39 页复制图标或其他对象：

```json
{"op":"copy_shape_from_template","slide":2,"template_slide":38,"selector":{"left":121.1,"top":93.05},"left":80,"top":250,"width":40,"height":44}
```

## 5. 表格操作

```json
{"op":"set_table_cell","slide":3,"role":"table","row":2,"column":3,"text":"125","font":{"name":"Arial","size":12,"color":"#000000"}}
{"op":"add_table_row","slide":3,"role":"table"}
{"op":"delete_table_row","slide":3,"role":"table","row":5}
{"op":"add_table_column","slide":3,"role":"table"}
{"op":"delete_table_column","slide":3,"role":"table","column":4}
```

修改模板图表数据：

```json
{
  "op":"set_chart_data",
  "slide":4,
  "role":"chart",
  "categories":["Q1","Q2","Q3","Q4"],
  "series":[{"name":"占比","values":[28,32,25,15]}],
  "data_labels":"percentage",
  "title":"季度构成"
}
```

`data_labels` 支持 `percentage`、`value`、`none`；省略时保留现有标签。`set_chart_data` 会保留原图表类型和视觉样式，但需要本机 Excel 支持 PowerPoint 的嵌入式图表数据编辑。执行后必须重新渲染并检查标签、图例和色彩。

## 6. 导出操作

```json
{"op":"save"}
{"op":"export_pdf","path":"交付/输出.pdf"}
{"op":"export_png","directory":"交付/preview","width":1600,"height":900}
```

## 7. 选择器

优先使用已注册的 `role`。没有角色时使用：

```json
{"selector":{"name":"Picture 3","left":420,"top":0,"tolerance":3}}
```

支持字段：`name`、`path`、`type`、`left`、`top`、`text_contains`、`tolerance`、`match_index`。

先运行 `inspect-ppt.ps1` 获取准确对象路径与几何位置。

## 8. 完整示例

```json
{
  "new_deck": true,
  "output": "交付/东鹏经营报告.pptx",
  "overwrite": false,
  "operations": [
    {"op":"insert_template_slide","layout_id":"cover","position":1},
    {"op":"set_text","slide":1,"role":"cover_title","text":"2026 年度经营报告"},
    {"op":"set_text","slide":1,"role":"cover_subtitle","text":"DONGPENG ANNUAL BUSINESS REVIEW"},
    {"op":"insert_template_slide","layout_id":"image-text","position":2},
    {"op":"set_text","slide":2,"role":"title","text":"渠道增长与终端焕新"},
    {"op":"replace_image","slide":2,"role":"hero_image","path":"images/showroom.jpg","fit":"cover"},
    {"op":"insert_template_slide","layout_id":"closing","position":3},
    {"op":"export_pdf","path":"交付/东鹏经营报告.pdf"},
    {"op":"export_png","directory":"交付/preview"}
  ]
}
```
