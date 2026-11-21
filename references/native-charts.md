# 数据审查与原生 PowerPoint 图表

把原始表格、研究数据或已经校对的观测清单当作数值来源。让视觉草稿决定图表的位置、色板和阅读顺序。现有截图中的标签可帮助定位需要核对的图表。用了生成图时，图内出现的数字、刻度、样本个数、线条高低和趋势都需要根据原始资料重建。

## 复算后再画

1. 记录权威来源、原始表格位置、单位、样本范围以及缺失值的处理。核对行列对应关系和日期。无原始数据时，向用户指出对应图表缺哪份资料或数值。数字经用户确认后可继续。
2. 独立计算正文、脚注和图表需要的中间量。统计页要对照检验方向、假设、显著性水平、自由度、统计量、p 值、区间和最终判断。保留未舍入值用于画图，正文通常只显示必要精度。计算型图还应记录函数、参数和生成方法。
3. 柱形图和条形图保留零基线。离散类别可用柱形或条形，等间距序列可用折线。数值横轴、参考线、分布和置信区间使用 XY，直接写入每个系列的真实横纵坐标。比较线、区间线、样本点都能作为独立系列。

## 清单中的图表元素

示例使用两组演示数值，源图预留了干净的图表区。

~~~json
{
  "id": "trial-points",
  "kind": "chart",
  "box": [120, 200, 1000, 445],
  "chart_type": "scatter",
  "series": [
    {"name": "观测值", "x_values": [1, 2, 3], "values": [4, 6, 5],
     "color": "#147D76", "line_width_px": 0, "marker": "circle"},
    {"name": "基准", "x_values": [1, 3], "values": [5, 5],
     "color": "#BE623A", "dash": "dash"}
  ],
  "x_axis": {"min": 1, "max": 3, "major_unit": 1, "title": "观测编号"},
  "value_axis": {"min": 3, "max": 7, "major_unit": 1, "title": "测量值"},
  "font": "Microsoft YaHei",
  "show_legend": true,
  "source_note": "文档演示数据",
  "erase_mode": "none"
}
~~~

- `chart_type` 支持 `column`、`bar`、`line`、`scatter`。前三类使用 `categories` 字符串数组，按顺序与每组 `values` 一一对应，并可指定 `category_axis_title`。XY 接受每组独立的数字数组 `x_values`，长度与对应 `values` 相同。
- `source_note` 必填，核查时写清链接、表号或计算出处。它会进入图表对象和构建报告。可用 `title`、`font`、`font_size_px`、`foreground`、`background`、`show_legend`、`data_labels`。线系列可用 `line_width_px`、`dash` 与 `marker`。数值轴选项有 `min`、`max`、`major_unit`、`title`、`number_format`、`visible`、`show_tick_labels`、`show_major_gridlines`。
- 底图已画过图表时，用 `erase` 框住整个旧图及残留标签，指定 `erase_fill` 或匹配尺寸的 `erase_patch`。默认清除范围是 `box`。确认原图是空白占位区时才用 `erase_mode: "none"`。清单编辑完成后再生成原生图表，新图上的文字和数据随 PowerPoint 对象保存。

## 交付前核对

- `verify.py` 检查图表类型、系列及数字，沿 PowerPoint 文件的关系找到内嵌 Excel，逐个核对系列实际引用的单元格。验收以最终 PPTX 和审核后的清单为准。
- 用 PowerPoint 渲染，检查柱子的零点、XY 数据间距、轴单位、裁切、图例、色彩区分、数据标签和旧图表残影。视觉差异图只是定位辅助，图稿里的数值错误本就需要修正。
- 抽查 PowerPoint 的「编辑数据」，核对能定位到同一组原始数值。交付检查以实际数据、统计结论、画面和可编辑性都正确为完成条件。
