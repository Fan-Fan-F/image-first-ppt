# 像素坐标、清单和验收

## 本地依赖与运行

使用 Python 3.10 或更新版本。优先调用已经装好依赖的解释器。需要主动配置时，可在工作目录建 Python 虚拟环境，再安装技能目录下的 `requirements.txt`。Windows 实际导出需要 PowerPoint 桌面版和 Windows PowerShell。权限或渲染器不可用时，`--renderer none` 可单跑结构检查。若已用演示软件导出同尺寸 PNG，可用 `--render-dir <dir>` 导入，按 `slide-001.png`、`slide-002.png` 命名。

`scripts/rebuild.py init <job-dir> <image...>` 按输入顺序生成 `s001.png` 等锁定的 RGB 原图与 `manifest.json`，同时应用照片方向元数据，将透明区铺白。后续用该目录的 PNG 做清点。支持任意统一的宽高比。前期调整画幅用裁切或补边。CLI 参数 `--slide-width-in` 默认约 13.333 英寸，它决定默认页面和字号的物理尺寸。

`python scripts/source_guides.py <job-dir>/manifest.json <job-dir>/guides --page s001` 导出像素标尺和相互重叠的原尺寸裁片，可在清单为空时先运行。裁片文件名和 `source-guides.json` 都附原图坐标。在裁片里识别的框需要加上裁片的原图左上角坐标，再写回整页清单。页数多时可按需逐页生成导览。

## 清单格式

- 顶层为 `{"version":1,"slide_width_in":13.33333333,"slides":[...]}`。保留 `init` 写入的 `id`、`source`、`width`、`height` 和 `sha256`，用原图分析结果更新各页的 `elements`。网页清点一般只返回单页对象，需要按 ID 合回本地清单。
- `elements` 从底向顶顺序排列。每页对象使用唯一 ID，限英文、数字、连字符、下划线，首字符为英文字母或数字。每个对象的 `box` 是左上为原点的整数 `[x,y,w,h]`，左上包含、右下不含，单位为源图像素。
- 所有颜色写成 `#RRGGBB`。可在页面和元素写 `review` 说明仍需确认的对象、分组、遮挡或刻意更改。把只有固定底层的装饰和有意义的可移动元素区分开，连续底图之外的独立元素用 ID 记录。

| 类型 | 必填和常用字段 |
|---|---|
| `text` | `text` 为最终文字，实际换行写 JSON 的 `\n`。可用 `font`、`font_ea`、`font_px`、`bold`、`italic`、`color`、`align`（left/center/right）、`valign`（top/middle/bottom）、`rotation`（度，顺时针）、`line_height_px`。局部混合字号或颜色可用多个文字框精确摆放。 |
| `image` | `box` 是放置范围。可用 `polygon` 或 `mask` 选取原图内的像素。零件中断、孔洞或复杂轮廓优先用二值图。 |
| `shape` | `shape` 选 rect、round_rect、ellipse 或 triangle，附明确的 `fill`。可用 `stroke`、`stroke_px`。原图上占据的像素范围也支持 `polygon` 或 `mask`。PPT 默认圆角与原图差异大时改回图片层。 |

### 文字去除和图形补底

- `erase` 是原始字形所在的 `[x,y,w,h]`，默认等同 `box`，可以独立于新文字框调整。按需要给 `erase_polygon`（顶点都在 `erase` 内）或 `erase_mask`（相对目录下与原图同尺寸的二值 PNG，选中像素都在 `erase` 内）。环绕花纹处把擦除范围收窄成字形，或分多段登记。多色字体、描边和艺术字需仔细核对。
- `erase_fill` 表示此字下已知底色。省略时 `erase_mode` 默认 auto，根据局部边缘选择平涂或局部补洞，可指定 flat 或 inpaint。补洞只能估计原图中被覆盖的像素，对文字压住复杂纹理的情况必须检查边缘。
- `polygon` 的顶点是位于 `box` 内的原图整数像素。 `mask` 是相对清单目录、与原图同宽高的二值 PNG，白像素选中，黑色或全透明像素忽略。使用固定大小的二值图可表达不相连的部件和孔洞。
- `underlay_fill` 用于已知的图形背后底色。省略会局部估计。重叠处按从底至顶的实际归属分配可见像素。透明边缘按原稿中已合成的 RGB 保留，放在原位更贴合，移动到不同颜色的底图上可能显示原底色边缘。原图中完全遮挡的内容需要原素材才能恢复。
- 局部补图可用 `erase_patch` 或 `underlay_patch` 引用本地 PNG。其尺寸必须与 `erase` 或 `box` 完全相同。去字补图需要窄 `erase_mask` 或 `erase_polygon`，去图补底只采用当前物体可见遮罩范围内的像素。补图和显式填充色二选一。导入后检查边缘是否接上。
- 文字颜色未知时脚本从原图与清除后图像的差异估计字色。单行、未旋转的文字会结合本机字体样本匹配字宽、字高和笔画。字体、字色、字号仍要看真实渲染。显式填写字体名可覆盖选择，交付前核对接收方机器是否具备相同字体。

## 工作样例

```json
{
  "version": 1,
  "slide_width_in": 13.33333333,
  "slides": [{
    "id": "s001", "source": "s001.png", "width": 1600, "height": 900,
    "elements": [
      {"id":"photo","kind":"image","box":[868,128,620,640]},
      {"id":"tag","kind":"shape","box":[92,530,436,100],"shape":"rect","fill":"#F1F5FA","underlay_fill":"#FFFFFF"},
      {"id":"title","kind":"text","box":[96,116,740,142],"erase":[100,147,650,78],
       "text":"给城市留一点绿","font":"Arial","font_ea":"Microsoft YaHei",
       "font_px":75,"color":"#0F2431","bold":true}
    ],
    "review": []
  }]
}
```

## 检查结果的含义

- 构建报告列出各对象的像素处理、文字设置以及遮罩外改动。检查擦除范围、完全遮挡及色彩/字号有待复核的条目。
- `build` 的 `--layers-dir <dir>` 可导出处理后的底图和各个裁切层。检查底图里是否留有图形残影。在 PowerPoint 中移动有重叠的图层时另查露出的背景与不可见边缘。
- 结构检查从实际 PPTX 读取对象名、种类、叠层、位置和文字。渡过结构检查后仍需看原图清点是否完整。渲染差异是定位工具，文字正常的抗锯齿变动与严重漏件都可能在差异图亮起。逐个核对文字内容、基线、视觉字重、轮廓和物体边缘。
- 该工作流重建单页截图能观察到的内容。复杂图表的数据源、被遮挡的结构、图层的原始透明度和字体的精确身份，需要原文件或额外资料确认。复杂图表若需编辑数据，拿到可靠数据后用原生图表重做并逐值比对。
