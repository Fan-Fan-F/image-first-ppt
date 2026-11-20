# 让被挡住的图案也能完整拖动

一张合成图只记录露出的像素。重建时，`box` 是整个图案的放置范围，`polygon` 或 `mask` 描述原稿能看到的轮廓，`completion` 描述独立对象需要的完整外观。数组仍从底到顶排列。形状、图片各用一个独立对象。

## 选择补全方式

- 原本握有完整透明图时，用 `completion.mode: "rgba"`，`origin: "original"`。文件为与 `box` 同尺寸的 PNG，按原坐标对齐。最好留存同一张图稿在叠放前的素材。为补全复杂隐藏细节制作的候选 PNG 用 `origin: "reconstructed"`。
- 确定整块都是单色、形状也明确时，用 `completion.mode: "solid"`，提供完整 `geometry`、`polygon` 或 `mask` 之一。构建器会检查可见像素是否支持单色假设，再补上被挡的区域。需要在 PPT 里改填色时，可选吻合轮廓的原生 `shape`，它本身就包含完整几何。
- 照片、纹理和插画可在网页端或外部工具中只修遮挡区域。返回与原图中整个对象同尺寸的透明 PNG，保留对象全部零件、留白孔洞和外部透明。优先用额外照片或素材对齐候选。无法确认的遮挡纹理只能近似，并应在交付时指出。

```json
{
  "id": "blue-card",
  "kind": "image",
  "box": [120, 80, 280, 160],
  "completion": {"mode": "solid", "geometry": "rect", "color": "#346D9E"},
  "underlay_fill": "#F4F3EF"
}
```

```json
{
  "id": "back-illustration",
  "kind": "image",
  "box": [710, 260, 420, 320],
  "mask": "back-illustration-visible.png",
  "completion": {
    "mode": "rgba",
    "file": "back-illustration-full.png",
    "origin": "original",
    "visible": "asset",
    "mask": "back-illustration-silhouette.png"
  }
}
```

两个示例作为单页 `elements` 中的项目使用。把挡在前面的元素写在相应项目后面。最终 PNG 像素归属有三条规则：

1. 挡住部分由对齐的完整候选层提供，其非透明部分可以延伸到已登记的前层下面。露出的区域须和源图轮廓吻合。RGBA 的可选 `completion.mask` 是对照用的全画布二值轮廓，选中的像素应与 PNG 中所有 Alpha 大于零的像素一致。对完全被挡住的对象，只有已知叠放关系并有对齐的原始素材时才登记，且需提供相符的轮廓或使用原图尺寸内的矩形范围。
2. `visible` 默认 `"source"`，用原图的可见 RGB 锁住配色、位置和细节。边缘来自合成后的 RGB，移到新底色时可能看到旧底色。拥有真实透明素材且确认配色、光影和大小一致时，可以指定 `"asset"` 保留原始透明通道，再用原位渲染核对。
3. `solid` 的 `geometry` 支持 rect、ellipse、round_rect 和 triangle。round_rect 可指定 `radius_px`。其它完整轮廓用全图二值 `mask` 或原图整数 `polygon`。纯色估计使用源图露出的颜色，有阴影或纹理时用完整 PNG。

## 验收原位和独立对象

1. 用 `build --require-complete` 生成正式候选。没有补全的确定或疑似重叠会在 `*.build.json` 中列出。疑似贴边但经原素材或放大核对确认没有遮挡时，可用 `"occlusion": "none"` 标记。它不能消除确实重合的遮罩。
2. 用 `verify.py` 确认原位叠放后，图案的位置、颜色和文字仍贴近原图。
3. 用 `check_layers.py` 导出 `isolated.pptx`，打开检查每个重叠对象自身的边界、内部纹理和天然镂空。也可在副本里实际移走挡在上方的对象，检查原位留下的下层是否完整、底板是否有残影。目标范围可用 `--object s001/blue-card` 选择。
4. 在交付时分清完整原素材、按几何重建、近似补画和仍待确认的对象。导入补图时必须看到实际对象，尤其检查被遮挡的范围如何连接到露出部分。
