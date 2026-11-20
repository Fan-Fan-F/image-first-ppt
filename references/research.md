# GitHub 参考与实现取舍

- **PPT Rebuilder King** 的 R28 仍经过背景、整页前景和组件板生成。其公布的实际 PowerPoint 样本有 27 个原生文字对象和 14 张 PNG，同时出现了三张分离卡片合并、标题新增描边及字重和位置偏差。源图坐标和局部原图像素因此应作为主依据，验收同时检查逐元素内容、布局和视觉差异。原项目第一方代码和文档采用 MIT。 [项目仓库](https://github.com/frederick-kk/ppt-rebuilder-king) · [真实页面验收](https://github.com/frederick-kk/ppt-rebuilder-king/blob/main/docs/RELEASE_VALIDATION.zh-CN.md)
- **codex-ppt-skill** 以图片建立视觉风格，适合作为生图阶段的流程参考。**image-to-editable-ppt-skill** 提供逐页清单和局部返修的实现参考。新的流程可让本地裁切承担大多数图层提取，减少网页端反复重画。 [codex-ppt-skill](https://github.com/ningzimu/codex-ppt-skill) · [image-to-editable-ppt-skill](https://github.com/ningzimu/image-to-editable-ppt-skill)
- **px-image2pptx** 将 OCR 区域和局部字形像素遮罩求交，再局部补底并生成文字框。对边框、标记和图标邻近文字的场景，遮罩要贴合字形并经人工或渲染核对。PaddleOCR 可以提供多边形与识别置信度。 [px-image2pptx](https://github.com/JadeLiu-tech/px-image2pptx) · [PaddleOCR 输出格式](https://github.com/PaddlePaddle/PaddleOCR/blob/main/docs/version3.x/pipeline_usage/OCR.en.md)
- **pngtopptx** 从源图像素坐标驱动重建，单独核对包结构、对象可编辑性和渲染相似度。**img2pptx** 展示了文字、简单形状和局部 PNG 的混合重建，用已装字体清单和渲染差异定位返修区域。 [pngtopptx](https://github.com/CAPTW/pngtopptx) · [img2pptx](https://github.com/KeNan624/img2pptx)
- **SAM 2** 可用区域或点提示提取候选遮罩，原图 RGB 仍可供裁切；**Qwen-Image-Layered** 可产生多张 RGBA 层用于困难场景的对照候选，需重新核查位置、边缘与内容。复杂源图须结合遮挡关系决定是否合并为可移动图片组。 [SAM 2](https://github.com/facebookresearch/sam2) · [Qwen-Image-Layered](https://github.com/QwenLM/Qwen-Image-Layered)
- **PowerPoint** 支持指定像素尺寸导出实际渲染。DrawingML 对中文和其他东亚字形可显式指定字体，结合本机实际字体检验行高与换行。 [PowerPoint PNG 导出](https://learn.microsoft.com/en-us/office/vba/api/powerpoint.presentation.export) · [DrawingML EastAsianFont](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.eastasianfont?view=openxml-3.0.1)

## 2026-11-20 对比与 v1.1 取舍

v1.0 用可见遮罩裁下层图片，重叠处会变成透明。上层在原位时，渲染对照可能完全一致。移开上层后，透明缺口才会显现。R28 的公开验收也记录了有透明图层、但三张卡片被合并的案例。两者都说明需要检查实际可移动对象。 [v1.0 裁切实现](https://github.com/Fan-Fan-F/image-first-ppt/blob/v1.0.0/scripts/raster.py) · [R28 验收记录](https://github.com/frederick-kk/ppt-rebuilder-king/blob/main/docs/RELEASE_VALIDATION.zh-CN.md)

v1.1 将物体完整轮廓、原图可见像素和遮挡区补全分别处理。有完整透明素材时，用它提供隐藏部分和边缘依据。符合纯色条件的物体可按确认过的完整轮廓补齐。复核要同时看原位叠放和移开后露出的部分，检查轮廓、缺口及底图残影。复杂纹理、文字和结构若只存在于被遮挡的区域，需用源素材确认，或把估计结果标为待复核。

- **复杂图层候选。** Qwen-Image-Layered 能输出多张 RGBA 图层，官方模型卡标注 20B。CVPR 2025 的 Open-World Amodal Appearance Completion 明确分开可见区域分割、遮挡分析和隐藏区域补画。可参考这些机制接入可选补全来源，逐层与原图校对，避免将模型生成结果当作源素材。 [Qwen-Image-Layered 模型卡](https://huggingface.co/Qwen/Qwen-Image-Layered) · [CVPR 2025 实现](https://github.com/saraao/amodal)
- **字体和逐项校验。** pngtopptx 记录原字体到实际字体的映射，分别报告结构与视觉结果。PageIR 工作流使用 OCR 候选、逐对象校正，并提供 LibreOffice 渲染路线。后续可增加带置信度的 OCR 清单导入、缺字体与换行复核，以及跨平台渲染。 [pngtopptx 验收约定](https://github.com/CAPTW/pngtopptx#the-trust-contract) · [PageIR 工具链](https://github.com/algerchen2024/image-to-editable-pptx#local-toolchain)
