---
name: image-first-ppt
description: Generate slide design images first, then rebuild them as editable PowerPoint with native text, closely matched fonts and colors, movable source-pixel graphics, and before/after render checks. Use for image-to-editable-PPT, 生图后拆分成可编辑 PPT, screenshot reconstruction, or a manual ChatGPT-to-Codex handoff.
---

# 生图后重建可编辑 PowerPoint

用整页图片锁定版式、色彩和视觉风格。把可读文字重建为 PowerPoint 文字框，把独立的复杂图形拆成可移动图片。只有能够贴合原图的简单形状才原生重画。优先保证元素齐全、轮廓和位置接近原图。

## 1. 生成或接收原图

- 有原图就从原图开始。需要从主题生成时，先定每页的实际文案和画幅，再用图像生成工具出整页视觉稿。若用户想节省 Codex 侧调用，可在 ChatGPT 网页版手工生图和清点，按 [网页版交接](references/chatgpt-handoff.md) 接收文件。实际用量取决于所用的功能和账号。
- 同一套 PPT 锁定统一画幅、字体角色、色板和版心。提示生图工具将文字放在干净底面，减少文字压住细线、阴影或复杂物件。画幅需要修正时，在冻结前确定最终裁边或补边。禁止拉伸图片凑比例。
- 保留用户定稿，按页对照图片文字。模糊字、错误数字和 AI 拼写错误由定稿或原始数据校正，并记入复核记录。

## 2. 逐页清点和拆层

1. 用 `scripts/rebuild.py init` 冻结原图、尺寸、哈希和页序。只用这份原图的像素坐标。按 [清单和运行说明](references/scene.md) 填写清单。
2. 先看全页，再按重叠分区查漏。可用 `scripts/source_guides.py` 输出像素标尺和带原图坐标的裁片。扫标题、正文、数字、图例、数据标签、脚注，同时逐个登记有意义的视觉对象，特别注意小图标、细线、嵌套标签、重复卡片、边角装饰。放大确认断开的部件、阴影和孔洞。交叉对照 OCR、用户文案或清单中的 ID。
3. 将文字标为 `text`，分别记录原字清除范围和新文字框。清除范围只覆盖字形，避开邻接图案。中西文字体分别写入，先从本机候选中找相似字形和字重，再校准字号、行高、基线与颜色。文字颜色可从原图字形采样。
4. 从同一原图的去字版本截取 `image`。图案上有文字时先去字再裁切。矩形照片可直接按框截取。轮廓不规则、形状相交或带透明细节时，用多边形或二值遮罩标出实际可见像素，保留磕断的线段、影子和缝隙。原图里确实被挡住的部分可按组保留，或拆为可见的部分并记录移动限制。
5. 保留连续底色、纹理和不可独立恢复的景物。优先局部平涂或局部补底。若确实要用生成式修补，只提交遮罩范围的局部裁片，合成时锁住遮罩以外的像素。整页前景和组件板重画会改变原有物件位置和外观。

## 3. 构建、渲染、修复

```powershell
python scripts/rebuild.py init <job-dir> <slide-1.png> <slide-2.png>
# 填写 <job-dir>/manifest.json
python scripts/rebuild.py build <job-dir>/manifest.json <output.pptx>
python scripts/verify.py <job-dir>/manifest.json <output.pptx> <qa-dir> --renderer auto
```

CLI 均在本地运行。Windows 上优先用 PowerPoint 实际导出。如渲染不可用，交付中注明已执行的检查和未验证的项目。环境和命令参数见 [清单和运行说明](references/scene.md)。

- 逐页检查 `verify.py` 的清单标注图、实际渲染、并排图和差异图。清单里的每个 ID 都应在成品有对应对象或在复核记录有明确解释。重新扫原图以找出**清单本身漏掉**的东西。全页平均差异很低也可能漏掉一个小图标。
- 发现漏件补清单或拆窄遮罩。发现外形漂移先回到原始像素和真实轮廓。发现错字、残影、重叠或截断，局部改文字和去字范围，重建并重新渲染。主要文字需能直接选中和修改。测试移动一两个图片对象，看是否带着其它图案或有明显残影。
- 对可编辑度分开说明。文字是原生文本，标为 `shape` 的简单形状可修改外观，复杂图片可独立移动和替换。遮挡造成的不可见部分、贴合纹理的残影、艺术字样式的近似或保留底图的细节要明确标出。

GitHub 设计依据见 [研究记录](references/research.md)。
