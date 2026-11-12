---
name: huiben-02-ppt
description: 把绘本原件（扫描页 / 翻拍照片 / 截图 / 已有逐页插画）转换成可直接交给 huiben-01-creater 做视频的逐镜 PPTX：一镜一页，每页 = 本机 ComfyUI 生成的无文字 16:9 底图 + Python 叠加的中文原文 + PPTX 内可编辑文本框。**只要用户提到"绘本转 PPT""把绘本做成 PPT""绘本原件""扫描页拆镜""逐镜拆页""给绘本配图排版""用 Comfy 画绘本 PPT""Comfy 出图、文字 Python 加"，或发来绘本扫描页/翻拍照片并要求整理成 PPT/逐页文档，就用这个 skill**。也适用于只做其中一段：只转写文字、只拆镜、只补底图、只重排版。下游 skill 是 huiben-01-creater（绘本短视频生成）。
description_zh: 绘本原件 → ComfyUI 无字底图 + Python 叠字 + 可编辑逐镜 PPTX
description_en: Picture-book scans to per-shot PPTX with local ComfyUI illustrations and Python-rendered Chinese text
version: 2.2.2
allowed-tools: Bash, Read, Write, Edit
display_name: "绘本PPT生成"
tags:
  - huiben
  - ppt
  - picture-book
  - image-edit
---

# 绘本 PPT 生成（huiben-02-ppt）

把一本绘本的原件页面或已有逐页插画，转成一份逐镜 PPTX。默认做法是：

> **本机 ComfyUI 只画无文字 16:9 底图；所有中文文字、数字、引号、标点由 Python/Pillow 叠加；PPTX 里同时保留可编辑的楷体文本框。**

这样既能让画面扩展成 16:9、按每页留白重新构图，又不会让图像模型生成乱码文字。
**最高优先级版式规则：整张无字底图必须铺满 16:9，不要为了放字把图缩小、裁切或改成图文分栏；
文字可以直接叠在图上，但只能落在空白纸面/背景色上，不得覆盖角色、人脸、身体、关键道具或主体轮廓。
不要把文字塞进独立白色侧栏或另铺一块大底板来“腾位置”；找不到干净空白时，先缩字号、换位置或重生成底图。
宁可缩小文字、换位置或重生成底图，也不能把字压在人物画像上。**
《胖胖兔减肥》11 页实测通过：Qwen-Image 2.1 Turbo 出底图，Pillow 用
`C:/Windows/Fonts/simkai.ttf` 叠字，PPTX 文本框用楷体，中文/阿拉伯数字/弯引号/标点均正常。

## 上下游

```
绘本原件 / 逐页插画 ──huiben-02-ppt──▶ 逐镜 PPTX ──huiben-01-creater──▶ 绘本短视频 MP4
                       ├─ 原文逐字.md（逐页原文）
                       ├─ slides.json（拆镜与逐字文本）
                       ├─ 01-Comfy底图/（无字 16:9 画面）
                       ├─ 02-文字合成/（叠字 PNG，视觉预览）
                       ├─ 04-验收/shot-duration-plan.json（逐页字数 → 视频镜头时长预算）
                       └─ <书名>.pptx（底图 + 可编辑文字）
```

PPTX 里的底图就是后续主视觉的风格与角色锚点；文字是台本和字幕的唯一来源。
每页文字数量并不相同，必须把逐页字数和镜头时长预算一起交给下游，
避免 `huiben-01-creater` 用等长镜头迫使 TTS 加速。

## 默认流程（优先使用）

### 1. 先对齐三件事

1. **是否允许删改原文**：默认一字不改，含引号、感叹号、叠词、阿拉伯数字。
   要压缩或改写必须先问用户，见坑 3。
2. **拆镜粒度**：默认一个编号/一个场景 = 一页；问答型绘本把同一轮问答合在一页。
3. **底图来源与文字安全区**：默认用本机 ComfyUI 图生图/扩图生成无字 16:9 底图。
   只有用户明确要求「不重画、只排版」时，才允许直接嵌原件裁剪图；必须在交付说明里写明「插画含原件文字」。
   在出图前就要指定「文字安全区」：优先利用页内天然空白纸面，可以是上方、下方、左右边缘或窄栏；
   安全区内不得出现角色、人脸、身体、关键道具或主体轮廓。整图仍铺满页面，文字叠在空白处即可。
   默认不要做「左文右图」「上图下文」这种硬分栏；只有用户明确要求分栏时才分栏。

### 2. 逐字转写与拆镜

1. 每张扫描页先看图再抄字；OCR 或视觉模型只当草稿。
2. 有编号块按编号排序，不按视觉顺序；忽略页码和出版社信息。
3. 页边截断字按上下文补全，并在 `原文逐字.md` 文末列补全清单。
4. 产出 `slides.json`。每页至少有 `id`、`kind`、`image`、`text`、`textSafeZone`（封面用 `title`）。
   文本的逐字来源是这个文件，后续 Python/PPTX 都从它读取，不再手工抄一遍。
   `textSafeZone` 用 `[left, top, right, bottom]` 像素坐标写明可放文字的空白区域；
   没有通过安全区检查的页不得进入批量组装。

### 3. 用本机 ComfyUI 生成无字底图

默认配置见 `scripts/comfy_ppt/example-config.json`，可复用入口是
`scripts/comfy_ppt/build_comfy_ppt.py` + `scripts/comfy_ppt/comfy_client.py`。

硬规则：

- 先检查 `http://127.0.0.1:8188/system_stats` 与 `/queue`；队列为空再提交。
- 模型优先用已验证的 Qwen-Image 2.1 Turbo 编辑工作流：
  `UNETLoader → QwenImage21Cache → TextEncodeQwenImage21 → CFGGuider → SamplerCustomAdvanced → VAEDecode → SaveImage`。
  参考图挂 `TextEncodeQwenImage21.inputs["images.image_1"]`；16:9 输出画布用
  `EmptyImage(1344×768) → VAEEncode → SamplerCustomAdvanced.latent_image`。
- 正向提示词必须逐页指定「主体集中在哪一侧、哪一块保留完全干净的纸面」，
  并明确「文字安全区内不得出现角色、人脸、身体、关键道具或主体轮廓」，
  再接一段全片复用的画风锚点和角色恒定描述。
- 负向提示词必须明确禁止 `Chinese text, words, letters, numbers, captions, labels,
  watermarks, page numbers, speech bubbles, typography`。
- 串行生成，不要一次把全书全塞进队列。每页结束后看联系表；
  扩图把别的镜头角色/球/山坡混进留白区时，只重跑该页并改写「删除其他元素」提示词。
- 已验收的底图可以复用；在报告里标明 `reusedValidatedRaw`。

### 4. Python 叠字 + 组装 PPTX

- **保留整图铺满，不要把底图缩小或切成图文两栏。** 文字直接叠在图上，允许压纸纹、留白或背景色；
  但不得压角色画像、人脸、身体、关键道具或主体轮廓。能不铺底板就不铺，底板不能用来遮盖角色。
- 用 Pillow 在 1344×768 的无字底图上叠字，字体用 `C:/Windows/Fonts/simkai.ttf`；
  中文文本框按字数自动缩字号，允许自动换行，但拼接后必须与 `slides.json` 原文逐字一致。
- 正文文字必须完整位于 `textSafeZone`；可编辑文本框和安全区底板都不能越界。
  安全区不足时优先换位置、缩字号或增加换行；仍不足则重生成底图，把主体收拢并腾出空白，
  不要自动改成左文右图/上下分栏，也不要裁切角色。禁止用「文字侧固定为左/右」的批量模板硬套所有页。
- PNG 只用于视觉预览与联系表；PPTX 必须同时插入同一张无字底图，并用
  `python-pptx` 放可编辑文本框，字体同时设置 latin / ea / cs 为楷体。
- PPTX 插入的底图必须是原始整图版本，不能插入缩小、裁切或分栏后的安全区版本。
- 默认命令：

```powershell
$env:PYTHONUTF8 = "1"
$env:PYTHONIOENCODING = "utf-8"
python <skill>/scripts/comfy_ppt/build_comfy_ppt.py <项目>/comfy-config.json
python <skill>/scripts/comfy_ppt/build_comfy_ppt.py <项目>/comfy-config.json --final-only
```

第二条只重排文字和 PPTX，不触发 GPU；改字号、修文字侧、调整底板时优先用它。

### 4b. 产出逐页视频时长预算

PPT 完成后运行：

```powershell
python <skill>/scripts/comfy_ppt/plan_shot_durations.py `
  <项目>/03-转写与拆镜/slides.json `
  --out <项目>/04-验收/shot-duration-plan.json
```

脚本逐页统计朗读字符、行数、对白切换，给出 `estimatedTtsSec`、
`plannedSlotSec` 和符合 H3 `17n+5` 规则的 `plannedFrames`；默认自然语速预算
**3.5 字/秒（210 字/分钟）**，`maxNaturalAtempo = 1.15`。这是低幼绘本旁白的
讲故事速度，不是日常对话速度。长页会得到更长镜头，短页得到更短镜头。
超过单镜 12s 容量的页标记 `needsSplit: true`，由视频阶段先拆镜或经用户确认
允许超长镜头；脚本不会默默把语速压快。
这份规划是 `huiben-01-creater` 阶段 0 的必读输入；不要只交 PPTX 不交字数分布。

### 5. 人眼复核 + 自检

先看 `11页图文联系表.png`：文字是否压住角色、文字安全区是否干净、整图是否仍铺满、
角色是否走形、风格是否一致、底图是否有残留字或水印。再看 `11页Comfy无字底图联系表.png`。
必须逐页放大检查文字框边界：安全区内不能有角色、人脸、身体、关键道具或主体轮廓。

自检门槛：

| 项 | 方法 | 门槛 |
|---|---|---|
| 页数与镜号 | python-pptx | `== slides.json` 条数 |
| 每页底图 | 遍历 shape 统计 picture | 恰好 1 张 |
| 文字逐字 | 抽取 PPTX 文本与 `slides.json` 比对 | 拼接后逐字一致，含标点/数字/引号 |
| PNG 文字 | 报告里的 `renderedLines` 拼接 | 与原文一致 |
| 图片尺寸 | Pillow | 底图与合成图均 1344×768 |
| 整图铺满 | python-pptx / 联系表 | 底图 1344×768 铺满整页；无缩小、裁切或强制分栏 |
| 文字安全区 | 逐页像素诊断 + 放大人眼 | 文字框 100% 落在安全区；0 压角色、0 压主体轮廓 |
| 版式 | 联系表人眼 | 0 乱码、0 重复文字、整图铺满、字不压角色 |
| 插画无文字 | 联系表 + 放大抽查 | 0 残留 |
| 逐页时长预算 | `shot-duration-plan.json` | 页数 == slides.json；长页 `plannedFrames` > 短页；`check.passed == true` |

以上由 `build_comfy_ppt.py` 输出 `04-验收/全本自检.json`，`allPassed` 必须为 `true`。

## 产出目录

```
<项目>/
├── 03-转写与拆镜/原文逐字.md
├── 03-转写与拆镜/slides.json
├── 03-工作流与脚本/comfy-config.json
├── 01-Comfy底图/S01-comfy底图.png …
├── 02-文字合成/S01-图文.png …
├── 04-验收/11页图文联系表.png
├── 04-验收/11页Comfy无字底图联系表.png
├── 04-验收/全本生成报告.json
├── 04-验收/shot-duration-plan.json
├── 04-验收/全本自检.json
└── <书名>.pptx
```

## 兼容旧项目

`scripts/build_pptx.py`、`scripts/contact_sheet.py`、`scripts/selftest.py` 和
`references/layout-spec.md` 保留给《小花籽找快乐》那种 1536×1024 3:2 插画项目。
旧脚本默认把 3:2 插画按高度铺满 16:9 页面，不适合作为新项目默认流程；
新项目一律用 `scripts/comfy_ppt/` 的 16:9 全幅底图 + 叠字流程。

## 坑（都踩过）

1. **跨页扫描的文字块编号与视觉顺序无关**：按编号排。
2. **页边截断字**：补全并记录清单；不补会让下游字幕缺字。
3. **改写原文会产生断句坑**：「把雪山当作自己的家」缩成「把雪山当作家」后，
   TTS 读成「当作家」。默认逐字照抄；要缩写先问用户。
4. **不要让图像模型画中文**：Qwen/图像编辑容易把字画成乱码、残字或多笔画。
   ComfyUI 只负责无字画面，文字永远交给 Python。
5. **扩图会把别的镜头元素补进留白**：提示词要写「只保留本页这一组主体，删除其他角色/球/山/植物」，
   发现乱入只重跑该页。
6. **文字可以压图，但不得压角色**：整图铺满页面，文字叠在天然空白纸面上。先做原画索引，
   再按每页角色位置决定 `textSafeZone`，不是简单按左/右固定文字侧。文字、文本框和任何底板
   都不得压到角色、人脸、身体、关键道具或主体轮廓；不要用缩小整图、裁切或图文分栏来回避。
   发现覆盖时，换位置、缩字号或重生成底图，不要硬挪文本框、不要用底栏遮住角色。
7. **PPTX 文字仍要可编辑**：PNG 只是预览；交付的 PPTX 必须有独立文本框。
8. **楷体要同时设 latin / ea / cs**，否则数字和标点会掉成默认字体。
9. **联系表必须看**：单张图看不出风格漂移，11 页拼一起立刻现形。
10. **批量前先单页试跑**：确认 ComfyUI 工作流、参考图挂载、16:9 全幅画布、字体、
    文字安全区和「不压角色」验收图，再跑全本。
11. **默认禁止图文分栏救场**：整图铺满优先。不要自动把画面缩小放到一侧、再加白栏放字；
    除非用户明确要求分栏，否则用页内空白、换行、缩字号或重生成底图解决。
12. **每页字数不一样，视频镜头不能等长**：PPT 交付必须带 `shot-duration-plan.json`。
    下游若按固定 5–7 秒/镜生成，长页 TTS 会被迫加速到 1.35–1.45 倍；
    让下游先读字数预算，按长页生成更长 H3 镜头，再回填实测音频时长。

## 交付

1. `<书名>.pptx`（每页无字底图 + 可编辑原文文本框）
2. `原文逐字.md`（含补全清单）
3. `slides.json`
4. 11 页图文联系表与无字底图联系表
5. `全本生成报告.json`（Comfy prompt / seed / 输出 / 复用页）
6. `shot-duration-plan.json`（逐页字数、计划秒数、17n+5 帧数、语速门禁）
7. `全本自检.json`（`allPassed: true`）
8. 一句话验收结论：哪些页文字补全、哪些底图复用、是否有人眼复核、最长/最短镜头
