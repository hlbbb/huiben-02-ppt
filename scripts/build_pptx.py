# -*- coding: utf-8 -*-
"""slides.json -> 逐镜 PPTX（一镜一页：插画 + 原文文本框）。

版式默认值来自《小花籽找快乐》实测：
  页面 12191695x6858000 EMU（16:9），内容页插画 10287000x6858000 按高铺满水平居中，
  封面插画 7818120x5212080 居中，标题楷体 54pt #D84A38，正文楷体 24pt #4A3828。

用法：
    python build_pptx.py slides.json
    python build_pptx.py slides.json --out 自定义.pptx

slides.json 最小结构：
    {
      "title": "小花籽找快乐",
      "output": "小花籽找快乐.pptx",
      "slides": [
        {"id": "S01", "kind": "cover", "image": "02-插画/S01.png", "title": "小花籽找快乐"},
        {"id": "S02", "kind": "content", "image": "02-插画/S02.png",
         "text": ["第一行", "第二行"]}
      ]
    }
可选字段：slideSizeEmu、imageBoxEmu（content/cover 各自）、textBoxEmu、
font、coverFontPt、fontPt、titleColor、textColor、titleBoxEmu。
"""
import argparse
import json
import os
import sys

from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import PP_ALIGN
from pptx.oxml.ns import qn
from pptx.util import Emu, Pt

EMU_W, EMU_H = 12191695, 6858000
COVER_IMG = (2185416, 1463040, 7818120, 5212080)
COVER_TITLE = (2441448, 274320, 7315200, 1188720)
CONTENT_IMG = (952347, 0, 10287000, 6858000)
CONTENT_TEXT = (7635240, 2377440, 3291840, 4297680)
FONT = "楷体"
TITLE_COLOR = "D84A38"
TEXT_COLOR = "4A3828"


def set_font(run, name, size_pt, color_hex):
    run.font.name = name
    run.font.size = Pt(size_pt)
    run.font.color.rgb = RGBColor.from_string(color_hex)
    rPr = run._r.get_or_add_rPr()
    for tag in ("a:latin", "a:ea", "a:cs"):
        el = rPr.find(qn(tag))
        if el is None:
            el = rPr.makeelement(qn(tag), {})
            rPr.append(el)
        el.set("typeface", name)


def add_textbox(slide, box, lines, font, size_pt, color_hex, align=PP_ALIGN.LEFT):
    x, y, w, h = box
    tb = slide.shapes.add_textbox(Emu(x), Emu(y), Emu(w), Emu(h))
    tf = tb.text_frame
    tf.word_wrap = True
    for i, line in enumerate(lines):
        para = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        para.alignment = align
        run = para.add_run()
        run.text = line
        set_font(run, font, size_pt, color_hex)
    return tb


def add_image(slide, path, box):
    x, y, w, h = box
    return slide.shapes.add_picture(path, Emu(x), Emu(y), Emu(w), Emu(h))


def add_text_panel(slide, box, color="FFFFFF"):
    x, y, w, h = box
    shape = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Emu(x), Emu(y), Emu(w), Emu(h))
    shape.fill.solid()
    shape.fill.fore_color.rgb = RGBColor.from_string(color)
    shape.line.fill.background()
    return shape


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("config")
    ap.add_argument("--out", default=None, help="覆盖 json 里的 output")
    args = ap.parse_args()
    cfg = json.load(open(args.config, encoding="utf-8"))

    base = os.path.dirname(os.path.abspath(args.config))
    out = args.out or cfg.get("output") or "output.pptx"
    if not os.path.isabs(out):
        out = os.path.join(base, out)

    emu_w, emu_h = cfg.get("slideSizeEmu", [EMU_W, EMU_H])
    font = cfg.get("font", FONT)
    title_color = cfg.get("titleColor", TITLE_COLOR)
    text_color = cfg.get("textColor", TEXT_COLOR)
    cover_img = cfg.get("coverImageBoxEmu", list(COVER_IMG))
    cover_title = cfg.get("titleBoxEmu", list(COVER_TITLE))
    content_img = cfg.get("imageBoxEmu", list(CONTENT_IMG))
    content_text = cfg.get("textBoxEmu", list(CONTENT_TEXT))
    cover_pt = cfg.get("coverFontPt", 54)
    body_pt = cfg.get("fontPt", 24)

    prs = Presentation()
    prs.slide_width = Emu(emu_w)
    prs.slide_height = Emu(emu_h)
    blank = prs.slide_layouts[6]

    for spec in cfg["slides"]:
        slide = prs.slides.add_slide(blank)
        image = spec["image"]
        if not os.path.isabs(image):
            image = os.path.join(base, image)
        if not os.path.exists(image):
            sys.exit("缺图: %s (%s)" % (spec.get("id"), image))
        kind = spec.get("kind", "content")
        if kind == "cover":
            add_image(slide, image, spec.get("imageBoxEmu", cover_img))
            add_textbox(slide, spec.get("textBoxEmu", cover_title),
                        [spec.get("title", cfg.get("title", ""))],
                        font, spec.get("fontPt", cover_pt), title_color,
                        align=PP_ALIGN.CENTER)
        else:
            add_image(slide, image, spec.get("imageBoxEmu", content_img))
            lines = spec.get("text") or [spec.get("title", "")]
            panel = spec.get("textPanelEmu") or cfg.get("textPanelEmu")
            if panel:
                add_text_panel(slide, panel, cfg.get("textPanelColor", "FFFFFF"))
            add_textbox(slide, spec.get("textBoxEmu", content_text),
                        lines, font, spec.get("fontPt", body_pt), text_color)

    os.makedirs(os.path.dirname(os.path.abspath(out)) or ".", exist_ok=True)
    prs.save(out)
    print("saved: %s (%d slides)" % (out, len(cfg["slides"])))


if __name__ == "__main__":
    main()
