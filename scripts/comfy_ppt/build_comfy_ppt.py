# -*- coding: utf-8 -*-
"""复用 ComfyUI 底图并生成《胖胖兔减肥》11 页可编辑 PPT。"""
import argparse
import json
import math
import shutil
import time
import uuid
from pathlib import Path

import numpy as np
from PIL import Image, ImageDraw, ImageFilter, ImageFont
from pptx import Presentation
from pptx.enum.text import PP_ALIGN
from pptx.util import Emu, Pt

import comfy_client as helper


SCRIPT_DIR = Path(__file__).resolve().parent
EMU_W, EMU_H = 12191695, 6858000


def load_pages(config):
    slides_path = (SCRIPT_DIR / config["slidesJson"]).resolve()
    slides = json.loads(slides_path.read_text(encoding="utf-8"))["slides"]
    layouts = {item["id"]: item for item in config["pageLayouts"]}
    source_root = (SCRIPT_DIR / config["sourceRoot"]).resolve()
    pages = []
    for slide in slides:
        page = dict(layouts[slide["id"]])
        page["id"] = slide["id"]
        page["kind"] = slide.get("kind", "content")
        page["ref"] = source_root / "02-插画" / Path(slide["image"]).name
        page["text"] = [slide["title"]] if page["kind"] == "cover" else list(slide["text"])
        page["textSafeZone"] = slide.get("textSafeZone") or page.get("textSafeZone")
        pages.append(page)
    return pages


def detail_coverage(image, safe_zone, config):
    pixels = np.asarray(image.convert("RGB"), dtype=np.int16)
    luminance = pixels.mean(axis=2)
    chroma = pixels.max(axis=2) - pixels.min(axis=2)
    gx = np.abs(np.diff(luminance, axis=1, prepend=luminance[:, :1]))
    gy = np.abs(np.diff(luminance, axis=0, prepend=luminance[:1, :]))
    mask = (
        (chroma > config.get("safeChroma", 38))
        | (np.maximum(gx, gy) > config.get("safeEdge", 18))
    ).astype(np.uint8) * 255
    dilated = Image.fromarray(mask).filter(ImageFilter.MaxFilter(5))
    left, top, right, bottom = safe_zone
    region = np.asarray(dilated.crop((left, top, right, bottom)), dtype=np.uint8)
    return float((region > 0).mean())


def create_safe_zone_canvas(raw_path, page, config, destination):
    image = Image.open(raw_path).convert("RGB")
    if image.size != (config["width"], config["height"]):
        image = image.resize((config["width"], config["height"]), Image.Resampling.LANCZOS)
    if not page["text"]:
        image.save(destination, "PNG")
        return image, [0, 0, image.width, image.height], {"layoutMode": "full-bleed-no-text"}
    safe_zone = page.get("textSafeZone")
    if not safe_zone:
        raise RuntimeError(f"textSafeZone required for full-bleed text: {page['id']}")
    image.save(destination, "PNG")
    detection = {
        "layoutMode": "full-bleed-overlay",
        "fullBleedImage": str(raw_path),
        "textPanel": safe_zone,
        "fullArtVisible": True,
        "textPanelOverIllustration": False,
    }
    return image, [int(value) for value in safe_zone], detection


def fit_text(page, config, safe_zone):
    opening = "“‘（《【〈"
    closing = "，。！？；：、”’）》】〉"
    left, top, right, bottom = safe_zone
    padding = config.get("safeZonePadding", 28)
    max_width = right - left - 2 * padding
    max_height = bottom - top - 2 * padding
    requested = int(page["fontPx"])
    minimum = config.get("minFontPx", 16) if page["kind"] == "content" else 36
    dummy = ImageDraw.Draw(Image.new("RGB", (1, 1)))
    for font_px in range(requested, minimum - 1, -1):
        font = ImageFont.truetype(config["font"], font_px)
        rendered = []
        for line in page["text"]:
            current = ""
            for char in line:
                candidate = current + char
                if current and dummy.textlength(candidate, font=font) > max_width:
                    if char in closing:
                        current = candidate
                        continue
                    if current[-1] in opening:
                        rendered.append(current[:-1])
                        current = current[-1] + char
                        continue
                    rendered.append(current)
                    current = char
                else:
                    current = candidate
            rendered.append(current)
        line_height = int(font_px * config["lineSpacing"])
        total_height = line_height * len(rendered)
        if total_height <= max_height and all(dummy.textlength(line, font=font) <= max_width for line in rendered):
            return font_px, rendered, line_height, total_height, max_width, left + padding, top + (max_height - total_height) // 2
    raise RuntimeError(f"Text does not fit: {page['id']}")


def add_text(safe_image, page, config, destination, layout, safe_zone):
    image = safe_image.convert("RGB")
    font_px, lines, line_height, total_height, box_width, box_x, top = layout
    draw = ImageDraw.Draw(image)
    font = ImageFont.truetype(config["font"], font_px)
    color = tuple(int(page["color"][index:index + 2], 16) for index in (0, 2, 4))
    for index, line in enumerate(lines):
        draw.text(
            (box_x, top + index * line_height),
            line,
            font=font,
            fill=color,
            stroke_width=2,
            stroke_fill=(255, 255, 255),
        )
    image.save(destination, "PNG")
    left, zone_top, right, bottom = safe_zone
    return {
        "fontPx": font_px,
        "lines": lines,
        "lineHeight": line_height,
        "totalHeight": total_height,
        "boxX": box_x,
        "boxWidth": box_width,
        "top": top,
        "textExactConcat": "".join(lines) == "".join(page["text"]),
        "textSafeZone": safe_zone,
        "textBox": [box_x, top, box_x + box_width, top + total_height],
        "textPanelOverIllustration": False,
        "textBoxInsideSafeZone": (
            box_x >= left
            and top >= zone_top
            and box_x + box_width <= right
            and top + total_height <= bottom
        ),
    }


def make_grid_contact_sheet(paths, page_ids, destination, font_path):
    thumb_width, thumb_height, columns = 448, 256, 3
    rows = math.ceil(len(paths) / columns)
    sheet = Image.new("RGB", (columns * thumb_width, rows * thumb_height), (245, 242, 236))
    draw = ImageDraw.Draw(sheet)
    label_font = ImageFont.truetype(font_path, 18)
    for index, path in enumerate(paths):
        thumb = Image.open(path).convert("RGB").resize(
            (thumb_width, thumb_height), Image.Resampling.LANCZOS
        )
        left = (index % columns) * thumb_width
        top = (index // columns) * thumb_height
        sheet.paste(thumb, (left, top))
        draw.rectangle((left + 8, top + 8, left + 70, top + 32), fill=(255, 253, 247))
        draw.text((left + 15, top + 8), page_ids[index], font=label_font, fill=(70, 55, 42))
    destination.parent.mkdir(parents=True, exist_ok=True)
    sheet.save(destination, "PNG")


def make_pptx(pages, raw_paths, layouts, config, destination):
    presentation = Presentation()
    presentation.slide_width = Emu(EMU_W)
    presentation.slide_height = Emu(EMU_H)
    blank = presentation.slide_layouts[6]
    pixels_per_inch = config["width"] / (EMU_W / 914400)
    font_scale = 72 / pixels_per_inch
    for page, raw_path in zip(pages, raw_paths):
        layout = layouts[page["id"]]
        slide = presentation.slides.add_slide(blank)
        slide.shapes.add_picture(str(raw_path), 0, 0, Emu(EMU_W), Emu(EMU_H))
        textbox = slide.shapes.add_textbox(
            Emu(int(layout["boxX"] * EMU_W / config["width"])),
            Emu(int(layout["top"] * EMU_H / config["height"])),
            Emu(int(layout["boxWidth"] * EMU_W / config["width"])),
            Emu(int(layout["totalHeight"] * 1.25 * EMU_H / config["height"])),
        )
        frame = textbox.text_frame
        frame.word_wrap = False
        frame.margin_left = 0
        frame.margin_right = 0
        frame.margin_top = 0
        frame.margin_bottom = 0
        for index, line in enumerate(layout["lines"]):
            paragraph = frame.paragraphs[0] if index == 0 else frame.add_paragraph()
            paragraph.alignment = PP_ALIGN.LEFT
            paragraph.line_spacing = config["lineSpacing"]
            paragraph.space_before = Pt(0)
            paragraph.space_after = Pt(0)
            run = paragraph.add_run()
            run.text = line
            helper.set_pptx_font(run, config["fontName"], layout["fontPx"] * font_scale, page["color"])
    presentation.save(destination)


def verify_output(pptx_path, pages, raw_paths, final_paths, layouts, config, report):
    presentation = Presentation(pptx_path)
    if len(presentation.slides) != len(pages):
        raise RuntimeError("PPTX page count mismatch")
    checks = []
    for index, page in enumerate(pages):
        slide = presentation.slides[index]
        pictures = [shape for shape in slide.shapes if shape.shape_type == 13]
        runs = [
            run.text
            for shape in slide.shapes
            if shape.has_text_frame
            for paragraph in shape.text_frame.paragraphs
            for run in paragraph.runs
        ]
        raw = Image.open(raw_paths[index])
        final = Image.open(final_paths[index])
        checks.append(
            {
                "id": page["id"],
                "pictures": len(pictures),
                "textExactConcat": "".join(runs) == "".join(page["text"]),
                "pngTextExactConcat": layouts[page["id"]]["textExactConcat"],
                "rawSize": list(raw.size),
                "finalSize": list(final.size),
                "fontPx": layouts[page["id"]]["fontPx"],
                "renderedLines": layouts[page["id"]]["lines"],
                "safeZoneDetailCoverage": layouts[page["id"]]["safeZoneDetailCoverage"],
            }
        )
    return {
        "pptx": str(pptx_path),
        "slideCount": len(presentation.slides),
        "slideSizeEmu": [presentation.slide_width, presentation.slide_height],
        "textRenderer": "Pillow + simkai.ttf; PPTX uses editable 楷体 text boxes",
        "checks": checks,
        "allPassed": all(
            item["pictures"] == 1
            and item["textExactConcat"]
            and item["pngTextExactConcat"]
            and layouts[item["id"]]["textBoxInsideSafeZone"]
            and not layouts[item["id"]]["textPanelOverIllustration"]
            and (
                not pages[index]["text"]
                or item["safeZoneDetailCoverage"] <= config.get("safeZoneMaxDetailCoverage", 0.02)
            )
            and item["rawSize"] == [config["width"], config["height"]]
            and item["finalSize"] == [config["width"], config["height"]]
            for item in checks
        ),
        "comfyRuns": sum("promptId" in entry for entry in report["pages"]),
    }


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("config")
    parser.add_argument("--only", action="append", default=[], help="只处理指定页号，可重复")
    parser.add_argument("--force", action="store_true", help="强制重跑选中页的 ComfyUI 底图")
    parser.add_argument("--final-only", action="store_true", help="只用现有底图重排")
    args = parser.parse_args()

    config_path = Path(args.config).resolve()
    config = json.loads(config_path.read_text(encoding="utf-8"))
    pages = load_pages(config)
    selected = set(args.only)
    if selected:
        pages = [page for page in pages if page["id"] in selected]
    root = (SCRIPT_DIR / config["outputRoot"]).resolve()
    validated_raw_root = (SCRIPT_DIR / config["validatedRawRoot"]).resolve()
    raw_dir = root / "01-Comfy底图"
    text_dir = root / "02-文字合成"
    verify_dir = root / "04-验收"
    for directory in (raw_dir, text_dir, verify_dir):
        directory.mkdir(parents=True, exist_ok=True)

    report = {
        "comfyui": helper.COMFY,
        "model": "Qwen-Image 2.1 Turbo (local ComfyUI /prompt)",
        "sourceSlides": str((SCRIPT_DIR / config["slidesJson"]).resolve()),
        "reusedValidatedRaw": [],
        "pages": [],
    }
    raw_paths = []
    final_paths = []
    layouts = {}
    for index, page in enumerate(pages):
        raw_path = raw_dir / f"{page['id']}-comfy底图.png"
        final_path = text_dir / f"{page['id']}-图文.png"
        page_report = {
            "id": page["id"],
            "ref": str(page["ref"]),
            "raw": str(raw_path),
            "final": str(final_path),
            "text": page["text"],
            "textSide": page["textSide"],
        }
        needs_comfy = not raw_path.exists() or args.force
        if page.get("useValidatedRaw") and needs_comfy:
            validated = validated_raw_root / f"{page['id']}-comfy底图.png"
            shutil.copy2(validated, raw_path)
            report["reusedValidatedRaw"].append(page["id"])
            page_report["source"] = "validated-three-page-test"
            needs_comfy = False
        if needs_comfy and not args.final_only:
            image_name = helper.upload_image(page["ref"])
            workflow = helper.build_workflow(page, config, image_name, config["seedBase"] + index)
            queued = helper.post_json(
                "/prompt",
                {"prompt": workflow, "client_id": f"pangpangtu-full-{uuid.uuid4().hex}"},
            )
            started = time.time()
            entry = helper.wait_for_prompt(queued["prompt_id"])
            image_info = helper.find_output_image(entry, page["id"])
            raw_path.write_bytes(helper.fetch_image(image_info))
            page_report.update(
                {
                    "promptId": queued["prompt_id"],
                    "seed": config["seedBase"] + index,
                    "elapsedSec": round(time.time() - started, 1),
                    "outputFilename": image_info["filename"],
                    "status": entry.get("status", {}).get("status_str"),
                    "source": "local-comfyui",
                }
            )
        if not raw_path.exists():
            raise FileNotFoundError(raw_path)
        canvas_path = raw_dir / f"{page['id']}-安全区底图.png"
        safe_image, safe_zone, detection = create_safe_zone_canvas(raw_path, page, config, canvas_path)
        coverage = detail_coverage(safe_image, safe_zone, config)
        if page["text"] and coverage > config.get("safeZoneMaxDetailCoverage", 0.02):
            raise RuntimeError(f"Text safe zone covers details: {page['id']} {coverage:.4f}")
        layout = fit_text(page, config, safe_zone)
        layouts[page["id"]] = add_text(safe_image, page, config, final_path, layout, safe_zone)
        layouts[page["id"]]["safeZoneDetailCoverage"] = round(coverage, 5)
        page_report["safeCanvas"] = str(canvas_path)
        page_report["safeZoneDetection"] = detection
        page_report.update(layouts[page["id"]])
        report["pages"].append(page_report)
        raw_paths.append(raw_path)
        final_paths.append(final_path)
        print(f"{page['id']} done: {final_path}")

    pptx_path = root / f"{config['title']}.pptx"
    make_pptx(pages, raw_paths, layouts, config, pptx_path)
    contact_path = verify_dir / "11页图文联系表.png"
    raw_contact_path = verify_dir / "11页Comfy无字底图联系表.png"
    make_grid_contact_sheet(final_paths, [page["id"] for page in pages], contact_path, config["labelFont"])
    make_grid_contact_sheet(raw_paths, [page["id"] for page in pages], raw_contact_path, config["labelFont"])
    report_path = verify_dir / "全本生成报告.json"
    report_path.write_text(json.dumps(report, ensure_ascii=False, indent=2), encoding="utf-8")
    verification = verify_output(pptx_path, pages, raw_paths, final_paths, layouts, config, report)
    verification_path = verify_dir / "全本自检.json"
    verification_path.write_text(json.dumps(verification, ensure_ascii=False, indent=2), encoding="utf-8")
    note_path = root / "全本说明.md"
    note_path.write_text(
        "\n".join(
            [
                "# 《胖胖兔减肥》ComfyUI 全本 PPT",
                "",
                "- 11 页，16:9，1344×768。",
                "- 底图由本机 ComfyUI Qwen-Image 2.1 Turbo 生成；S01/S02/S07 复用已验收测试底图。",
                "- 中文、数字、标点全部由 Python/Pillow 叠加，PPTX 内为可编辑楷体文本框。",
                "- 逐字自检见 `04-验收/全本自检.json`，运行记录见 `04-验收/全本生成报告.json`。",
                "",
                "复现：",
                "",
                "```powershell",
                '$env:PYTHONUTF8 = "1"',
                '$env:PYTHONIOENCODING = "utf-8"',
                'python "<skill>/scripts/comfy_ppt/build_comfy_ppt.py" "<项目>/comfy-config.json" --final-only',
                "```",
                "",
            ]
        ),
        encoding="utf-8",
    )
    print(f"pptx: {pptx_path}")
    print(f"contact sheet: {contact_path}")
    print(f"verification: {verification_path}")
    print(f"allPassed: {verification['allPassed']}")


if __name__ == "__main__":
    main()
