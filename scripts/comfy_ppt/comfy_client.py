# -*- coding: utf-8 -*-
"""用本机 ComfyUI 生成无字底图，再用 PIL 叠加中文原文。"""
import argparse
import io
import json
import time
import urllib.parse
import urllib.request
import uuid
from pathlib import Path

from PIL import Image, ImageDraw, ImageFont
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.oxml.ns import qn
from pptx.util import Emu, Pt


COMFY = "http://127.0.0.1:8188"
SIGMAS = "1.0, 0.978453, 0.95418, 0.926626, 0.89508, 0.845148, 0.704534, 0.414568, 0"
SCRIPT_DIR = Path(__file__).resolve().parent
EMU_W, EMU_H = 12191695, 6858000


def post_json(path, payload):
    data = json.dumps(payload, ensure_ascii=False).encode("utf-8")
    request = urllib.request.Request(
        COMFY + path,
        data=data,
        headers={"Content-Type": "application/json"},
        method="POST",
    )
    with urllib.request.urlopen(request, timeout=30) as response:
        return json.load(response)


def get_json(path):
    with urllib.request.urlopen(COMFY + path, timeout=30) as response:
        return json.load(response)


def upload_image(path):
    boundary = "----PangpangtuComfy" + uuid.uuid4().hex
    filename = path.name
    content = path.read_bytes()
    body = io.BytesIO()
    body.write(f"--{boundary}\r\n".encode())
    body.write(
        f'Content-Disposition: form-data; name="image"; filename="{filename}"\r\n'.encode()
    )
    body.write(b"Content-Type: image/png\r\n\r\n")
    body.write(content)
    body.write(f"\r\n--{boundary}--\r\n".encode())
    request = urllib.request.Request(
        COMFY + "/upload/image",
        data=body.getvalue(),
        headers={"Content-Type": f"multipart/form-data; boundary={boundary}"},
        method="POST",
    )
    with urllib.request.urlopen(request, timeout=60) as response:
        result = json.load(response)
    return result["name"]


def build_workflow(page, config, image_name, seed):
    width, height = config["width"], config["height"]
    prompt = page["prompt"] + config["styleTail"]
    return {
        "1": {
            "class_type": "UNETLoader",
            "inputs": {
                "unet_name": "qwen_image_2.1_turbo_int8_convrot.safetensors",
                "weight_dtype": "default",
            },
        },
        "2": {
            "class_type": "CLIPLoader",
            "inputs": {
                "clip_name": "qwen3vl_8b_int8_convrot.safetensors",
                "type": "qwen_image",
                "device": "default",
            },
        },
        "3": {
            "class_type": "VAELoader",
            "inputs": {"vae_name": "qwen_image_2.1_vae_bf16.safetensors"},
        },
        "4": {
            "class_type": "QwenImage21Cache",
            "inputs": {"model": ["1", 0], "device": "auto", "dtype": "default"},
        },
        "5": {
            "class_type": "TextEncodeQwenImage21",
            "inputs": {
                "clip": ["2", 0],
                "prompt": prompt,
                "negative_prompt": config["negative"],
                "vae": ["3", 0],
                "resolution": config["resolution"],
                "images.image_1": ["15", 0],
            },
        },
        "6": {
            "class_type": "CFGGuider",
            "inputs": {
                "model": ["4", 0],
                "positive": ["5", 0],
                "negative": ["5", 1],
                "cfg": 1,
            },
        },
        "7": {
            "class_type": "RandomNoise",
            "inputs": {"noise_seed": seed},
        },
        "8": {
            "class_type": "KSamplerSelect",
            "inputs": {"sampler_name": "euler"},
        },
        "9": {
            "class_type": "ManualSigmas",
            "inputs": {"sigmas": SIGMAS},
        },
        "10": {
            "class_type": "SamplerCustomAdvanced",
            "inputs": {
                "noise": ["7", 0],
                "guider": ["6", 0],
                "sampler": ["8", 0],
                "sigmas": ["9", 0],
                "latent_image": ["14", 0],
            },
        },
        "11": {
            "class_type": "VAEDecode",
            "inputs": {"samples": ["10", 0], "vae": ["3", 0]},
        },
        "12": {
            "class_type": "SaveImage",
            "inputs": {
                "images": ["11", 0],
                "filename_prefix": f"PangpangtuTest/{page['id']}",
            },
        },
        "13": {
            "class_type": "EmptyImage",
            "inputs": {
                "width": width,
                "height": height,
                "batch_size": 1,
                "color": 0,
            },
        },
        "14": {
            "class_type": "VAEEncode",
            "inputs": {"pixels": ["13", 0], "vae": ["3", 0]},
        },
        "15": {
            "class_type": "LoadImage",
            "inputs": {"image": image_name},
        },
    }


def wait_for_prompt(prompt_id, timeout=300):
    deadline = time.time() + timeout
    while time.time() < deadline:
        history = get_json(f"/history/{prompt_id}")
        if prompt_id in history:
            entry = history[prompt_id]
            status = entry.get("status", {})
            if status.get("completed") or status.get("status_str") == "success":
                return entry
            if status.get("status_str") in {"error", "failed"}:
                raise RuntimeError(json.dumps(status, ensure_ascii=False))
        time.sleep(2)
    raise TimeoutError(f"ComfyUI prompt timeout: {prompt_id}")


def find_output_image(entry, page_id):
    for node_output in entry.get("outputs", {}).values():
        for image in node_output.get("images", []):
            if image.get("filename", "").startswith(page_id):
                return image
    raise RuntimeError(f"No output image for {page_id}: {json.dumps(entry, ensure_ascii=False)[:1000]}")


def fetch_image(image_info):
    query = urllib.parse.urlencode(
        {
            "filename": image_info["filename"],
            "subfolder": image_info.get("subfolder", ""),
            "type": image_info.get("type", "output"),
        }
    )
    with urllib.request.urlopen(COMFY + "/view?" + query, timeout=60) as response:
        return response.read()


def wrap_lines(lines, font, max_width, draw):
    opening = "“‘（《【〈"
    closing = "，。！？；：、”’）》】〉"
    result = []
    for line in lines:
        current = ""
        for char in line:
            candidate = current + char
            if current and draw.textlength(candidate, font=font) > max_width:
                if char in closing:
                    current = candidate
                    continue
                if current[-1] in opening:
                    result.append(current[:-1])
                    current = current[-1] + char
                    continue
                result.append(current)
                current = char
            else:
                current = candidate
        result.append(current)
    return result


def add_text(source, page, config, destination):
    image = Image.open(source).convert("RGB")
    if image.size != (config["width"], config["height"]):
        image = image.resize((config["width"], config["height"]), Image.Resampling.LANCZOS)
    draw = ImageDraw.Draw(image)
    font_size = int(page["fontPt"] * config["width"] / 1344)
    font = ImageFont.truetype(config["font"], font_size)
    margin = 78
    side = page["textSide"]
    box_width = 520 if side == "right" else 560
    box_x = config["width"] - box_width - margin if side == "right" else margin
    lines = wrap_lines(page["text"], font, box_width, draw)
    line_height = int(font_size * 1.62)
    total_height = line_height * len(lines)
    if page["kind"] == "cover":
        top = (config["height"] - total_height) // 2
    else:
        top = max(110, (config["height"] - total_height) // 2)
    color = tuple(int(page["color"][index:index + 2], 16) for index in (0, 2, 4))
    for index, line in enumerate(lines):
        y = top + index * line_height
        draw.text(
            (box_x, y),
            line,
            font=font,
            fill=color,
            stroke_width=2,
            stroke_fill=(255, 253, 246),
        )
    image.save(destination, "PNG")
    return lines


def make_contact_sheet(paths, destination):
    thumb_width, thumb_height = 672, 384
    sheet = Image.new("RGB", (thumb_width, thumb_height * len(paths)), (245, 242, 236))
    for index, path in enumerate(paths):
        thumb = Image.open(path).convert("RGB").resize(
            (thumb_width, thumb_height), Image.Resampling.LANCZOS
        )
        sheet.paste(thumb, (0, index * thumb_height))
    sheet.save(destination, "PNG")


def set_pptx_font(run, font_name, size_pt, color_hex):
    run.font.name = font_name
    run.font.size = Pt(size_pt)
    run.font.color.rgb = RGBColor.from_string(color_hex)
    rpr = run._r.get_or_add_rPr()
    for tag in ("a:latin", "a:ea", "a:cs"):
        element = rpr.find(qn(tag))
        if element is None:
            element = rpr.makeelement(qn(tag), {})
            rpr.append(element)
        element.set("typeface", font_name)


def make_pptx(pages, config, destination):
    presentation = Presentation()
    presentation.slide_width = Emu(EMU_W)
    presentation.slide_height = Emu(EMU_H)
    blank = presentation.slide_layouts[6]
    pixels_per_inch = config["width"] / (EMU_W / 914400)
    font_size_scale = 72 / pixels_per_inch
    font_name = "楷体"
    for page, raw_path, lines in pages:
        slide = presentation.slides.add_slide(blank)
        slide.shapes.add_picture(str(raw_path), 0, 0, Emu(EMU_W), Emu(EMU_H))
        margin = 78
        box_width = 560
        box_x = config["width"] - box_width - margin if page["textSide"] == "right" else margin
        line_height = int(page["fontPt"] * 1.62)
        total_height = line_height * len(lines)
        top = (config["height"] - total_height) // 2
        if page["kind"] != "cover":
            top = max(110, top)
        textbox = slide.shapes.add_textbox(
            Emu(int(box_x * EMU_W / config["width"])),
            Emu(int(top * EMU_H / config["height"])),
            Emu(int(box_width * EMU_W / config["width"])),
            Emu(int(total_height * 1.25 * EMU_H / config["height"])),
        )
        frame = textbox.text_frame
        frame.word_wrap = False
        frame.margin_left = 0
        frame.margin_right = 0
        frame.margin_top = 0
        frame.margin_bottom = 0
        for index, line in enumerate(lines):
            paragraph = frame.paragraphs[0] if index == 0 else frame.add_paragraph()
            paragraph.alignment = PP_ALIGN.LEFT
            paragraph.line_spacing = 1.62
            paragraph.space_before = Pt(0)
            paragraph.space_after = Pt(0)
            run = paragraph.add_run()
            run.text = line
            set_pptx_font(run, font_name, page["fontPt"] * font_size_scale, page["color"])
    presentation.save(destination)


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("config")
    parser.add_argument("--skip-comfy", action="store_true")
    parser.add_argument("--only", action="append", default=[], help="只跑指定页号，可重复")
    parser.add_argument("--all-final", action="store_true", help="用现有底图重排全部三页")
    args = parser.parse_args()
    config_path = Path(args.config).resolve()
    config = json.loads(config_path.read_text(encoding="utf-8"))
    root = (SCRIPT_DIR / config["outputRoot"]).resolve()
    source_root = (SCRIPT_DIR / config["sourceRoot"]).resolve()
    raw_dir = root / "01-Comfy底图"
    text_dir = root / "02-文字合成"
    verify_dir = root / "04-验收"
    for directory in (raw_dir, text_dir, verify_dir):
        directory.mkdir(parents=True, exist_ok=True)

    report = {
        "comfyui": COMFY,
        "model": "Qwen-Image 2.1 Turbo (local ComfyUI /prompt)",
        "textRenderer": "Pillow + simkai.ttf",
        "pages": [],
    }
    composed_paths = []
    raw_paths = []
    pptx_pages = []
    selected_ids = set(args.only)
    if args.all_final:
        selected_ids = set()
    for index, page in enumerate(config["pages"]):
        if selected_ids and page["id"] not in selected_ids:
            continue
        source_path = (source_root / page["ref"]).resolve()
        raw_path = raw_dir / f"{page['id']}-comfy底图.png"
        final_path = text_dir / f"{page['id']}-图文测试.png"
        page_report = {
            "id": page["id"],
            "ref": str(source_path),
            "raw": str(raw_path),
            "final": str(final_path),
            "text": page["text"],
        }
        if not args.skip_comfy and not args.all_final:
            image_name = upload_image(source_path)
            workflow = build_workflow(page, config, image_name, config["seedBase"] + index)
            queued = post_json("/prompt", {"prompt": workflow, "client_id": f"pangpangtu-{uuid.uuid4().hex}"})
            started = time.time()
            entry = wait_for_prompt(queued["prompt_id"])
            image_info = find_output_image(entry, page["id"])
            raw_path.write_bytes(fetch_image(image_info))
            page_report.update(
                {
                    "promptId": queued["prompt_id"],
                    "seed": config["seedBase"] + index,
                    "elapsedSec": round(time.time() - started, 1),
                    "outputFilename": image_info["filename"],
                }
            )
        lines = add_text(raw_path, page, config, final_path)
        page_report["renderedLines"] = lines
        page_report["textExact"] = [line.strip() for line in lines] == page["text"] or lines == page["text"]
        report["pages"].append(page_report)
        composed_paths.append(final_path)
        raw_paths.append(raw_path)
        pptx_pages.append((page, raw_path, lines))
        print(f"{page['id']} done: {final_path}")

    contact_path = verify_dir / "三页图文联系表.png"
    make_contact_sheet(composed_paths, contact_path)
    raw_contact_path = verify_dir / "三页Comfy无字底图联系表.png"
    make_contact_sheet(raw_paths, raw_contact_path)
    pptx_path = root / "胖胖兔减肥-Comfy图文测试.pptx"
    make_pptx(pptx_pages, config, pptx_path)
    report_path = verify_dir / "测试报告.json"
    if report_path.exists():
        previous = json.loads(report_path.read_text(encoding="utf-8"))
        merged = {entry["id"]: entry for entry in previous.get("pages", [])}
        for entry in report["pages"]:
            if entry["id"] in merged and "promptId" not in entry and "promptId" in merged[entry["id"]]:
                for key in ("promptId", "seed", "elapsedSec", "outputFilename"):
                    entry[key] = merged[entry["id"]][key]
    report_path.write_text(json.dumps(report, ensure_ascii=False, indent=2), encoding="utf-8")
    print(f"contact sheet: {contact_path}")
    print(f"raw contact sheet: {raw_contact_path}")
    print(f"pptx: {pptx_path}")
    print(f"report: {report_path}")


if __name__ == "__main__":
    main()
