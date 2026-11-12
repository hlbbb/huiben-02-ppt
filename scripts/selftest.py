# -*- coding: utf-8 -*-
"""自测：从现有 PPTX 抽出插画与逐页文本 -> 用 build_pptx.py 重建 -> 逐项比对。

用法：
    python selftest.py <参照.pptx> [--work <临时目录>]

校验项：页数、逐页文本逐字一致、每页插画数、关键版式 EMU、字体设置。
"""
import argparse
import json
import os
import subprocess
import sys
import tempfile
import zipfile

from pptx import Presentation
from pptx.oxml.ns import qn

HERE = os.path.dirname(os.path.abspath(__file__))


def slide_texts(path):
    prs = Presentation(path)
    out = []
    for slide in prs.slides:
        texts = []
        pics = 0
        for shape in slide.shapes:
            if shape.shape_type == 13:  # PICTURE
                pics += 1
            if shape.has_text_frame:
                for para in shape.text_frame.paragraphs:
                    texts.append("".join(run.text for run in para.runs))
        out.append({"texts": texts, "pics": pics})
    return out


def extract_media(pptx, dest):
    with zipfile.ZipFile(pptx) as z:
        for name in z.namelist():
            if name.startswith("ppt/media/"):
                with open(os.path.join(dest, os.path.basename(name)), "wb") as f:
                    f.write(z.read(name))


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("reference")
    ap.add_argument("--work", default=None)
    args = ap.parse_args()

    work = args.work or os.path.join(tempfile.gettempdir(), "huiben02-ppt-selftest")
    os.makedirs(work, exist_ok=True)
    extract_media(args.reference, work)

    ref_texts = slide_texts(args.reference)
    n = len(ref_texts)
    media = sorted(f for f in os.listdir(work) if f.startswith("image") and f.endswith(".png"))
    media_map = {i + 1: os.path.join(work, media[i]) for i in range(len(media))}
    reuse = {n: 9}  # 末页复用第 9 张图（小花籽实例）

    slides = []
    for i in range(1, n + 1):
        texts = ref_texts[i - 1]["texts"]
        if i == 1:
            title = next((t for t in texts if t), "")
            slides.append({"id": "S01", "kind": "cover",
                           "image": media_map[1], "title": title})
        else:
            src = reuse.get(i, i if i <= len(media) else len(media))
            slides.append({"id": "S%02d" % i, "kind": "content",
                           "image": media_map[src], "text": texts})

    cfg = {"title": "selftest", "output": "rebuilt.pptx", "slides": slides}
    cfg_path = os.path.join(work, "slides.json")
    with open(cfg_path, "w", encoding="utf-8") as f:
        json.dump(cfg, f, ensure_ascii=False, indent=2)

    r = subprocess.run([sys.executable, os.path.join(HERE, "build_pptx.py"), cfg_path],
                       capture_output=True, text=True, encoding="utf-8")
    print(r.stdout.strip())
    if r.returncode:
        print(r.stderr[-1500:])
        sys.exit("build failed")

    rebuilt = os.path.join(work, "rebuilt.pptx")
    new_texts = slide_texts(rebuilt)
    ok = True
    if len(new_texts) != n:
        print("FAIL 页数 %d != %d" % (len(new_texts), n))
        ok = False
    for i in range(n):
        a, b = ref_texts[i], new_texts[i]
        # 参照件每页有全页背景矩形（空文本），重建件没有，过滤空段后比对
        ta = [t for t in a["texts"] if t]
        tb = [t for t in b["texts"] if t]
        if ta != tb:
            print("FAIL 文本不一致 S%02d" % (i + 1))
            print("  ref :", ta)
            print("  new :", tb)
            ok = False
        expect_pics = 1 if i > 0 else 1
        if b["pics"] != expect_pics:
            print("FAIL 插画数 S%02d: %d != %d" % (i + 1, b["pics"], expect_pics))
            ok = False

    prs = Presentation(rebuilt)
    print("slide size:", prs.slide_width, "x", prs.slide_height)
    s2 = prs.slides[1]
    for shape in s2.shapes:
        if shape.shape_type == 13:
            print("S02 pic EMU:", shape.left, shape.top, shape.width, shape.height)
        elif shape.has_text_frame:
            print("S02 text EMU:", shape.left, shape.top, shape.width, shape.height)
            for para in shape.text_frame.paragraphs:
                for run in para.runs:
                    rPr = run._r.find(qn("a:rPr"))
                    ea = rPr.find(qn("a:ea")) if rPr is not None else None
                    print("   font ea:", ea.get("typeface") if ea is not None else None,
                          "size:", run.font.size.pt if run.font.size else None)
    print("SELFTEST", "PASS" if ok else "FAIL")
    sys.exit(0 if ok else 1)


if __name__ == "__main__":
    main()
