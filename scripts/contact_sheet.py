# -*- coding: utf-8 -*-
"""把一批插画拼成联系表 PNG，用于人眼检查风格一致性/角色走形/文字残留。

用法：
    python contact_sheet.py -o 联系表.png 02-插画/*.png
    python contact_sheet.py -o 联系表.png --dir 02-插画
    python contact_sheet.py -o 联系表.png --dir 02-插画 --cols 4 --cell 1024
"""
import argparse
import glob
import math
import os

from PIL import Image, ImageDraw


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("-o", "--out", required=True)
    ap.add_argument("images", nargs="*")
    ap.add_argument("--dir", dest="dir", default=None)
    ap.add_argument("--cols", type=int, default=0, help="列数，默认自动")
    ap.add_argument("--cell", type=int, default=768, help="单元格短边像素")
    args = ap.parse_args()

    paths = list(args.images)
    if args.dir:
        exts = (".png", ".jpg", ".jpeg", ".webp")
        paths += [p for p in sorted(glob.glob(os.path.join(args.dir, "*")))
                  if p.lower().endswith(exts)]
    if not paths:
        raise SystemExit("没有输入图片")

    cell_h = args.cell
    cell_w = int(cell_h * 1.5)
    cols = args.cols or max(1, min(len(paths), int(math.ceil(math.sqrt(len(paths) * 1.5)))))
    rows = int(math.ceil(len(paths) / cols))
    sheet = Image.new("RGB", (cols * cell_w, rows * (cell_h + 28)), "white")
    draw = ImageDraw.Draw(sheet)
    for i, p in enumerate(paths):
        im = Image.open(p).convert("RGB")
        im.thumbnail((cell_w, cell_h))
        x = (i % cols) * cell_w
        y = (i // cols) * (cell_h + 28)
        sheet.paste(im, (x + (cell_w - im.width) // 2, y))
        draw.text((x + 6, y + cell_h + 6), os.path.basename(p), fill="black")
    os.makedirs(os.path.dirname(os.path.abspath(args.out)) or ".", exist_ok=True)
    sheet.save(args.out)
    print("saved: %s (%dx%d, %d images)" % (args.out, cols, rows, len(paths)))


if __name__ == "__main__":
    main()
