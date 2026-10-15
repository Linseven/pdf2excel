#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""生成应用图标：assets/icon.png / icon.icns (macOS) / icon.ico (Windows)。

用法：.venv/bin/python tools/make_icons.py
"""

from __future__ import annotations

import os
import shutil
import subprocess
import sys
import tempfile

from PIL import Image, ImageDraw

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
ASSETS = os.path.join(ROOT, "assets")
SIZE = 1024
C_TOP = (99, 102, 241)    # #6366F1
C_BOTTOM = (79, 70, 229)  # #4F46E5
C_LINE = (79, 70, 229)


def lerp(a, b, t):
    return int(a + (b - a) * t)


def make_base(size: int) -> Image.Image:
    """圆角渐变底板 + 白色表格符号。"""
    s = size
    img = Image.new("RGBA", (s, s), (0, 0, 0, 0))

    # 渐变圆角底板（macOS 风格约 22.5% 圆角率）
    radius = int(s * 0.225)
    mask = Image.new("L", (s, s), 0)
    ImageDraw.Draw(mask).rounded_rectangle([0, 0, s - 1, s - 1], radius=radius, fill=255)
    grad = Image.new("RGBA", (s, s))
    gd = ImageDraw.Draw(grad)
    for y in range(s):
        t = y / max(s - 1, 1)
        gd.line([(0, y), (s, y)], fill=(lerp(C_TOP[0], C_BOTTOM[0], t),
                                        lerp(C_TOP[1], C_BOTTOM[1], t),
                                        lerp(C_TOP[2], C_BOTTOM[2], t), 255))
    img.paste(grad, (0, 0), mask)

    # 白色表格：表头实心 + 下方 2×3 网格
    d = ImageDraw.Draw(img)
    k = s / SIZE
    x0, y0, x1, y1 = 272 * k, 328 * k, 752 * k, 696 * k
    header_h = 118 * k
    r = 26 * k
    d.rounded_rectangle([x0, y0, x1, y1], radius=r, fill=(255, 255, 255, 255))

    lw = 13 * k
    hy = y0 + header_h                     # 表头分隔线
    by = y0 + header_h + (y1 - y0 - header_h) / 2  # 行分隔线
    cx1 = x0 + (x1 - x0) / 3               # 列分隔线 1
    cx2 = x0 + (x1 - x0) * 2 / 3           # 列分隔线 2
    line = (*C_LINE, 255)
    d.rectangle([x0, hy - lw / 2, x1, hy + lw / 2], fill=line)
    d.rectangle([x0, by - lw / 2, x1, by + lw / 2], fill=line)
    d.rectangle([cx1 - lw / 2, hy, cx1 + lw / 2, y1], fill=line)
    d.rectangle([cx2 - lw / 2, hy, cx2 + lw / 2, y1], fill=line)
    return img


def make_icns(base: Image.Image) -> bool:
    """macOS：借助系统 iconutil 生成 .icns。"""
    if shutil.which("iconutil") is None:
        print("跳过 icon.icns（未找到 iconutil，非 macOS 环境）")
        return False
    with tempfile.TemporaryDirectory() as td:
        iconset = os.path.join(td, "AppIcon.iconset")
        os.makedirs(iconset)
        for sz in (16, 32, 64, 128, 256, 512, 1024):
            base.resize((sz, sz), Image.LANCZOS).save(os.path.join(iconset, f"icon_{sz}x{sz}.png"))
            if sz <= 512:  # @2x 变体
                base.resize((sz * 2, sz * 2), Image.LANCZOS).save(os.path.join(iconset, f"icon_{sz}x{sz}@2x.png"))
        out = os.path.join(ASSETS, "icon.icns")
        subprocess.run(["iconutil", "-c", "icns", iconset, "-o", out], check=True)
        print("已生成", out)
        return True


def main() -> int:
    os.makedirs(ASSETS, exist_ok=True)
    base = make_base(SIZE)

    png_path = os.path.join(ASSETS, "icon.png")
    base.save(png_path)
    print("已生成", png_path)

    ico_path = os.path.join(ASSETS, "icon.ico")
    base.resize((256, 256), Image.LANCZOS).save(
        ico_path, sizes=[(16, 16), (24, 24), (32, 32), (48, 48), (64, 64), (128, 128), (256, 256)])
    print("已生成", ico_path)

    make_icns(base)
    return 0


if __name__ == "__main__":
    sys.exit(main())
