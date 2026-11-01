# -*- coding: utf-8 -*-
"""生成小虾记账应用图标（ICO 多尺寸 + PNG）

用法: python tools/generate_icon.py
输出: assets/icon.ico, assets/icon.png
"""
from pathlib import Path

from PIL import Image, ImageDraw, ImageFont

ROOT = Path(__file__).resolve().parent.parent
ASSETS = ROOT / "assets"
ASSETS.mkdir(parents=True, exist_ok=True)

SIZE = 1024
# 品牌色
BLUE_TOP = (37, 99, 235)     # #2563eb
BLUE_BOTTOM = (30, 58, 138)  # #1e3a8a
GOLD = (245, 158, 11)        # #f59e0b


def find_font() -> str:
    candidates = [
        r"C:\Windows\Fonts\msyhbd.ttc",
        r"C:\Windows\Fonts\msyh.ttc",
        r"C:\Windows\Fonts\simhei.ttf",
        r"C:\Windows\Fonts\segoeuib.ttf",
    ]
    for c in candidates:
        if Path(c).exists():
            return c
    raise SystemExit("找不到可用字体")


def vertical_gradient(size, top, bottom):
    grad = Image.new("RGB", (1, size), top)
    for y in range(size):
        t = y / max(size - 1, 1)
        color = tuple(int(top[i] + (bottom[i] - top[i]) * t) for i in range(3))
        grad.putpixel((0, y), color)
    return grad.resize((size, size))


def rounded_mask(size, radius):
    mask = Image.new("L", (size, size), 0)
    ImageDraw.Draw(mask).rounded_rectangle([0, 0, size - 1, size - 1], radius=radius, fill=255)
    return mask


def build_icon() -> Image.Image:
    canvas = Image.new("RGBA", (SIZE, SIZE), (0, 0, 0, 0))

    # 背景：圆角 + 纵向渐变
    bg = vertical_gradient(SIZE, BLUE_TOP, BLUE_BOTTOM).convert("RGBA")
    bg.putalpha(rounded_mask(SIZE, int(SIZE * 0.22)))
    canvas.alpha_composite(bg)

    draw = ImageDraw.Draw(canvas)

    # 主字「虾」：白色
    font_path = find_font()
    font_size = int(SIZE * 0.56)
    font = ImageFont.truetype(font_path, font_size)
    text = "虾"
    bbox = draw.textbbox((0, 0), text, font=font)
    tw, th = bbox[2] - bbox[0], bbox[3] - bbox[1]
    tx = (SIZE - tw) / 2 - bbox[0]
    ty = (SIZE - th) / 2 - bbox[1] - int(SIZE * 0.045)
    draw.text((tx, ty), text, font=font, fill=(255, 255, 255, 255))

    # 底部金色短线，增加识别度
    bar_w = int(SIZE * 0.26)
    bar_h = int(SIZE * 0.045)
    bx = (SIZE - bar_w) / 2
    by = SIZE - int(SIZE * 0.145)
    draw.rounded_rectangle([bx, by, bx + bar_w, by + bar_h], radius=bar_h // 2, fill=GOLD + (255,))

    return canvas


def main():
    icon = build_icon()

    png_path = ASSETS / "icon.png"
    icon.resize((512, 512), Image.LANCZOS).save(png_path, "PNG")

    ico_path = ASSETS / "icon.ico"
    sizes = [(16, 16), (24, 24), (32, 32), (48, 48), (64, 64), (128, 128), (256, 256)]
    icon.save(ico_path, format="ICO", sizes=sizes)

    print(f"[OK] {png_path}")
    print(f"[OK] {ico_path}")


if __name__ == "__main__":
    main()
