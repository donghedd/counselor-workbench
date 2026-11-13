# -*- coding: utf-8 -*-
"""生成安装包所需的美术资源：app.ico / wizard.bmp / wizard_small.bmp

用法：python installer/make_assets.py
仅依赖 Pillow。改图标只需改这个文件后重跑，再编译 installer.iss。
"""
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from PIL import Image, ImageDraw, ImageFont

try:
    import version
except Exception:
    version = type("v", (), {"VERSION": "0.0.0"})

HERE = os.path.dirname(os.path.abspath(__file__))
ASSETS = os.path.join(HERE, "assets")

BLUE_TOP = (61, 125, 237)
BLUE_BOT = (29, 78, 216)
FONT_BD = r"C:\Windows\Fonts\msyhbd.ttc"
FONT_RG = r"C:\Windows\Fonts\msyh.ttc"


def _gradient(size, top, bottom):
    w, h = size
    img = Image.new("RGB", (1, h))
    d = ImageDraw.Draw(img)
    for y in range(h):
        t = y / max(1, h - 1)
        d.point((0, y), tuple(int(top[i] + (bottom[i] - top[i]) * t) for i in range(3)))
    return img.resize((w, h), Image.BILINEAR)


def _mask_round(size, radius):
    m = Image.new("L", size, 0)
    ImageDraw.Draw(m).rounded_rectangle([0, 0, size[0] - 1, size[1] - 1], radius=radius, fill=255)
    return m


def icon_master(S=1024):
    """1024 母图：蓝色圆角方块 + 白色卡片 + 三行「点名册」条目"""
    img = _gradient((S, S), BLUE_TOP, BLUE_BOT).convert("RGBA")
    img.putalpha(_mask_round((S, S), int(S * 0.22)))

    d = ImageDraw.Draw(img)
    # 白色卡片
    pad = int(S * 0.24)
    card = [pad, int(S * 0.20), S - pad, S - int(S * 0.20)]
    d.rounded_rectangle(card, radius=int(S * 0.055), fill=(255, 255, 255, 255))

    # 三行：左侧圆点 + 右侧横线
    cw = card[2] - card[0]
    row_h = int(cw * 0.115)
    gap = int(cw * 0.075)
    total = row_h * 3 + gap * 2
    y = card[1] + (card[3] - card[1] - total) // 2
    dot = int(row_h * 0.72)
    for i in range(3):
        cx = card[0] + int(cw * 0.13)
        d.ellipse([cx, y + (row_h - dot) // 2, cx + dot, y + (row_h - dot) // 2 + dot],
                  fill=BLUE_BOT + (255,))
        lx = cx + dot + int(cw * 0.075)
        line_w = int(cw * (0.52 if i < 2 else 0.34))
        d.rounded_rectangle([lx, y + int(row_h * 0.24), lx + line_w, y + int(row_h * 0.76)],
                            radius=int(row_h * 0.26), fill=(90, 116, 160, 255))
        y += row_h + gap
    return img


def build_ico(master):
    sizes = [16, 24, 32, 48, 64, 128, 256]
    imgs = [master.resize((s, s), Image.LANCZOS) for s in sizes]
    out = os.path.join(ASSETS, "app.ico")
    imgs[-1].save(out, format="ICO", sizes=[(s, s) for s in sizes])
    print("app.ico", os.path.getsize(out), "bytes")
    return imgs


def wizard_panel(master, k):
    """Inno 现代向导左侧大图，基准 164x314，k 为缩放倍数（2 = 高 DPI 用）"""
    W, H = 164 * k, 314 * k
    bg = _gradient((W, H), (37, 99, 235), (21, 58, 170)).convert("RGBA")
    img = Image.new("RGBA", (W, H), (255, 255, 255, 0))
    img.paste(bg, (0, 0))

    d = ImageDraw.Draw(img)
    s = 92 * k
    im = master.resize((s, s), Image.LANCZOS)
    img.paste(im, ((W - s) // 2, 46 * k), im)

    def center(text, font, y, fill=(255, 255, 255, 255)):
        w = d.textlength(text, font=font)
        d.text(((W - w) / 2, y), text, font=font, fill=fill)

    center("辅导员工作台", ImageFont.truetype(FONT_BD, 15 * k), 160 * k)
    d.line([(W // 2 - 26 * k, 186 * k), (W // 2 + 26 * k, 186 * k)],
           fill=(255, 255, 255, 90), width=max(1, k))
    center("学生数据 · 一屏掌握", ImageFont.truetype(FONT_RG, 10 * k), 196 * k, (214, 228, 255, 235))
    center("v" + version.VERSION, ImageFont.truetype(FONT_RG, 9 * k), H - 26 * k, (170, 196, 245, 220))
    return img


def build_wizard(master, ico_small):
    for k, name in ((1, "wizard.bmp"), (2, "wizard@2x.bmp")):
        p = os.path.join(ASSETS, name)
        wizard_panel(master, k).convert("RGB").save(p, format="BMP")
        print(name, "164x314" if k == 1 else "328x628")

    for k, name in ((1, "wizard_small.bmp"), (2, "wizard_small@2x.bmp")):
        p = os.path.join(ASSETS, name)
        ico_small.resize((55 * k, 55 * k), Image.LANCZOS).convert("RGB").save(p, format="BMP")
        print(name, "%dx%d" % (55 * k, 55 * k))


if __name__ == "__main__":
    os.makedirs(ASSETS, exist_ok=True)
    master = icon_master()
    imgs = build_ico(master)
    build_wizard(master, imgs[5])
