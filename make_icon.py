#!/usr/bin/env python3
"""Erzeugt das Programm-Icon (assets/icon.png, icon_256.png, icon.ico, icon.icns).

Eigenes Design: Farbverlauf-Kachel, kleines Fenster mit rotem Klick-Ring und
darunter eine Kette aus drei Schritten. Aufruf:  python make_icon.py
"""
import os
import sys

from PIL import Image, ImageChops, ImageDraw, ImageFilter

HERE = os.path.dirname(os.path.abspath(__file__))
OUT = os.path.join(HERE, "assets")
S = 2048                                  # Zeichnen in hoher Aufloesung
BLUE, VIOLET = (37, 99, 235), (124, 58, 237)
RED, ACCENT = (239, 68, 68), (59, 130, 246)


def px(v: float) -> int:
    return int(v * S)


def build() -> Image.Image:
    # Diagonaler Farbverlauf
    v = Image.linear_gradient("L").resize((S, S))
    h = v.rotate(90)
    diag = ImageChops.add(v, h, scale=2)
    bg = Image.composite(Image.new("RGB", (S, S), VIOLET),
                         Image.new("RGB", (S, S), BLUE), diag).convert("RGBA")
    # Weicher Glanz oben links
    glow = Image.new("RGBA", (S, S), (0, 0, 0, 0))
    ImageDraw.Draw(glow).ellipse((px(-.3), px(-.4), px(.6), px(.35)),
                                 fill=(255, 255, 255, 46))
    bg = Image.alpha_composite(bg, glow.filter(ImageFilter.GaussianBlur(px(.14))))
    # Kachelform (squircle-aehnlich)
    mask = Image.new("L", (S, S), 0)
    pad = px(.055)
    ImageDraw.Draw(mask).rounded_rectangle(
        (pad, pad, S - pad, S - pad), radius=px(.225), fill=255)
    bg.putalpha(mask)

    art = Image.new("RGBA", (S, S), (0, 0, 0, 0))
    d = ImageDraw.Draw(art)
    card = (px(.19), px(.19), px(.81), px(.65))
    # Schatten der Karte
    sh = Image.new("RGBA", (S, S), (0, 0, 0, 0))
    ImageDraw.Draw(sh).rounded_rectangle(
        (card[0], card[1] + px(.025), card[2], card[3] + px(.025)),
        radius=px(.05), fill=(15, 23, 42, 110))
    art = Image.alpha_composite(art, sh.filter(ImageFilter.GaussianBlur(px(.02))))
    d = ImageDraw.Draw(art)
    d.rounded_rectangle(card, radius=px(.05), fill=(255, 255, 255, 255))
    # Titelleiste mit drei Punkten
    bar = (card[0], card[1], card[2], card[1] + px(.085))
    d.rounded_rectangle(bar, radius=px(.05), fill=(226, 232, 240, 255))
    d.rectangle((bar[0], bar[1] + px(.04), bar[2], bar[3]),
                fill=(226, 232, 240, 255))
    for i in range(3):
        cx, cy, r = card[0] + px(.055 + .04 * i), card[1] + px(.0425), px(.012)
        d.ellipse((cx - r, cy - r, cx + r, cy + r), fill=(148, 163, 184, 255))
    # Inhalt: zwei Textbalken und ein blauer Knopf
    d.rounded_rectangle((px(.25), px(.33), px(.62), px(.355)), px(.012),
                        fill=(203, 213, 225, 255))
    d.rounded_rectangle((px(.25), px(.385), px(.52), px(.41)), px(.012),
                        fill=(226, 232, 240, 255))
    d.rounded_rectangle((px(.25), px(.465), px(.47), px(.55)), px(.03),
                        fill=ACCENT + (255,))
    # Roter Klick-Ring + Punkt (wie in den Screenshots der App)
    cx, cy = px(.64), px(.535)
    for rad, wid, col in ((px(.105), px(.034), (255, 255, 255, 255)),
                          (px(.105), px(.02), RED + (255,))):
        d.ellipse((cx - rad, cy - rad, cx + rad, cy + rad), outline=col, width=wid)
    r = px(.032)
    d.ellipse((cx - r, cy - r, cx + r, cy + r), fill=RED + (255,))
    # Schrittkette: drei Punkte, verbunden
    y = px(.795)
    xs = (px(.34), px(.5), px(.66))
    d.line((xs[0], y, xs[2], y), fill=(255, 255, 255, 190), width=px(.014))
    for i, x in enumerate(xs):
        r = px(.036)
        d.ellipse((x - r, y - r, x + r, y + r),
                  fill=(255, 255, 255, 255) if i < 2 else RED + (255,))
    icon = Image.alpha_composite(bg, art)
    return icon.resize((1024, 1024), Image.LANCZOS)


def main() -> None:
    os.makedirs(OUT, exist_ok=True)
    img = build()
    img.save(os.path.join(OUT, "icon.png"))
    img.resize((256, 256), Image.LANCZOS).save(os.path.join(OUT, "icon_256.png"))
    img.save(os.path.join(OUT, "icon.ico"),
             sizes=[(s, s) for s in (16, 24, 32, 48, 64, 128, 256)])
    try:
        img.save(os.path.join(OUT, "icon.icns"))
    except Exception as exc:                        # sehr alte Pillow-Version
        print("icns uebersprungen:", exc, file=sys.stderr)
    print("Icons erzeugt in", OUT)


if __name__ == "__main__":
    main()
