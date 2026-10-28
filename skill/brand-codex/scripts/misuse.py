"""Generate 'incorrect logo usage' example tiles from a real logo file.

Each function returns an RGB PIL image (tile) showing one violation, so the
guide can show the brand's own logo misused instead of generic placeholders.
"""

import math
import random

from PIL import Image, ImageDraw, ImageFilter, ImageFont, ImageOps

from brandlib import hex_to_rgb, rel_lum

TILE = (900, 520)


def _font(size):
    for name in ("DejaVuSans.ttf", "Arial.ttf", "Helvetica.ttc"):
        try:
            return ImageFont.truetype(name, size)
        except Exception:  # noqa
            continue
    try:
        return ImageFont.load_default(size=size)
    except TypeError:
        return ImageFont.load_default()


def _canvas(bg):
    return Image.new("RGBA", TILE, hex_to_rgb(bg) + (255,))


def _fit(logo, max_w, max_h):
    im = logo.copy()
    im.thumbnail((max_w, max_h), Image.LANCZOS)
    return im


def _paste_center(canvas, im, dx=0, dy=0):
    x = (canvas.width - im.width) // 2 + dx
    y = (canvas.height - im.height) // 2 + dy
    canvas.alpha_composite(im, (max(x, 0), max(y, 0)))
    return canvas


def base_logo(logo):
    return _fit(logo, int(TILE[0] * 0.62), int(TILE[1] * 0.5))


def stretch(logo, bg, **_):
    im = base_logo(logo)
    im = im.resize((int(im.width * 1.45), int(im.height * 0.7)), Image.LANCZOS)
    im = _fit(im, int(TILE[0] * 0.85), TILE[1])
    return _paste_center(_canvas(bg), im).convert("RGB")


def squash(logo, bg, **_):
    im = base_logo(logo)
    im = im.resize((int(im.width * 0.6), int(im.height * 1.35)), Image.LANCZOS)
    return _paste_center(_canvas(bg), im).convert("RGB")


def rotate(logo, bg, **_):
    im = base_logo(logo).rotate(14, expand=True, resample=Image.BICUBIC)
    return _paste_center(_canvas(bg), im).convert("RGB")


def recolor(logo, bg, **_):
    im = base_logo(logo)
    a = im.getchannel("A")
    rgb = im.convert("RGB")
    lum = sum(rel_lum(p) for p in rgb.resize((40, 20)).getdata()) / 800
    if lum < 0.04 or lum > 0.85:
        # near-mono logo: hue shift does nothing, tint it instead
        tint = Image.new("RGB", im.size, (214, 120, 28))
        rgb = Image.blend(rgb, tint, 0.85)
    else:
        hsv = rgb.convert("HSV")
        h, s, v = hsv.split()
        h = h.point(lambda x: (x + 110) % 256)
        rgb = Image.merge("HSV", (h, s, v)).convert("RGB")
    out = rgb.convert("RGBA")
    out.putalpha(a)
    return _paste_center(_canvas(bg), out).convert("RGB")


def effects(logo, bg, **_):
    im = base_logo(logo)
    pad = 60
    shadow = Image.new("RGBA", (im.width + pad * 2, im.height + pad * 2), (0, 0, 0, 0))
    sh = Image.new("RGBA", im.size, (0, 0, 0, 255))
    sh.putalpha(im.getchannel("A").point(lambda v: int(v * 0.75)))
    shadow.alpha_composite(sh, (pad + 18, pad + 18))
    shadow = shadow.filter(ImageFilter.GaussianBlur(12))
    shadow.alpha_composite(im, (pad, pad))
    return _paste_center(_canvas(bg), shadow).convert("RGB")


def outline(logo, bg, **_):
    im = base_logo(logo)
    pad = 20
    a = Image.new("L", (im.width + pad * 2, im.height + pad * 2), 0)
    a.paste(im.getchannel("A"), (pad, pad))
    grown = a.filter(ImageFilter.MaxFilter(13))
    col = (255, 214, 0) if rel_lum(hex_to_rgb(bg)) < 0.5 else (230, 0, 120)
    stroke = Image.new("RGBA", a.size, col + (255,))
    stroke.putalpha(grown)
    stroke.alpha_composite(im, (pad, pad))
    return _paste_center(_canvas(bg), stroke).convert("RGB")


def busy(logo, bg, palette=None, **_):
    rnd = random.Random(7)
    cols = [hex_to_rgb(c) for c in (palette or [])] or [(200, 60, 60), (60, 120, 200), (240, 200, 40)]
    c = Image.new("RGBA", TILE, (180, 180, 180, 255))
    d = ImageDraw.Draw(c)
    for _ in range(140):
        x, y = rnd.randint(-80, TILE[0]), rnd.randint(-80, TILE[1])
        r = rnd.randint(20, 110)
        col = rnd.choice(cols + [(255, 255, 255), (30, 30, 30)])
        d.ellipse([x, y, x + r, y + r], fill=col + (255,))
    for i in range(0, TILE[0] + TILE[1], 38):
        d.line([(i, 0), (i - TILE[1], TILE[1])], fill=(255, 255, 255, 120), width=6)
    c = c.filter(ImageFilter.GaussianBlur(1.2))
    return _paste_center(c, base_logo(logo)).convert("RGB")


def wrong_mode(logo, bg, **_):
    """Put the given (light-mode) logo on a dark background, or vice versa."""
    return _paste_center(_canvas(bg), base_logo(logo)).convert("RGB")


def low_contrast(logo, bg, **_):
    return _paste_center(_canvas(bg), base_logo(logo)).convert("RGB")


def crowd(logo, bg, **_):
    c = _canvas(bg)
    im = _fit(logo, int(TILE[0] * 0.5), int(TILE[1] * 0.32))
    x = (TILE[0] - im.width) // 2
    y = (TILE[1] - im.height) // 2
    c.alpha_composite(im, (x, y))
    d = ImageDraw.Draw(c)
    dark = rel_lum(hex_to_rgb(bg)) < 0.4
    ink = (235, 235, 235, 255) if dark else (60, 60, 60, 255)
    blk = (120, 120, 120, 255)
    f = _font(30)
    d.text((x, y - 44), "Introducing our new platform", font=f, fill=ink)
    d.text((x + 10, y + im.height + 6), "Learn more at example.com", font=f, fill=ink)
    d.rectangle([x + im.width + 6, y - 10, x + im.width + 120, y + im.height + 10], fill=blk)
    d.rectangle([x - 120, y + 10, x - 6, y + im.height - 10], fill=blk)
    return c.convert("RGB")


def sentence(logo, bg, **_):
    c = _canvas(bg)
    d = ImageDraw.Draw(c)
    dark = rel_lum(hex_to_rgb(bg)) < 0.4
    ink = (235, 235, 235, 255) if dark else (50, 50, 50, 255)
    f = _font(44)
    pre, post = "Try", "today for free."
    im = _fit(logo, 400, 54)
    w1 = d.textlength(pre + " ", font=f)
    w2 = d.textlength(" " + post, font=f)
    total = w1 + im.width + w2
    x = int((TILE[0] - total) / 2)
    y = TILE[1] // 2 - 26
    d.text((x, y), pre + " ", font=f, fill=ink)
    c.alpha_composite(im, (int(x + w1), y + 26 - im.height // 2 + 4))
    d.text((x + w1 + im.width, y), " " + post, font=f, fill=ink)
    return c.convert("RGB")


GENERATORS = {
    "stretch": (stretch, "Do not stretch or distort the logo"),
    "squash": (squash, "Do not squash or change its proportions"),
    "rotate": (rotate, "Do not rotate, angle, or skew the logo"),
    "recolor": (recolor, "Do not recolor the logo or use unapproved colors"),
    "effects": (effects, "Do not add drop shadows, glows, or other effects"),
    "outline": (outline, "Do not outline or stroke the logo"),
    "busy": (busy, "Do not place the logo on busy backgrounds or patterns"),
    "wrong_mode": (wrong_mode, "Do not use the light-mode logo on dark backgrounds"),
    "low_contrast": (low_contrast, "Do not place the logo on low-contrast colors"),
    "crowd": (crowd, "Do not crowd the logo or ignore its clear space"),
    "sentence": (sentence, "Do not use the logo as a word in a sentence"),
}

DEFAULT_ORDER = ["stretch", "rotate", "recolor", "effects", "outline", "busy",
                 "wrong_mode", "crowd", "low_contrast", "sentence", "squash"]
