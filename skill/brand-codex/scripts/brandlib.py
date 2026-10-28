"""Shared helpers for Brand Codex scripts: loading brand.json, color math,
and logo normalization. Standard library + Pillow only (cairosvg optional
for SVG logos)."""

import json
import os
import re
import io

from PIL import Image

HEX_RE = re.compile(r"^#?([0-9a-fA-F]{6})$")


# ---------- loading ----------

def load_brand(path):
    with open(path, "r", encoding="utf-8") as f:
        brand = json.load(f)
    brand["_dir"] = os.path.dirname(os.path.abspath(path))
    return brand


def resolve(brand, rel):
    """Resolve an asset path relative to brand.json."""
    if not rel:
        return None
    if os.path.isabs(rel):
        return rel
    return os.path.join(brand["_dir"], rel)


def g(d, *keys, default=None):
    """Safe nested get."""
    cur = d
    for k in keys:
        if isinstance(cur, dict) and k in cur:
            cur = cur[k]
        else:
            return default
    return cur if cur not in ("", None, [], {}) else default


# ---------- color ----------

def norm_hex(h):
    if not isinstance(h, str):
        return None
    m = HEX_RE.match(h.strip())
    return ("#" + m.group(1).upper()) if m else None


def hex_to_rgb(h):
    h = norm_hex(h)
    if not h:
        return None
    return tuple(int(h[i:i + 2], 16) for i in (1, 3, 5))


def rgb_to_hex(rgb):
    return "#%02X%02X%02X" % tuple(int(v) for v in rgb)


def parse_rgb(v):
    """Accept '16, 25, 33', [16,25,33], or 'rgb(16,25,33)'."""
    if v is None:
        return None
    if isinstance(v, (list, tuple)) and len(v) == 3:
        return tuple(int(x) for x in v)
    nums = re.findall(r"\d+", str(v))
    if len(nums) == 3:
        return tuple(int(x) for x in nums)
    return None


def rel_lum(rgb):
    def ch(c):
        c = c / 255.0
        return c / 12.92 if c <= 0.03928 else ((c + 0.055) / 1.055) ** 2.4
    r, g_, b = rgb
    return 0.2126 * ch(r) + 0.7152 * ch(g_) + 0.0722 * ch(b)


def contrast(h1, h2):
    a, b = rel_lum(hex_to_rgb(h1)), rel_lum(hex_to_rgb(h2))
    hi, lo = max(a, b), min(a, b)
    return (hi + 0.05) / (lo + 0.05)


def wcag_label(ratio):
    if ratio >= 7:
        return "AAA"
    if ratio >= 4.5:
        return "AA"
    if ratio >= 3:
        return "AA Large"
    return "Fail"


def is_dark(h):
    return rel_lum(hex_to_rgb(h)) < 0.18


def readable_on(bg, light="#FFFFFF", dark="#111111"):
    """Pick whichever of light/dark text has more contrast on bg."""
    return light if contrast(bg, light) >= contrast(bg, dark) else dark


def all_swatches(brand):
    out = []
    for grp in g(brand, "colors", "groups", default=[]) or []:
        for sw in grp.get("swatches", []):
            out.append((grp.get("name", ""), sw))
    return out


def color_by_ref(brand, ref):
    """A ref may be a HEX or the name of a swatch."""
    if ref is None:
        return None
    h = norm_hex(ref)
    if h:
        return h
    for _, sw in all_swatches(brand):
        if sw.get("name", "").strip().lower() == str(ref).strip().lower():
            return norm_hex(sw.get("hex"))
    return None


def swatch_name(brand, hexv):
    for _, sw in all_swatches(brand):
        if norm_hex(sw.get("hex")) == norm_hex(hexv):
            return sw.get("name")
    return None


# ---------- logos ----------

def load_logo(path, target_px=1600):
    """Open PNG/JPG/WEBP/SVG as RGBA, trimmed to visible content."""
    ext = os.path.splitext(path)[1].lower()
    if ext == ".svg":
        try:
            import cairosvg
        except ImportError as e:
            raise RuntimeError(
                "SVG logo found but cairosvg is not installed. Run "
                "`pip install cairosvg` or supply a PNG export.") from e
        png = cairosvg.svg2png(url=path, output_width=target_px)
        im = Image.open(io.BytesIO(png)).convert("RGBA")
    else:
        im = Image.open(path).convert("RGBA")
        if not has_transparency(im):
            im = knock_out_background(im)
    return trim(im)


def knock_out_background(im, tol=38):
    """Make a flat background (sampled from the four corners) transparent,
    so opaque JPG/PNG logos can still sit on brand colors. Rough edges are
    expected; the validator still asks for a proper transparent file."""
    w, h = im.size
    corners = [im.getpixel(p)[:3] for p in ((0, 0), (w - 1, 0), (0, h - 1), (w - 1, h - 1))]
    bg = max(set(corners), key=corners.count)
    px = im.load()
    for y in range(h):
        for x in range(w):
            r, g_, b, a = px[x, y]
            d = abs(r - bg[0]) + abs(g_ - bg[1]) + abs(b - bg[2])
            if d < tol:
                px[x, y] = (r, g_, b, 0)
            elif d < tol * 2:
                px[x, y] = (r, g_, b, int(255 * (d - tol) / tol))
    return im


def source_is_opaque(path):
    if path.lower().endswith(".svg"):
        return False
    try:
        return not has_transparency(Image.open(path).convert("RGBA"))
    except Exception:  # noqa
        return False


def has_transparency(im):
    a = im.getchannel("A")
    return a.getextrema()[0] < 250


def trim(im):
    """Trim transparent margins; if fully opaque, trim uniform border color."""
    if has_transparency(im):
        bbox = im.getchannel("A").point(lambda v: 255 if v > 8 else 0).getbbox()
    else:
        from PIL import ImageChops
        bg = Image.new("RGBA", im.size, im.getpixel((0, 0)))
        diff = ImageChops.difference(im, bg).convert("L").point(
            lambda v: 255 if v > 18 else 0)
        bbox = diff.getbbox()
    return im.crop(bbox) if bbox else im


def logo_pixels(im):
    """Visible logo pixels. For opaque files, ignore the background color
    (taken from the corner pixel) so it is not counted as artwork."""
    small = im.copy()
    small.thumbnail((200, 200))
    data = list(small.getdata())
    if has_transparency(small):
        return [p[:3] for p in data if p[3] > 200]
    bg = data[0][:3]
    return [p[:3] for p in data if sum(abs(a - b) for a, b in zip(p[:3], bg)) > 40]


def visible_colors(im, n=5):
    """Dominant opaque colors of a logo, as HEX list."""
    px = logo_pixels(im)
    if not px:
        return []
    q = Image.new("RGB", (len(px), 1))
    q.putdata(px)
    q = q.quantize(colors=n, method=Image.Quantize.MEDIANCUT)
    pal = q.getpalette()[: n * 3]
    counts = sorted(q.getcolors(), reverse=True)
    return [rgb_to_hex(pal[i * 3:i * 3 + 3]) for _, i in counts]


def mean_luminance(im):
    px = logo_pixels(im)
    if not px:
        return 0.5
    return sum(rel_lum(p) for p in px) / len(px)


def logo_variants(brand):
    return g(brand, "logos", "variants", default=[]) or []


def preferred_variant(brand):
    vs = logo_variants(brand)
    pref = g(brand, "logos", "preferred")
    for v in vs:
        if v.get("id") == pref:
            return v
    return vs[0] if vs else None


def variant_for(brand, mode):
    """First variant meant for a 'light' or 'dark' background."""
    for v in logo_variants(brand):
        if v.get("use_on") == mode and v.get("type", "full") in ("full", "horizontal", "primary"):
            return v
    for v in logo_variants(brand):
        if v.get("use_on") == mode:
            return v
    return None


def suggest_background(im):
    """'light', 'dark', or 'either', from the average contrast of the logo's
    visible pixels against white and against near-black."""
    px = logo_pixels(im)
    if not px:
        return "either", 0, 0
    lw, lb = rel_lum((255, 255, 255)), rel_lum((17, 17, 17))

    def cr(a, b):
        return (max(a, b) + 0.05) / (min(a, b) + 0.05)
    lums = [rel_lum(p) for p in px]
    # weight toward the weakest parts: a logo is only as legible as its
    # least-contrasting element, so use the 20th percentile contrast
    on_white = sorted(cr(l, lw) for l in lums)[len(lums) // 5]
    on_black = sorted(cr(l, lb) for l in lums)[len(lums) // 5]
    if on_white >= 3 and on_black >= 3:
        return "either", on_white, on_black
    return ("light" if on_white >= on_black else "dark"), on_white, on_black
