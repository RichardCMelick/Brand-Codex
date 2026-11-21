#!/usr/bin/env python3
"""Build a brand guidelines PowerPoint from brand.json.

Every value on every slide comes from brand.json, so the deck cannot
contradict itself. Sections with no data are skipped, not filled with
placeholder text.

Usage:
  python build_deck.py brand.json out.pptx [--fallback-fonts]

--fallback-fonts  set slide text in each typeface's fallback (use when the
                  brand fonts are not installed on the machines that will
                  open the deck)
"""

import io
import os
import sys
import copy

from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.oxml.ns import qn
from lxml import etree
from PIL import Image, ImageDraw

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from brandlib import (load_brand, resolve, g, norm_hex, hex_to_rgb, parse_rgb,
                      contrast, wcag_label, is_dark, readable_on, color_by_ref,
                      load_logo, logo_variants, preferred_variant, variant_for,
                      all_swatches, rel_lum)
import misuse

W, H = 13.333, 7.5
LEFT_W = 4.35          # text column width
RX = 4.85              # right visual area x (inside the inset panel)
RW = W - RX - 0.85     # right visual area width
PANEL_R = 0.18         # corner radius of the inset visual panel
HAIR = 0.0104          # hairline thickness (0.75pt) in inches


# ======================================================================
# Theme
# ======================================================================

class Theme:
    def __init__(self, brand, fallback_fonts=False):
        self.brand = brand
        modes = g(brand, "colors", "modes", default={}) or {}
        lm, dm = modes.get("light", {}) or {}, modes.get("dark", {}) or {}
        sw = [norm_hex(s.get("hex")) for _, s in all_swatches(brand) if norm_hex(s.get("hex"))]
        darkest = min(sw, key=lambda h: rel_lum(hex_to_rgb(h))) if sw else "#111111"
        self.light_bg = color_by_ref(brand, lm.get("background")) or "#FFFFFF"
        self.light_surface = color_by_ref(brand, lm.get("surface")) or self._shade(self.light_bg, -0.04)
        self.light_text = color_by_ref(brand, lm.get("text")) or readable_on(self.light_bg)
        self.light_muted = color_by_ref(brand, lm.get("text_muted")) or self._mix(self.light_text, self.light_bg, 0.35)
        self.dark_bg = color_by_ref(brand, dm.get("background")) or darkest
        self.dark_text = color_by_ref(brand, dm.get("text")) or readable_on(self.dark_bg)
        self.dark_muted = color_by_ref(brand, dm.get("text_muted")) or self._mix(self.dark_text, self.dark_bg, 0.35)
        acc = color_by_ref(brand, lm.get("accent"))
        if not acc:
            prim = (g(brand, "colors", "groups", default=[{}]) or [{}])[0].get("swatches", [])
            cands = [norm_hex(s.get("hex")) for s in prim if norm_hex(s.get("hex"))]
            cands = [c for c in cands if 0.03 < rel_lum(hex_to_rgb(c)) < 0.85] or cands
            acc = cands[0] if cands else "#2F6FED"
        self.accent = acc
        self.accent_ink = readable_on(acc)
        self.border = color_by_ref(brand, lm.get("border")) or self._mix(self.light_text, self.light_bg, 0.85)
        # hairline for dividers: lighter than a border so lists read as rows, not boxes
        self.hair = self._mix(self.light_text, self.light_bg, 0.82)
        self.good, self.bad = "#2F7D5B", "#B5473A"
        fams = g(brand, "typography", "families", default=[]) or []

        def pick(roles):
            for f in fams:
                if str(f.get("role", "")).lower() in roles:
                    return f
            return None
        head = pick({"headline", "display", "heading", "headings"}) or (fams[0] if fams else {})
        body = pick({"body", "text", "copy"}) or (fams[1] if len(fams) > 1 else head)
        if fallback_fonts:
            self.head_font = head.get("fallback") or head.get("name") or "Arial"
            self.body_font = body.get("fallback") or body.get("name") or "Arial"
        else:
            self.head_font = head.get("name") or "Arial"
            self.body_font = body.get("name") or "Arial"
        self.fallback_fonts = fallback_fonts

    @staticmethod
    def _mix(a, b, t):
        ra, rb = hex_to_rgb(a), hex_to_rgb(b)
        return "#%02X%02X%02X" % tuple(int(x + (y - x) * t) for x, y in zip(ra, rb))

    def _shade(self, h, amt):
        return self._mix(h, "#000000" if amt < 0 else "#FFFFFF", abs(amt))


# ======================================================================
# Drawing primitives
# ======================================================================

def rgb(h):
    r, g_, b = hex_to_rgb(h)
    return RGBColor(r, g_, b)


def rect(slide, x, y, w, h, fill, line=None, shape=MSO_SHAPE.RECTANGLE, radius=None, name=None):
    s = slide.shapes.add_shape(shape, Inches(x), Inches(y), Inches(w), Inches(h))
    if fill:
        s.fill.solid()
        s.fill.fore_color.rgb = rgb(fill)
    else:
        s.fill.background()
    if line:
        s.line.color.rgb = rgb(line)
        s.line.width = Pt(0.75)
    else:
        s.line.fill.background()
    s.shadow.inherit = False
    st = s._element.find(qn("p:style"))   # theme style carries the default drop shadow
    if st is not None:
        s._element.remove(st)
    if radius is not None and shape == MSO_SHAPE.ROUNDED_RECTANGLE:
        s.adjustments[0] = radius
    if name:
        s.name = name
    tf = s.text_frame
    tf.text = ""
    return s


def rrect(slide, x, y, w, h, fill, line=None, r=0.14, name=None):
    """Rounded rectangle with a corner radius given in inches."""
    frac = min(0.5, r / max(0.01, min(w, h)))
    return rect(slide, x, y, w, h, fill, line=line, shape=MSO_SHAPE.ROUNDED_RECTANGLE,
                radius=frac, name=name)


def rule(slide, x, y, w, color, weight=HAIR):
    """Hairline divider."""
    return rect(slide, x, y, w, weight, color)


def text(slide, x, y, w, h, content, size=12, color="#000000", font="Arial",
         bold=False, italic=False, align="left", anchor="top", spacing=None,
         bullets=False, line_spacing=None, caps=False, para_after=4):
    """content: str or list of str (paragraphs) or list of list-of-runs.
    A run is (text, {opts}) with opts size/color/bold/italic/font."""
    tb = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    tf = tb.text_frame
    tf.word_wrap = True
    tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
    tf.vertical_anchor = {"top": MSO_ANCHOR.TOP, "middle": MSO_ANCHOR.MIDDLE,
                          "bottom": MSO_ANCHOR.BOTTOM}[anchor]
    paras = content if isinstance(content, list) else [content]
    for i, para in enumerate(paras):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = {"left": PP_ALIGN.LEFT, "center": PP_ALIGN.CENTER,
                       "right": PP_ALIGN.RIGHT}[align]
        if line_spacing:
            p.line_spacing = line_spacing
        p.space_after = Pt(para_after)
        runs = para if isinstance(para, list) else [(para, {})]
        for t, o in runs:
            r = p.add_run()
            t = str(t)
            r.text = t.upper() if (caps or o.get("caps")) else t
            f = r.font
            f.size = Pt(o.get("size", size))
            f.bold = o.get("bold", bold)
            f.italic = o.get("italic", italic)
            f.name = o.get("font", font)
            f.color.rgb = rgb(o.get("color", color))
            if spacing is not None or o.get("spacing") is not None:
                rPr = r._r.get_or_add_rPr()
                rPr.set("spc", str(int((o.get("spacing", spacing)) * 100)))
        if bullets:
            _bullet(p, color)
    return tb


def _bullet(p, color):
    pPr = p._p.get_or_add_pPr()
    pPr.set("marL", str(Emu(Inches(0.18))))
    pPr.set("indent", str(-Emu(Inches(0.18))))
    for tag in ("a:buNone", "a:buChar", "a:buAutoNum"):
        for el in pPr.findall(qn(tag)):
            pPr.remove(el)
    buClr = etree.SubElement(pPr, qn("a:buClr"))
    srgb = etree.SubElement(buClr, qn("a:srgbClr"))
    srgb.set("val", norm_hex(color)[1:])
    bu = etree.SubElement(pPr, qn("a:buChar"))
    bu.set("char", "•")


def _round_corners(im, radius_px):
    im = im.convert("RGBA")
    ss = 3
    mask = Image.new("L", (im.width * ss, im.height * ss), 0)
    ImageDraw.Draw(mask).rounded_rectangle(
        [0, 0, im.width * ss - 1, im.height * ss - 1], radius=radius_px * ss, fill=255)
    mask = mask.resize(im.size, Image.LANCZOS)
    a = im.getchannel("A")
    from PIL import ImageChops
    im.putalpha(ImageChops.multiply(a, mask))
    return im


def picture(slide, pil_or_path, x, y, w, h, align="center", radius=None):
    """Place an image fitted inside the box (keeps aspect). radius (inches)
    rounds the corners."""
    if isinstance(pil_or_path, Image.Image):
        im = pil_or_path
    else:
        im = Image.open(pil_or_path)
    iw, ih = im.size
    scale = min(w / iw, h / ih)
    pw, ph = iw * scale, ih * scale
    if align == "center":
        px, py = x + (w - pw) / 2, y + (h - ph) / 2
    elif align == "left":
        px, py = x, y + (h - ph) / 2
    else:
        px, py = x + w - pw, y + (h - ph) / 2
    if radius:
        im = _round_corners(im, max(1, int(radius / pw * im.width)))
    buf = io.BytesIO()
    im.save(buf, format="PNG")
    buf.seek(0)
    return slide.shapes.add_picture(buf, Inches(px), Inches(py), Inches(pw), Inches(ph))


def set_bg(slide, color):
    bg = slide.background.fill
    bg.solid()
    bg.fore_color.rgb = rgb(color)


# ======================================================================
# Deck builder
# ======================================================================

class Deck:
    def __init__(self, brand, fallback_fonts=False):
        self.b = brand
        self.t = Theme(brand, fallback_fonts)
        self.prs = Presentation()
        self.prs.slide_width = Inches(W)
        self.prs.slide_height = Inches(H)
        self.blank = self.prs.slide_layouts[6]
        self.section = ""
        self.section_no = 0
        self.content_slides = []
        self.divider_slides = []
        self.toc_slots = []
        self.toc = []
        self.company = g(brand, "meta", "company", default="Company")
        self._logo_cache = {}
        self.warnings = []

    # ---------- assets ----------
    def logo(self, variant):
        if not variant:
            return None
        key = variant.get("file")
        if key not in self._logo_cache:
            p = resolve(self.b, key)
            try:
                self._logo_cache[key] = load_logo(p)
            except Exception as e:  # noqa
                self.warnings.append(f"Could not load logo {key}: {e}")
                self._logo_cache[key] = None
        return self._logo_cache[key]

    def mark_for(self, mode):
        for v in logo_variants(self.b):
            if v.get("type") == "mark" and v.get("use_on") == mode:
                return v
        return variant_for(self.b, mode)

    def variant_bg(self, v):
        if v.get("background"):
            return color_by_ref(self.b, v["background"]) or self.t.light_bg
        on = v.get("use_on", "light")
        if on == "dark":
            return self.t.dark_bg
        if on == "color":
            return self.t.accent
        return self.t.light_bg

    # ---------- slide scaffolds ----------
    def new(self, bg=None):
        s = self.prs.slides.add_slide(self.blank)
        set_bg(s, bg or self.t.light_bg)
        return s

    def label(self, s, x, y, w, label, color=None, size=10):
        """Quiet small-caps caption used inside visuals (never as a heading badge)."""
        return text(s, x, y, w, 0.3, label, size=size, color=color or self.t.light_muted,
                    font=self.t.body_font, bold=True, caps=True, spacing=1.2)

    def proposed_tag(self, s, x, y):
        """Status flag for values Claude suggested that are not yet approved."""
        w = 0.98
        rrect(s, x, y, w, 0.27, "#F4D58D", r=0.135)
        text(s, x, y, w, 0.27, "Proposed", size=9, color="#4A3500", font=self.t.body_font,
             bold=True, align="center", anchor="middle", caps=True, spacing=1)

    def _title_lines(self, title, size, width):
        per_line = max(6, int(width / (0.5 * size / 72)))
        words, lines, cur = title.split(), 1, 0
        for w_ in words:
            if cur and cur + 1 + len(w_) > per_line:
                lines += 1
                cur = len(w_)
            else:
                cur += (1 if cur else 0) + len(w_)
        return lines

    def content(self, title, body=None, eyebrow=None, right_bg=None, bullets=None):
        """Split slide: type column on the left, visual area on the right,
        set in a softly rounded inset panel instead of a hard half-page block."""
        t = self.t
        s = self.new(t.light_bg)
        if right_bg is not False:
            rrect(s, LEFT_W, 0.4, W - LEFT_W - 0.4, H - 0.8, right_bg or t.light_surface,
                  r=PANEL_R)
        colw = LEFT_W - 0.95
        size = 34 if len(title) <= 18 else 30 if len(title) <= 26 else 26
        text(s, 0.6, 0.85, colw, 1.8, title, size=size, color=t.light_text,
             font=t.head_font, line_spacing=0.98)
        y = 0.85 + self._title_lines(title, size, colw) * size * 1.02 / 72 + 0.35
        if body:
            paras = body if isinstance(body, list) else [body]
            text(s, 0.6, y, colw, 3.2, paras, size=13, color=t.light_muted,
                 font=t.body_font, line_spacing=1.22, para_after=8)
            per = int(colw / (0.5 * 13 / 72))
            y += sum(-(-len(p) // per) for p in paras) * 13 * 1.22 / 72 + 0.15 * len(paras) + 0.2
        if bullets:
            text(s, 0.6, y, colw, H - y - 1.0, bullets, size=12, color=t.light_text,
                 font=t.body_font, bullets=True, para_after=6, line_spacing=1.15)
        self.content_slides.append((s, self.section, True))
        return s

    def full(self, title, eyebrow=None, bg=None, subtitle=None):
        """Full-width slide with the heading across the top."""
        t = self.t
        bg = bg or t.light_bg
        ink = readable_on(bg, t.dark_text if is_dark(bg) else "#FFFFFF", t.light_text)
        s = self.new(bg)
        text(s, 0.6, 0.7, W - 1.2, 0.7, title, size=34, color=ink, font=t.head_font)
        if subtitle:
            text(s, 0.6, 1.42, 8.6, 0.6, subtitle, size=14, line_spacing=1.2,
                 color=t.dark_muted if is_dark(bg) else t.light_muted, font=t.body_font)
        self.content_slides.append((s, self.section, False))
        return s

    def divider(self, title, blurb=None):
        """Section opener. Type sits low on the page with room above it."""
        self.section_no += 1
        self.section = title
        t = self.t
        s = self.new(t.dark_bg)
        self.divider_slides.append((title, s))
        text(s, 0.8, 3.05, 10.5, 1.3, title, size=64, color=t.dark_text, font=t.head_font,
             anchor="bottom", line_spacing=0.95)
        rule(s, 0.84, 4.62, 0.7, t.accent, weight=0.03)
        if blurb:
            text(s, 0.82, 4.9, 6.6, 1.2, blurb, size=17, color=t.dark_muted,
                 font=t.body_font, line_spacing=1.25)
        v = self.mark_for("dark")
        im = self.logo(v)
        if im:
            picture(s, im, 0.8, 0.6, 0.8, 0.45, align="left")
        return s

    def finish(self):
        t = self.t
        total = len(self.prs.slides)
        mark = self.logo(self.mark_for("light"))
        order = {id(sl): i for i, sl in enumerate(self.prs.slides, start=1)}
        for slide, sec, split in self.content_slides:
            idx = order[id(slide)]
            fy = H - 0.62
            x = 0.6
            if mark:
                picture(slide, mark, x, fy - 0.01, 0.3, 0.26, align="left")
                x += 0.42
            text(slide, x, fy, 3.2, 0.25, f"{self.company}   {sec}", size=9,
                 color=t.light_muted, font=t.body_font, anchor="middle")
            if split:
                px, pw = LEFT_W - 1.1, 0.7
            else:
                px, pw = W - 1.3, 0.7
            text(slide, px, fy, pw, 0.25, f"{idx}", size=9, color=t.light_muted,
                 font=t.body_font, align="right", anchor="middle")
        # table of contents shows where each section starts
        for (title, sl), (box_slide, y_, x_) in zip(self.divider_slides, self.toc_slots):
            text(box_slide, x_, y_, 1.0, 0.5, str(order[id(sl)]), size=14, color=t.light_muted,
                 font=t.body_font, align="right", anchor="middle")

    # ==================================================================
    # Sections
    # ==================================================================

    def cover(self):
        t, b = self.t, self.b
        v = variant_for(b, "dark")
        dark = v is not None
        if not dark:
            v = variant_for(b, "light") or preferred_variant(b)
        bg = t.dark_bg if dark else t.light_bg
        ink = t.dark_text if dark else t.light_text
        muted = t.dark_muted if dark else t.light_muted
        s = self.new(bg)
        hero = resolve(b, g(b, "meta", "cover_image"))
        if hero and os.path.exists(hero):
            im = Image.open(hero).convert("RGB")
            tw, th = 1600, 900
            sc = max(tw / im.width, th / im.height)
            im = im.resize((int(im.width * sc), int(im.height * sc)))
            l, tp = (im.width - tw) // 2, (im.height - th) // 2
            im = im.crop((l, tp, l + tw, tp + th))
            ov = Image.new("RGB", im.size, hex_to_rgb(bg))
            im = Image.blend(im, ov, 0.45)
            picture(s, im, 0, 0, W, H)
        im = self.logo(v)
        if im:
            picture(s, im, 0.8, 0.7, 2.6, 0.75, align="left")
        title = g(b, "meta", "guide_title", default="Brand Guidelines")
        text(s, 0.8, 2.2, 10.8, 2.4, title, size=72, color=ink, font=t.head_font,
             anchor="bottom", line_spacing=0.92)
        sub = g(b, "meta", "subtitle") or g(b, "foundation", "tagline")
        if sub:
            text(s, 0.84, 4.95, 8.4, 0.9, sub, size=20, color=muted, font=t.body_font,
                 line_spacing=1.2)
        rule(s, 0.86, 4.72, 0.7, t.accent, weight=0.03)
        ver = "   ".join(x for x in [
            (f"Version {g(b, 'meta', 'version')}" if g(b, "meta", "version") else None),
            g(b, "meta", "date")] if x)
        text(s, 0.84, H - 0.85, 6, 0.3, ver, size=11, color=muted, font=t.body_font)
        text(s, W - 5.4, H - 0.85, 4.6, 0.3, f"© {self.company}. All rights reserved.",
             size=11, color=muted, font=t.body_font, align="right")

    def purpose(self):
        b, t = self.b, self.t
        p = g(b, "meta", "purpose")
        if not p:
            return
        s = self.new(t.light_bg)
        text(s, 0.9, 1.3, 7.4, 4.6, [p], size=24, color=t.light_text,
             font=t.head_font, line_spacing=1.22)
        y = 1.4
        for head, val in (("Who it is for", g(b, "meta", "audience")),
                          ("Where to get files", g(b, "governance", "asset_location"))):
            if not val:
                continue
            rule(s, 9.2, y, 3.3, t.hair)
            text(s, 9.2, y + 0.18, 3.3, 0.3, head, size=11, bold=True, color=t.light_text,
                 font=t.body_font)
            text(s, 9.2, y + 0.5, 3.3, 1.4, val, size=12, color=t.light_muted,
                 font=t.body_font, line_spacing=1.2)
            y += 2.0
        self.content_slides.append((s, "Introduction", False))

    def toc_slide(self, sections):
        t = self.t
        s = self.new(t.light_bg)
        rrect(s, 0.4, 0.4, LEFT_W - 0.4, H - 0.8, t.dark_bg, r=PANEL_R)
        text(s, 0.9, 0.95, LEFT_W - 1.3, 1, "Contents", size=36, color=t.dark_text,
             font=t.head_font)
        x0, cw = 5.4, W - 5.4 - 0.9
        rh = min(0.8, 5.6 / max(1, len(sections)))
        y0 = (H - rh * len(sections)) / 2
        for i, name in enumerate(sections):
            y = y0 + i * rh
            rule(s, x0, y, cw, t.hair)
            text(s, x0, y, cw - 1.2, rh, name, size=20, color=t.light_text,
                 font=t.head_font, anchor="middle")
            self.toc_slots.append((s, y + rh / 2 - 0.25, x0 + cw - 1.0))
        rule(s, x0, y0 + rh * len(sections), cw, t.hair)

    def mission(self, closing=False):
        b, t = self.b, self.t
        m = g(b, "foundation", "mission")
        if not m:
            return
        s = self.new(t.accent if not closing else t.dark_bg)
        ink = readable_on(t.accent) if not closing else t.dark_text
        size = 46 if len(m) < 110 else 38 if len(m) < 190 else 30
        text(s, 0.9, 1.2, 10.6, 5.0, m, size=size, color=ink, font=t.head_font,
             line_spacing=1.08, anchor="middle")
        if closing:
            im = self.logo(variant_for(b, "dark"))
            if im:
                picture(s, im, W - 3.4, H - 1.25, 2.6, 0.6, align="right")

    def foundation(self):
        b, t = self.b, self.t
        f = g(b, "foundation", default={}) or {}
        has = any(f.get(k) for k in ("vision", "positioning", "boilerplate", "pillars",
                                      "values", "voice", "direction"))
        if not has:
            return
        self.divider("Brand Foundation", "Who we are, what we stand for, and how we sound.")
        PX = RX            # content x inside the inset panel
        PW = RW

        # Positioning + vision + boilerplate
        if f.get("positioning") or f.get("boilerplate") or f.get("vision"):
            s = self.content("Who We Are", f.get("positioning"))
            y = 0.95
            if f.get("vision"):
                self.label(s, PX, y, PW, "Vision")
                n = len(f["vision"])
                vs = 28 if n < 90 else 22
                text(s, PX, y + 0.38, PW, 1.9, f["vision"], size=vs, color=t.light_text,
                     font=t.head_font, line_spacing=1.1)
                y += 0.38 + (1.3 if n < 90 else 1.0) + 0.55
            if f.get("boilerplate"):
                rule(s, PX, y, PW, t.hair)
                self.label(s, PX, y + 0.25, PW, "Approved company description")
                text(s, PX, y + 0.65, PW, H - y - 1.8, f["boilerplate"], size=13.5,
                     color=t.light_text, font=t.body_font, line_spacing=1.28)

        # Pillars / values
        pillars = f.get("pillars") or []
        if pillars:
            s = self.full(f.get("pillars_title") or "Brand Pillars",
                          subtitle=f.get("pillars_intro"))
            n = len(pillars[:4])
            gap = 0.5
            cw = (W - 1.2 - gap * (n - 1)) / n
            for i, p in enumerate(pillars[:4]):
                x = 0.6 + i * (cw + gap)
                rule(s, x, 2.35, cw, t.light_text, weight=0.015)
                text(s, x, 2.6, cw, 0.8, p.get("name", ""), size=34,
                     color=t.light_text, font=t.head_font)
                y = 3.55
                if p.get("description"):
                    text(s, x, y, cw, 1.3, p["description"], size=16,
                         color=t.light_muted, font=t.body_font, line_spacing=1.2)
                    y += 0.5 + 0.3 * (len(p["description"]) // 32)
                pts = p.get("points") or []
                if pts:
                    text(s, x, y, cw, H - y - 1.0, pts, size=14, color=t.light_text,
                         font=t.body_font, bullets=True, para_after=9)

        values = f.get("values") or []
        if values:
            s = self.full("Our Values")
            n = len(values[:6])
            cols = 3 if n > 4 else n
            rows = (n + cols - 1) // cols
            gap = 0.5
            cw = (W - 1.2 - gap * (cols - 1)) / cols
            ch = (4.6 - 0.3 * (rows - 1)) / rows
            for i, v in enumerate(values[:6]):
                x = 0.6 + (i % cols) * (cw + gap)
                y = 2.1 + (i // cols) * (ch + 0.3)
                rule(s, x, y, cw, t.hair)
                name = v.get("name") if isinstance(v, dict) else str(v)
                desc = v.get("description") if isinstance(v, dict) else None
                text(s, x, y + 0.25, cw, 0.5, name, size=22, color=t.light_text,
                     font=t.head_font)
                if desc:
                    text(s, x, y + 0.85, cw, ch - 1.0, desc, size=13,
                         color=t.light_muted, font=t.body_font, line_spacing=1.2)

        if f.get("direction") and f.get("direction_public", True):
            d = f["direction"]
            paras = d if isinstance(d, list) else [d]
            s = self.content("Where We Are Going",
                             "Our direction shapes how the brand shows up. Use it to "
                             "judge whether new work moves us forward.")
            text(s, PX, 0.9, PW, H - 1.8, paras,
                 size=28 if len(" ".join(paras)) < 260 else 20,
                 color=t.light_text, font=t.head_font, line_spacing=1.15, para_after=14,
                 anchor="middle")

        # Voice
        v = f.get("voice") or {}
        traits = v.get("traits") or []
        if traits or v.get("summary"):
            s = self.content("Voice and Tone", v.get("summary"),
                             bullets=v.get("tone_notes"))
            if traits:
                rows = traits[:6]
                rh = min(0.95, 5.2 / len(rows))
                half = PW / 2
                self.label(s, PX, 0.95, half, "We are", color=t.accent)
                self.label(s, PX + half, 0.95, half, "We are not", color=t.bad)
                for i, tr in enumerate(rows):
                    y = 1.45 + i * rh
                    rule(s, PX, y, PW, t.hair)
                    lines = [[(tr.get("we_are", ""), {"bold": True, "size": 16})]]
                    if tr.get("we_are_note"):
                        lines.append(tr["we_are_note"])
                    text(s, PX, y + 0.05, half - 0.3, rh - 0.1, lines, size=11.5,
                         color=t.light_text, font=t.body_font, anchor="middle", para_after=1)
                    text(s, PX + half, y + 0.05, half - 0.1, rh - 0.1, tr.get("we_are_not", ""),
                         size=16, color=t.light_muted, font=t.body_font, anchor="middle")

        # Audiences
        auds = g(self.b, "audiences", default=[]) or []
        if auds:
            s = self.full("Who We Speak To")
            n = len(auds[:3])
            gap = 0.5
            cw = (W - 1.2 - gap * (n - 1)) / n
            for i, a in enumerate(auds[:3]):
                x = 0.6 + i * (cw + gap)
                rule(s, x, 2.0, cw, t.light_text, weight=0.02)
                text(s, x, 2.25, cw, 0.5, a.get("name", ""), size=24,
                     color=t.light_text, font=t.head_font)
                y = 2.85
                if a.get("titles"):
                    text(s, x, y, cw, 0.6, a["titles"], size=12,
                         color=t.light_muted, font=t.body_font)
                    y += 0.7
                if a.get("needs"):
                    text(s, x, y, cw, 1.2, a["needs"], size=14,
                         color=t.light_text, font=t.head_font, italic=False, line_spacing=1.2)
                    y += 1.4
                if a.get("tone"):
                    text(s, x, y, cw, H - y - 1.0, a["tone"], size=12,
                         color=t.light_text, font=t.body_font, bullets=True, para_after=5)

    # ---------- logo ----------
    def logos(self):
        b, t = self.b, self.t
        vs = logo_variants(b)
        if not vs:
            return
        L = g(b, "logos", default={}) or {}
        self.divider("Logo", "Our most recognizable asset. Use only approved files, exactly as supplied.")

        # Primary logo: split, logo full-bleed on right
        pv = preferred_variant(b)
        bg = self.variant_bg(pv)
        s = self.content("Our Logo", L.get("description") or
                         "The logo is the primary visual signature of the brand. "
                         "Always use the approved artwork; never redraw or modify it.",
                         right_bg=bg, bullets=L.get("anatomy"))
        im = self.logo(pv)
        if im:
            picture(s, im, LEFT_W + 1.4, 1.6, W - LEFT_W - 2.8, H - 3.2)

        # Variations grid
        s = self.content("Logo Variations", L.get("variations_intro") or
                         "Choose the version by the background it sits on, not by taste.")
        tiles = vs[:6]
        cols = 3 if len(tiles) > 4 else 2
        rows = (len(tiles) + cols - 1) // cols
        gap = 0.3
        cw = (RW - gap * (cols - 1)) / cols
        ch = (H - 1.9 - gap * (rows - 1)) / rows
        for i, v in enumerate(tiles):
            x = RX + (i % cols) * (cw + gap)
            y = 0.85 + (i // cols) * (ch + gap)
            vbg = self.variant_bg(v)
            rrect(s, x, y, cw, ch - 0.95, vbg, line=t.border if not is_dark(vbg) else None, r=0.12)
            im = self.logo(v)
            if im:
                picture(s, im, x + cw * 0.18, y + 0.25, cw * 0.64, ch - 1.45)
            label = v.get("name", v.get("id", ""))
            if v.get("id") == L.get("preferred"):
                label += "  (preferred)"
            text(s, x, y + ch - 0.85, cw, 0.28, label, size=12, bold=True,
                 color=t.light_text, font=t.body_font)
            if v.get("usage"):
                text(s, x, y + ch - 0.57, cw, 0.5, v["usage"], size=10, color=t.light_muted,
                     font=t.body_font, line_spacing=1.1)
            if v.get("proposed"):
                self.proposed_tag(s, x + 0.12, y + 0.12)

        # Clear space + minimum size
        cs = L.get("clear_space") or {}
        ms = L.get("min_size") or {}
        if cs or ms:
            body = []
            if cs.get("rule"):
                body.append(cs["rule"])
            s = self.content("Clear Space and Minimum Size", body)
            pl = variant_for(b, "light") or pv
            im = self.logo(pl)
            if im and cs.get("ratio"):
                diag = clear_space_diagram(im, float(cs["ratio"]), t.light_bg, t.accent,
                                           cs.get("unit_label", "x"))
                picture(s, diag, RX, 0.8, RW, 3.9)
                text(s, RX, 4.75, RW, 0.4,
                     f"x = {cs.get('unit_description', str(cs['ratio']) + ' x logo height')}",
                     size=11, color=t.light_muted, font=t.body_font, align="center")
            elif im:
                picture(s, im, RX + 1.5, 1.2, RW - 3, 2.6)
            if ms:
                y = 5.3
                items = [(k.replace("_", " ").title(), v) for k, v in ms.items() if v]
                cw = RW / max(1, len(items))
                for i, (k, v) in enumerate(items):
                    x = RX + i * cw
                    rule(s, x, y, cw - 0.3, t.hair)
                    self.label(s, x, y + 0.18, cw - 0.3, k)
                    text(s, x, y + 0.5, cw - 0.3, 0.6, str(v), size=18,
                         color=t.light_text, font=t.head_font)
            if cs.get("small_size_rule"):
                text(s, 0.6, 5.3, LEFT_W - 0.95, 1.4, cs["small_size_rule"], size=11,
                     color=t.light_muted, font=t.body_font, line_spacing=1.2)

        # Background pairings (approved)
        pairs = L.get("backgrounds") or []
        if pairs:
            s = self.full("Approved Background Pairings",
                          subtitle="Each logo version paired with the backgrounds it is approved for.")
            n = len(pairs[:8])
            cols = 4 if n > 3 else n
            rows = (n + cols - 1) // cols
            cw = (W - 1.2 - 0.3 * (cols - 1)) / cols
            ch = min(2.6, (5.0 - 0.3 * (rows - 1)) / rows)
            for i, pr in enumerate(pairs[:8]):
                x = 0.6 + (i % cols) * (cw + 0.3)
                y = 2.2 + (i // cols) * (ch + 0.3)
                pbg = color_by_ref(b, pr.get("background")) or t.light_bg
                v = next((vv for vv in vs if vv.get("id") == pr.get("variant")), None)
                rrect(s, x, y, cw, ch - 0.55, pbg, line=t.border if not is_dark(pbg) else None, r=0.12)
                im = self.logo(v)
                if im:
                    picture(s, im, x + cw * 0.2, y + 0.25, cw * 0.6, ch - 0.95)
                status = pr.get("status", "approved")
                col = {"preferred": t.accent, "approved": t.light_text,
                       "minimal": "#B7791F"}.get(status, t.light_text)
                text(s, x, y + ch - 0.45, cw, 0.4,
                     [[(status.capitalize(), {"bold": True, "color": col, "size": 11}),
                       ("   " + (pr.get("note") or ""), {"size": 10})]],
                     color=t.light_muted, font=t.body_font)

        # Misuse grid
        self.logo_misuse()

        # Custom don'ts and where not to use
        custom = L.get("donts_custom") or []
        where = L.get("where_not") or []
        if custom or where:
            s = self.content("Brand-Specific Rules",
                             "These go beyond the generic rules above. They exist "
                             "because of how our brand has actually been misused.")
            y = 0.85
            if where:
                self.label(s, RX, y, RW, "Where our logo does not appear", color=t.bad)
                text(s, RX, y + 0.4, RW, 2.6, where[:8], size=13, color=t.light_text,
                     font=t.body_font, bullets=True, para_after=6)
                y += 0.75 + 0.36 * min(8, len(where))
            if custom:
                imgs = [c for c in custom if isinstance(c, dict) and c.get("image")]
                txts = [c if isinstance(c, str) else c.get("text", "") for c in custom
                        if not (isinstance(c, dict) and c.get("image"))]
                if txts:
                    self.label(s, RX, y + 0.1, RW, "Also do not", color=t.bad)
                    text(s, RX, y + 0.5, RW, 2.6, txts[:8], size=13, color=t.light_text,
                         font=t.body_font, bullets=True, para_after=6)
                    y += 0.8 + 0.36 * min(8, len(txts))
                if imgs and y < 5.2:
                    n = min(4, len(imgs))
                    cw = (RW - 0.2 * (n - 1)) / n
                    for i, c in enumerate(imgs[:n]):
                        x = RX + i * (cw + 0.2)
                        p = resolve(b, c["image"])
                        if p and os.path.exists(p):
                            picture(s, p, x, y + 0.1, cw, H - y - 1.6)
                        self._x_badge(s, x + 0.05, y + 0.15)
                        text(s, x, H - 1.35, cw, 0.6, c.get("text", ""), size=8.5,
                             color=t.light_text, font=t.body_font)

        # Lockups
        if L.get("lockup_rules"):
            self.content("Lockups and Extended Family",
                         L.get("lockup_intro") or "How names and sub-brands may appear with the logo.",
                         bullets=L["lockup_rules"])
            # (visual examples, if provided)
            ex = L.get("lockup_examples") or []
            s = self.content_slides[-1][0]
            for i, e in enumerate(ex[:3]):
                p = resolve(b, e)
                if p and os.path.exists(p):
                    picture(s, p, RX, 0.9 + i * 1.9, RW, 1.6)

    def list_slide(self, title, items, intro=None, kind="rule", dos=None):
        """Typographic list on hairline rows. kind: 'rule' (numbered), 'dont', 'do'.
        If dos is given, renders two columns: dos left, items (don'ts) right."""
        t = self.t
        s = self.full(title, subtitle=intro)
        top = (2.45 if intro else 1.95) + (0.45 if dos else 0)
        avail = H - top - 0.95

        def rows(lst, x0, colw, cols, k):
            lst = lst[:12]
            nrows = max(1, (len(lst) + cols - 1) // cols)
            rh = min(1.15, avail / nrows)
            fs = 20 if nrows <= 3 else 17 if nrows <= 5 else 15
            for i, it in enumerate(lst):
                x = x0 + (i % cols) * (colw + 0.5)
                y = top + (i // cols) * rh
                rule(s, x, y, colw, t.hair)
                if k == "rule":
                    text(s, x, y, 0.55, rh, f"{i + 1}", size=fs + 5, color=t.accent,
                         font=t.head_font, anchor="middle")
                    tx = x + 0.65
                else:
                    self._mark(s, x + 0.02, y + rh / 2 - 0.13, ok=(k == "do"))
                    tx = x + 0.55
                text(s, tx, y + 0.06, colw - (tx - x) - 0.1, rh - 0.12, str(it), size=fs,
                     color=t.light_text, font=t.body_font, anchor="middle", line_spacing=1.1)
        if dos:
            cw = (W - 1.2 - 0.6) / 2
            self.label(s, 0.6, top - 0.5, cw, "Do", color=t.good, size=11)
            self.label(s, 0.6 + cw + 0.6, top - 0.5, cw, "Don't", color=t.bad, size=11)
            rows(dos, 0.6, cw, 1, "do")
            rows(items, 0.6 + cw + 0.6, cw, 1, "dont")
        else:
            cols = 2 if len(items) > 4 else 1
            cw = (W - 1.2 - 0.5 * (cols - 1)) / cols
            rows(items, 0.6, cw, cols, kind)
        return s

    def _mark(self, s, x, y, ok=False, d=0.26):
        """Drawn status mark (filled disc with a stroked check or cross)."""
        col = self.t.good if ok else self.t.bad
        rect(s, x, y, d, d, col, shape=MSO_SHAPE.OVAL)
        u = d / 10.0
        pts = [(2.6, 5.3), (4.4, 7.0), (7.6, 3.2)] if ok else None
        strokes = [pts] if ok else [[(3.2, 3.2), (6.8, 6.8)], [(6.8, 3.2), (3.2, 6.8)]]
        for pl in strokes:
            fb = s.shapes.build_freeform(Inches(x + pl[0][0] * u), Inches(y + pl[0][1] * u))
            fb.add_line_segments([(Inches(x + px * u), Inches(y + py * u)) for px, py in pl[1:]],
                                 close=False)
            sh = fb.convert_to_shape()
            sh.fill.background()
            sh.line.color.rgb = rgb("#FFFFFF")
            sh.line.width = Pt(1.6)
            st = sh._element.find(qn("p:style"))
            if st is not None:
                sh._element.remove(st)

    def _x_badge(self, s, x, y, ok=False):
        self._mark(s, x, y, ok=ok, d=0.3)

    def logo_misuse(self):
        b, t = self.b, self.t
        L = g(b, "logos", default={}) or {}
        keys = L.get("donts_auto")
        if keys is None:
            keys = ["stretch", "rotate", "recolor", "effects", "outline", "busy",
                    "wrong_mode", "crowd"]
        keys = [k for k in keys if k in misuse.GENERATORS]
        if not keys:
            return
        light_v = variant_for(b, "light") or preferred_variant(b)
        dark_v = variant_for(b, "dark")
        light_im = self.logo(light_v)
        if light_im is None:
            return
        palette = [norm_hex(s.get("hex")) for _, s in all_swatches(b) if norm_hex(s.get("hex"))]
        # low-contrast background: a mid-luminance swatch near the logo's luminance
        from brandlib import mean_luminance
        ll = mean_luminance(light_im)
        mid = sorted(palette, key=lambda h: abs(rel_lum(hex_to_rgb(h)) - ll))
        low_bg = next((h for h in mid if contrast(h, t.light_bg) > 1.5), None) or "#8A8A8A"
        tiles = []
        for k in keys:
            fn, cap = misuse.GENERATORS[k]
            if k == "wrong_mode":
                if not dark_v:
                    continue
                img = fn(light_im, t.dark_bg)
            elif k == "low_contrast":
                img = fn(light_im, low_bg)
            else:
                img = fn(light_im, t.light_surface, palette=palette)
            tiles.append((img, cap))
        per = 8
        for page in range(0, len(tiles), per):
            chunk = tiles[page:page + per]
            title = "Incorrect Logo Usage" if page == 0 else "Incorrect Logo Usage (continued)"
            s = self.full(title, subtitle=L.get("donts_intro") or
                          "Never alter the logo. These examples use our real logo to show what not to do.")
            cols = 4
            gap = 0.25
            cw = (W - 1.2 - gap * (cols - 1)) / cols
            ch = cw * TILE_RATIO
            for i, (img, cap) in enumerate(chunk):
                x = 0.6 + (i % cols) * (cw + gap)
                y = 2.05 + (i // cols) * (ch + 0.78)
                picture(s, img, x, y, cw, ch, radius=0.12)
                self._x_badge(s, x + 0.1, y + 0.1)
                text(s, x, y + ch + 0.1, cw, 0.6, cap, size=11, color=t.light_text,
                     font=t.body_font, line_spacing=1.1)

    # ---------- architecture / co-branding ----------
    def architecture(self):
        b, t = self.b, self.t
        A = g(b, "architecture", default=None)
        if not A:
            return
        tiers = A.get("tiers") or []
        s = self.content("Brand Architecture", A.get("intro"), eyebrow="Logo")
        y = 0.9
        rh = min(1.3, (H - 2.0) / max(1, len(tiers)))
        for i, tr in enumerate(tiers[:5]):
            yy = y + i * rh
            rule(s, RX, yy, RW, t.hair)
            text(s, RX, yy + 0.2, 2.9, 0.4, tr.get("name", f"Tier {i + 1}"), size=18,
                 color=t.light_text, font=t.head_font)
            text(s, RX, yy + 0.62, 2.9, rh - 0.7, tr.get("rule", ""), size=11,
                 color=t.light_muted, font=t.body_font, line_spacing=1.15)
            ex = tr.get("examples") or []
            text(s, RX + 3.3, yy + 0.1, RW - 3.3, rh - 0.2, ", ".join(ex) if ex else "",
                 size=15, color=t.light_text, font=t.body_font, anchor="middle")

    def cobranding(self):
        b, t = self.b, self.t
        C = g(b, "cobranding", default=None)
        if not C:
            return
        self.section = "Co-branding"
        rules = C.get("rules") or []
        s = self.content("Partner Co-branding", C.get("intro") or
                         "How our logo appears alongside partner logos.",
                         eyebrow="Logo",
                         bullets=[(r if isinstance(r, str) else f"{r.get('title')}: {r.get('text')}")
                                  for r in rules[:7]])
        light_v = variant_for(b, "light") or preferred_variant(b)
        dark_v = variant_for(b, "dark")
        sep = (C.get("separator") or "none").lower()
        examples = [(t.light_bg, light_v, t.light_text)]
        if dark_v:
            examples.append((t.dark_bg, dark_v, t.dark_text))
        eh = 2.3
        for i, (bg, v, ink) in enumerate(examples):
            y = 0.9 + i * (eh + 0.35)
            rrect(s, RX, y, RW, eh, bg, line=t.border if not is_dark(bg) else None, r=0.14)
            im = self.logo(v)
            if im:
                picture(s, im, RX + 0.9, y + 0.6, RW / 2 - 1.5, eh - 1.2, align="right")
            cx = RX + RW / 2
            if sep in ("line", "vertical rule", "rule", "vertical line"):
                rect(s, cx - 0.01, y + 0.55, 0.02, eh - 1.1, ink)
            elif sep in ("plus", "+"):
                text(s, cx - 0.3, y, 0.6, eh, "+", size=30, color=ink, align="center",
                     anchor="middle", font=t.body_font)
            elif sep in ("dotted", "dotted rule"):
                for k in range(8):
                    rect(s, cx - 0.02, y + 0.55 + k * 0.15, 0.04, 0.06, ink)
            pr = rect(s, cx + 0.6, y + 0.75, RW / 2 - 1.5, eh - 1.5, None, line=ink,
                      shape=MSO_SHAPE.ROUNDED_RECTANGLE, radius=0.1)
            text(s, cx + 0.6, y + 0.75, RW / 2 - 1.5, eh - 1.5, "PARTNER LOGO", size=10,
                 color=ink, align="center", anchor="middle", font=t.body_font, spacing=2)
        if C.get("min_gap"):
            text(s, RX, 0.9 + len(examples) * (eh + 0.35), RW, 0.3,
                 f"Minimum space between logos: {C['min_gap']}", size=12,
                 color=t.light_muted, font=t.body_font)
        donts = C.get("donts") or []
        if donts:
            imgs = C.get("dont_images") or []
            if not imgs:
                self.list_slide("Co-branding: What Not To Do", donts, kind="dont")
                return
            self.content("Co-branding: What Not To Do", None, eyebrow="Logo",
                         bullets=donts[:10])
            s2 = self.content_slides[-1][0]
            for i, d in enumerate(imgs[:4]):
                p = resolve(b, d.get("image") if isinstance(d, dict) else d)
                if p and os.path.exists(p):
                    x = RX + (i % 2) * (RW / 2 + 0.1)
                    y = 0.9 + (i // 2) * 3.0
                    picture(s2, p, x, y, RW / 2 - 0.1, 2.4)
                    self._x_badge(s2, x + 0.05, y + 0.05)

    # ---------- color ----------
    def colors(self):
        b, t = self.b, self.t
        groups = g(b, "colors", "groups", default=[]) or []
        if not groups:
            return
        C = g(b, "colors", default={})
        self.divider("Color", C.get("intro") or "Color is one of the fastest ways people recognize us.")
        for grp in groups:
            sws = grp.get("swatches", [])
            if not sws:
                continue
            per_page = 8 if len(sws) <= 8 else 12
            for pg in range(0, len(sws), per_page):
                chunk = sws[pg:pg + per_page]
                gname = grp.get("name", "Color")
                if any(k in gname.lower() for k in ("color", "colour", "palette")):
                    title = gname
                else:
                    title = ("Neutral" if gname.lower() == "neutrals" else gname) + " Palette"
                if pg:
                    title += " (continued)"
                s = self.content(title, grp.get("usage"))
                self._swatch_grid(s, chunk)

        # Gradients
        grads = C.get("gradients") or []
        if grads:
            s = self.content("Gradients", C.get("gradient_rules"))
            n = len(grads[:6])
            cols = 2 if n > 1 else 1
            cw = (RW - 0.25 * (cols - 1)) / cols
            rows = (n + cols - 1) // cols
            ch = min(3.6 if n == 1 else 2.2, (H - 1.7) / rows - 0.25)
            for i, gd in enumerate(grads[:6]):
                x = RX + (i % cols) * (cw + 0.25)
                y = 0.85 + (i // cols) * (ch + 0.3)
                stops = [color_by_ref(b, c) for c in gd.get("stops", [])]
                stops = [c for c in stops if c]
                if len(stops) >= 2:
                    picture(s, gradient_image(stops, gd.get("angle", 90)), x, y, cw, ch - 0.55,
                            radius=0.14)
                text(s, x, y + ch - 0.45, cw, 0.4,
                     [[(gd.get("name", ""), {"bold": True}),
                       ("   " + " > ".join(stops) + (f"   {gd.get('angle')}°" if gd.get("angle") is not None else ""),
                        {"color": t.light_muted})]], size=11, color=t.light_text, font=t.body_font)

        self.modes()
        self.ratio()
        self.contrast_matrix()
        if C.get("donts"):
            self.list_slide("Color: What Not To Do", C["donts"], kind="dont")

    def _swatch_grid(self, s, chunk):
        """Large rounded color fields with open captions. A short palette gets
        one dominant swatch so the page has a focal point, not equal tiles."""
        t = self.t
        n = len(chunk)
        cols = 4 if n <= 8 else 6
        rows = (n + cols - 1) // cols
        gap = 0.28 if cols == 4 else 0.2
        small = cols == 6
        weights = [1.0] * cols
        if n <= 4 and n > 1:
            cols = n
            weights = [1.7] + [1.0] * (n - 1)
        wsum = sum(weights[:cols])
        unit = (RW - gap * (cols - 1)) / wsum
        top, avail = 0.85, H - 1.7
        most = max(1 + (1 if s_.get("cmyk") else 0) + (1 if s_.get("pms") else 0)
                   for s_ in chunk) + 1
        any_role = any(s_.get("role") for s_ in chunk)
        for i, sw in enumerate(chunk):
            r_, c_ = divmod(i, cols)
            cw = unit * weights[c_]
            x = RX + sum(unit * weights[k] + gap for k in range(c_))
            h = norm_hex(sw.get("hex")) or "#FF00FF"
            rgbv = parse_rgb(sw.get("rgb")) or hex_to_rgb(h)
            vals = [f"HEX  {h}", f"RGB  {', '.join(str(v) for v in rgbv)}"]
            if sw.get("cmyk"):
                vals.append(f"CMYK  {sw['cmyk']}")
            if sw.get("pms"):
                vals.append(f"PMS  {sw['pms']}")
            fs = 10 if small else 11.5
            cap_h = (0.34 if small else 0.4) + most * (fs * 1.45 / 72) + (0.22 if any_role else 0)
            row_h = min(5.7, (avail - gap * (rows - 1)) / rows)
            y = top + r_ * (row_h + gap)
            fh = row_h - cap_h - 0.12
            rrect(s, x, y, cw, fh, h, line=t.border if rel_lum(hex_to_rgb(h)) > 0.85 else None,
                  r=0.14)
            lines = [[(sw.get("name", ""), {"bold": True, "size": 13 if not small else 10.5,
                                             "color": t.light_text})]] + vals
            if sw.get("role"):
                lines.append([(sw["role"], {"italic": True})])
            text(s, x, y + fh + 0.12, cw, cap_h, lines, size=fs,
                 color=t.light_muted, font=t.body_font, para_after=1)
            if sw.get("proposed"):
                self.proposed_tag(s, x + 0.12, y + 0.12)

    def modes(self):
        b, t = self.b, self.t
        modes = g(b, "colors", "modes", default={}) or {}
        if not (modes.get("light") or modes.get("dark")):
            return
        s = self.full("Light Mode and Dark Mode",
                      subtitle=g(b, "colors", "modes_intro") or
                      "The same brand on two backgrounds. Each mode has its own text, accent, and logo version.")
        keys = ["background", "surface", "text", "text_muted", "border", "accent", "link"]
        which = [m for m in ("light", "dark") if modes.get(m)]
        pw = (W - 1.1 - 0.4 * (len(which) - 1)) / len(which)
        for i, m in enumerate(which):
            spec = modes[m]
            x = 0.55 + i * (pw + 0.4)
            y = 2.05
            bg = color_by_ref(b, spec.get("background")) or (t.light_bg if m == "light" else t.dark_bg)
            surf = color_by_ref(b, spec.get("surface")) or bg
            tx = color_by_ref(b, spec.get("text")) or readable_on(bg)
            mu = color_by_ref(b, spec.get("text_muted")) or tx
            bd = color_by_ref(b, spec.get("border")) or Theme._mix(tx, bg, 0.8)
            ac = color_by_ref(b, spec.get("accent")) or t.accent
            lk = color_by_ref(b, spec.get("link")) or ac
            ph = 3.15
            rrect(s, x, y, pw, ph, bg, line=bd, r=0.16)
            v = None
            for vv in logo_variants(b):
                if vv.get("id") == spec.get("logo_variant"):
                    v = vv
            v = v or variant_for(b, m)
            im = self.logo(v)
            if im:
                picture(s, im, x + 0.3, y + 0.25, 1.6, 0.4, align="left")
            text(s, x + pw - 2.3, y + 0.3, 2.0, 0.3, f"{m.upper()} MODE", size=10,
                 color=mu, bold=True, align="right", font=t.body_font, spacing=1.5)
            cx, cy, cwid = x + 0.3, y + 0.9, pw - 0.6
            rrect(s, cx, cy, cwid, 2.0, surf, line=bd, r=0.1)
            text(s, cx + 0.25, cy + 0.2, cwid - 0.5, 0.45, "Headline in primary text", size=19,
                 color=tx, font=t.head_font)
            text(s, cx + 0.25, cy + 0.68, cwid - 0.5, 0.5,
                 "Secondary text sits here and carries supporting detail.", size=11.5,
                 color=mu, font=t.body_font)
            btn = rrect(s, cx + 0.25, cy + 1.3, 1.5, 0.42, ac, r=0.21)
            text(s, cx + 0.25, cy + 1.3, 1.5, 0.42, "Primary action", size=9.5, bold=True,
                 color=readable_on(ac), align="center", anchor="middle", font=t.body_font)
            tb = text(s, cx + 1.95, cy + 1.36, 1.8, 0.35, "Text link", size=10, color=lk,
                      font=t.body_font, anchor="middle")
            tb.text_frame.paragraphs[0].runs[0].font.underline = True
            # token list
            ty = y + ph + 0.2
            cwt = pw / 4
            for k, key in enumerate(keys):
                val = color_by_ref(b, spec.get(key))
                if not val:
                    continue
                kx = x + (k % 4) * cwt
                ky = ty + (k // 4) * 0.62
                rrect(s, kx, ky + 0.04, 0.34, 0.34, val, line=t.border, r=0.17)
                text(s, kx + 0.44, ky, cwt - 0.5, 0.5,
                     [[(key.replace("_", " ").title(), {"bold": True, "color": t.light_text, "size": 10})],
                      val], size=9.5, color=t.light_muted, font=t.body_font, para_after=0)
            if spec.get("proposed"):
                self.proposed_tag(s, x + pw - 1.25, y - 0.38)

    def ratio(self):
        b, t = self.b, self.t
        r = g(b, "colors", "ratio", default=None)
        if not r:
            return
        s = self.full("Color Proportion",
                      subtitle=g(b, "colors", "ratio_intro") or
                      "Roughly how much of a layout each color should occupy.")
        total = sum(float(i.get("percent", 0)) for i in r) or 100
        x = 0.6
        gap = 0.06
        bw = W - 1.2 - gap * (len(r) - 1)
        for i in r:
            c = color_by_ref(b, i.get("color")) or "#999999"
            w = bw * float(i.get("percent", 0)) / total
            rrect(s, x, 2.4, w, 2.8, c, line=t.border if rel_lum(hex_to_rgb(c)) > 0.85 else None,
                  r=0.14)
            ink = readable_on(c)
            text(s, x + 0.25, 2.65, max(w - 0.3, 0.4), 0.9, f"{int(float(i.get('percent', 0)))}%",
                 size=44 if w > 1.6 else 20, color=ink, font=t.head_font)
            text(s, x + 0.02, 5.4, max(w - 0.1, 0.6), 1.1,
                 [[(i.get("label") or i.get("color"), {"bold": True, "color": t.light_text,
                                                          "size": 12})],
                  i.get("use", "")], size=11, color=t.light_muted, font=t.body_font,
                 para_after=1, line_spacing=1.1)
            x += w + gap

    def contrast_matrix(self):
        b, t = self.b, self.t
        modes = g(b, "colors", "modes", default={}) or {}
        bgs, fgs = [], []

        def addu(lst, h, label):
            if h and h not in [x[0] for x in lst]:
                lst.append((h, label))
        for m in ("light", "dark"):
            sp = modes.get(m) or {}
            addu(bgs, color_by_ref(b, sp.get("background")), f"{m} bg")
            addu(bgs, color_by_ref(b, sp.get("surface")), f"{m} surface")
            for k in ("text", "text_muted", "accent", "link"):
                addu(fgs, color_by_ref(b, sp.get(k)), f"{m} {k.replace('_', ' ')}")
        addu(bgs, t.accent, "accent")
        for h in ("#FFFFFF",):
            addu(fgs, h, "white")
        bgs, fgs = bgs[:5], fgs[:8]
        if len(bgs) < 2 or len(fgs) < 2:
            return
        s = self.full("Accessible Color Pairings",
                      subtitle="WCAG 2.x contrast ratios, calculated from the HEX values above. Body text "
                               "needs 4.5:1 (AA); large text and UI elements need 3:1.")
        rows, cols = len(fgs) + 1, len(bgs) + 1
        tw, th = W - 1.2, min(4.2, 0.46 * rows)
        tbl = s.shapes.add_table(rows, cols, Inches(0.6), Inches(2.4), Inches(tw), Inches(th)).table
        tbl.columns[0].width = Inches(2.6)
        for j in range(1, cols):
            tbl.columns[j].width = Inches((tw - 2.6) / (cols - 1))

        def cell(r_, c_, txt, fill, ink, bold=False, size=10):
            ce = tbl.cell(r_, c_)
            ce.fill.solid()
            ce.fill.fore_color.rgb = rgb(fill)
            tf = ce.text_frame
            tf.text = ""
            for k_, part in enumerate(str(txt).split("\n")):
                p = tf.paragraphs[0] if k_ == 0 else tf.add_paragraph()
                p.alignment = PP_ALIGN.CENTER if c_ else PP_ALIGN.LEFT
                rr = p.add_run()
                rr.text = part
                rr.font.size = Pt(size)
                rr.font.bold = bold if k_ == 0 else False
                rr.font.name = t.body_font
                rr.font.color.rgb = rgb(ink)
            ce.vertical_anchor = MSO_ANCHOR.MIDDLE
        cell(0, 0, "Text  \\  Background", t.light_surface, t.light_text, True)
        for j, (h, lab) in enumerate(bgs, 1):
            cell(0, j, f"{lab}\n{h}", h, readable_on(h), True, 9.5)
        for i, (fh, flab) in enumerate(fgs, 1):
            cell(i, 0, f"{flab}  {fh}", t.light_surface, t.light_text, False, 10)
            for j, (bh, _) in enumerate(bgs, 1):
                r_ = contrast(fh, bh)
                lab = wcag_label(r_)
                ce = tbl.cell(i, j)
                ce.fill.solid()
                ce.fill.fore_color.rgb = rgb(bh)
                tf = ce.text_frame
                tf.text = ""
                p = tf.paragraphs[0]
                p.alignment = PP_ALIGN.CENTER
                if r_ < 1.05:      # a color against itself carries no information
                    ce.vertical_anchor = MSO_ANCHOR.MIDDLE
                    continue
                for txt_, col_, sz_ in (("Aa  ", fh, 13), (f"{r_:.1f}:1 {lab}", fh if r_ >= 3 else readable_on(bh), 10)):
                    rr = p.add_run()
                    rr.text = txt_
                    rr.font.size = Pt(sz_)
                    rr.font.bold = True
                    rr.font.name = t.body_font
                    rr.font.color.rgb = rgb(col_)
                ce.vertical_anchor = MSO_ANCHOR.MIDDLE

    # ---------- typography ----------
    def typography(self):
        b, t = self.b, self.t
        T = g(b, "typography", default={}) or {}
        fams = T.get("families") or []
        if not fams:
            return
        self.divider("Typography", T.get("intro") or "Consistent type makes us recognizable before anyone reads a word.")
        for f in fams:
            font = f.get("fallback") if t.fallback_fonts and f.get("fallback") else f.get("name")
            body = [f.get("usage") or ""]
            if f.get("source"):
                body.append(f"Source / license: {f['source']}")
            if f.get("fallback"):
                body.append(f"Fallback when unavailable: {f['fallback']}")
            s = self.content(f.get("name", "Typeface"), [p for p in body if p],
                             eyebrow=f.get("role", "Typeface"))
            text(s, RX, 0.7, 3.6, 3.0, "Aa", size=170, color=t.light_text, font=font,
                 line_spacing=0.9)
            text(s, RX + 3.8, 1.25, RW - 3.8, 2.4,
                 ["ABCDEFGHIJKLM", "NOPQRSTUVWXYZ", "abcdefghijklm", "nopqrstuvwxyz",
                  "0123456789 !@#$%&?"], size=20, color=t.light_text, font=font,
                 line_spacing=1.15, para_after=4)
            weights = f.get("weights") or []
            if weights:
                rule(s, RX, 4.2, RW, t.hair)
                self.label(s, RX, 4.4, RW, "Weights")
                cw = RW / min(4, len(weights))
                for i, w in enumerate(weights[:8]):
                    x = RX + (i % 4) * cw
                    y = 4.85 + (i // 4) * 0.8
                    isbold = any(k in str(w).lower() for k in ("bold", "black", "heavy", "semi"))
                    text(s, x, y, cw - 0.1, 0.7, f"{w}", size=22, bold=isbold,
                         color=t.light_text, font=font, para_after=0)
            if f.get("reserved"):
                text(s, RX, H - 1.35, RW, 0.4, f"Reserved: {f['reserved']}", size=12,
                     color=t.bad, font=t.body_font)

        hier = T.get("hierarchy") or []
        if hier:
            s = self.full("Type Hierarchy", subtitle=T.get("hierarchy_intro"))
            n = len(hier[:9])
            rh = min(0.66, 4.5 / n)
            for i, h in enumerate(hier[:9]):
                y = 2.2 + i * rh
                fam = h.get("family") or t.body_font
                fobj = next((f for f in fams if f.get("name") == fam), {})
                font = fobj.get("fallback") if (t.fallback_fonts and fobj.get("fallback")) else fam
                try:
                    sz = float(str(h.get("size", "16")).lower().replace("px", "").replace("pt", "").split("/")[0])
                except ValueError:
                    sz = 16
                disp = max(11, min(34, sz * 0.75))
                rule(s, 0.6, y, W - 1.2, t.hair)
                text(s, 0.6, y, 1.2, rh, h.get("level", "").upper(), size=10, bold=True,
                     color=t.light_muted, font=t.body_font, anchor="middle", spacing=1)
                wt = str(h.get("weight", "")).lower()
                sample = h.get("sample") or ("The quick brown fox" if sz < 24 else "Headline sample")
                text(s, 1.9, y, 5.9, rh, sample, size=disp, color=t.light_text, font=font,
                     bold=any(k in wt for k in ("bold", "semi", "black", "600", "700", "800")),
                     anchor="middle", caps=str(h.get("case", "")).lower() == "uppercase")
                spec = "  |  ".join(str(x) for x in [fam, h.get("weight"), h.get("size"),
                                                      (f"LH {h['line_height']}" if h.get("line_height") else None),
                                                      (f"Tracking {h['tracking']}" if h.get("tracking") else None),
                                                      h.get("case")] if x)
                text(s, 8.0, y, W - 8.6, rh, spec, size=10.5, color=t.light_muted,
                     font=t.body_font, anchor="middle")
        if T.get("rules") and T.get("donts"):
            self.list_slide("Typography Rules", T["donts"], dos=T["rules"])
        elif T.get("rules") or T.get("donts"):
            self.list_slide("Typography Rules", T.get("rules") or T.get("donts"),
                            kind="rule" if T.get("rules") else "dont")

    # ---------- imagery ----------
    def imagery(self):
        b, t = self.b, self.t
        I = g(b, "imagery", default={}) or {}
        ph, ic, ge = I.get("photography"), I.get("iconography"), I.get("graphic_elements")
        il = I.get("illustration")
        if not (ph or ic or ge or il):
            return
        self.divider("Imagery", I.get("intro") or "Photography, illustration, and icons that look and feel like us.")
        for key, title in (("photography", "Photography"), ("illustration", "Illustration"),
                           ("iconography", "Iconography"), ("graphic_elements", "Graphic Elements")):
            sec = I.get(key)
            if not sec:
                continue
            if isinstance(sec, list):
                sec = {"principles": sec}
            pr = sec.get("principles") or []
            bul = [(p if isinstance(p, str) else f"{p.get('title')}: {p.get('text')}") for p in pr]
            dos = sec.get("do_images") or []
            if not (dos or sec.get("dont_images") or sec.get("examples")):
                if sec.get("dont"):
                    self.list_slide(title, sec["dont"], intro=sec.get("summary"), dos=bul)
                else:
                    self.list_slide(title, bul, intro=sec.get("summary"))
                continue
            s = self.content(title, sec.get("summary"), bullets=bul[:7])
            donts = sec.get("dont_images") or []
            ex = sec.get("examples") or []
            allimg = [(p, True) for p in dos] + [(p, False) for p in donts] + [(p, None) for p in ex]
            allimg = [(resolve(b, p), ok) for p, ok in allimg if resolve(b, p) and os.path.exists(resolve(b, p))]
            if allimg:
                n = min(6, len(allimg))
                cols = 3 if n > 4 else 2 if n > 1 else 1
                rows = (n + cols - 1) // cols
                cw = (RW - 0.25 * (cols - 1)) / cols
                ch = (H - 1.8 - 0.25 * (rows - 1)) / rows
                for i, (p, ok) in enumerate(allimg[:n]):
                    x = RX + (i % cols) * (cw + 0.25)
                    y = 0.9 + (i // cols) * (ch + 0.25)
                    im = Image.open(p).convert("RGB")
                    tr = cw / ch
                    if im.width / im.height > tr:
                        nw = int(im.height * tr)
                        im = im.crop(((im.width - nw) // 2, 0, (im.width - nw) // 2 + nw, im.height))
                    else:
                        nh = int(im.width / tr)
                        im = im.crop((0, (im.height - nh) // 2, im.width, (im.height - nh) // 2 + nh))
                    picture(s, im, x, y, cw, ch, radius=0.12)
                    if ok is not None:
                        self._x_badge(s, x + 0.08, y + 0.08, ok=ok)
            elif sec.get("dont"):
                self.label(s, RX, 0.95, RW, "Avoid", color=t.bad)
                text(s, RX, 1.35, RW, 4, sec["dont"], size=14, color=t.light_text,
                     font=t.body_font, bullets=True, para_after=7)
            if sec.get("dont") and allimg:
                pass

    # ---------- editorial ----------
    def editorial(self):
        b, t = self.b, self.t
        E = g(b, "editorial", default={}) or {}
        if not E:
            return
        self.divider("Writing Style", E.get("intro") or "How we write about ourselves and our products.")
        names = E.get("name_usage") or []
        if names:
            self.list_slide("Company and Product Names", names)
        terms = E.get("terms") or []
        if terms:
            for pg in range(0, len(terms), 10):
                chunk = terms[pg:pg + 10]
                s = self.full("Preferred Terms" + (" (continued)" if pg else ""))
                rows = len(chunk) + 1
                tbl = s.shapes.add_table(rows, 3, Inches(0.6), Inches(1.9), Inches(W - 1.2),
                                         Inches(min(4.6, 0.46 * rows))).table
                tbl.first_row = False
                tbl.horz_banding = False
                tbl.columns[0].width = Inches(3.2)
                tbl.columns[1].width = Inches(3.2)
                tbl.columns[2].width = Inches(W - 1.2 - 6.4)
                for j, hh in enumerate(["Use", "Instead of", "Why"]):
                    self._cell(tbl, 0, j, hh.upper(), t.light_bg, t.light_muted, True, size=10,
                               line=t.light_text)
                for i, tr in enumerate(chunk, 1):
                    self._cell(tbl, i, 0, tr.get("use", ""), t.light_bg, t.light_text, True)
                    self._cell(tbl, i, 1, tr.get("avoid", ""), t.light_bg, t.bad)
                    self._cell(tbl, i, 2, tr.get("note", ""), t.light_bg, t.light_muted)
        fmt = E.get("formatting") or []
        if fmt:
            self.list_slide("Formatting Rules", fmt,
                            intro=E.get("base_style_guide") and
                            f"For anything not covered here, follow {E['base_style_guide']}.")
        if E.get("donts"):
            self.list_slide("Writing: What Not To Do", E["donts"], kind="dont")

    def _cell(self, tbl, r_, c_, txt, fill, ink, bold=False, size=13, line=None):
        """Table cell with a hairline underneath instead of a filled grid."""
        ce = tbl.cell(r_, c_)
        ce.fill.solid()
        ce.fill.fore_color.rgb = rgb(fill)
        tf = ce.text_frame
        tf.text = ""
        rr = tf.paragraphs[0].add_run()
        rr.text = str(txt)
        rr.font.size = Pt(size)
        rr.font.bold = bold
        rr.font.name = self.t.body_font
        rr.font.color.rgb = rgb(ink)
        ce.vertical_anchor = MSO_ANCHOR.MIDDLE
        ce.margin_left = Inches(0.02)
        tcPr = ce._tc.get_or_add_tcPr()
        for edge in ("a:lnL", "a:lnR", "a:lnT", "a:lnB"):
            for el in tcPr.findall(qn(edge)):
                tcPr.remove(el)
        for edge in ("a:lnL", "a:lnR", "a:lnT"):
            ln = etree.SubElement(tcPr, qn(edge), w="0")
            etree.SubElement(ln, qn("a:noFill"))
        ln = etree.SubElement(tcPr, qn("a:lnB"), w="9525")
        sf = etree.SubElement(ln, qn("a:solidFill"))
        etree.SubElement(sf, qn("a:srgbClr"), val=norm_hex(line or self.t.hair)[1:])
        # border elements must precede the fill element in tcPr
        fills = [el for el in tcPr if el.tag in (qn("a:solidFill"), qn("a:noFill"))]
        for el in fills:
            tcPr.remove(el)
            tcPr.append(el)

    # ---------- governance ----------
    def governance(self):
        b, t = self.b, self.t
        G = g(b, "governance", default={}) or {}
        checklist = list(G.get("checklist") or [])
        auto = ["Is this the current logo file from the official source?",
                "Is the logo version right for the background (light vs. dark)?",
                "Is the logo clear space respected and above minimum size?",
                "Are all colors from the approved palette, using exact values?",
                "Do text and background pairs meet contrast guidelines?",
                "Are the brand typefaces (or approved fallbacks) used?"]
        if g(b, "editorial"):
            auto.append("Are company and product names written correctly?")
        if g(b, "imagery", "photography"):
            auto.append("Does photography match the brand direction?")
        if g(b, "cobranding"):
            auto.append("Are partner logos balanced and spaced per co-branding rules?")
        items = checklist or auto
        self.divider("Using This Guide")
        s = self.full("Brand Checklist", subtitle="Run through this before anything ships.")
        half = (len(items) + 1) // 2
        for col, chunk in enumerate((items[:half], items[half:])):
            for i, it in enumerate(chunk[:7]):
                x = 0.6 + col * 6.2
                y = 2.1 + i * 0.64
                rrect(s, x, y + 0.04, 0.3, 0.3, None, line=t.accent, r=0.09)
                text(s, x + 0.52, y, 5.4, 0.6, it, size=14, color=t.light_text,
                     font=t.body_font, line_spacing=1.1)
        contacts = G.get("contacts") or []
        hist = G.get("version_history") or []
        if contacts or G.get("asset_location") or hist:
            s = self.content("Questions and Resources",
                             G.get("approval_note") or "When in doubt, ask before you publish.")
            y = 0.95
            for c in contacts[:4]:
                rule(s, RX, y, RW, t.hair)
                text(s, RX, y + 0.15, RW, 0.8,
                     [[(c.get("role", ""), {"bold": True, "color": t.light_muted, "size": 11})],
                      [("   ".join(x for x in [c.get("name"), c.get("email")] if x), {"size": 17})]],
                     size=15, color=t.light_text, font=t.body_font, para_after=3)
                y += 1.05
            if G.get("asset_location"):
                rule(s, RX, y, RW, t.hair)
                text(s, RX, y + 0.15, RW, 0.9,
                     [[("Logo files, fonts, and templates", {"bold": True, "color": t.light_muted, "size": 11})],
                      [(G["asset_location"], {"size": 15})]],
                     size=15, color=t.light_text, font=t.body_font, para_after=3)
                y += 1.15
            if hist:
                rule(s, RX, y, RW, t.hair)
                self.label(s, RX, y + 0.17, RW, "Version history")
                text(s, RX, y + 0.55, RW, 2, [f"{h.get('version')}   {h.get('date')}   {h.get('notes', '')}"
                                              for h in hist[:5]],
                     size=12, color=t.light_muted, font=t.body_font, para_after=4)

    # ==================================================================
    def build(self, out):
        b = self.b
        self.cover()
        self.purpose()
        sections = []
        if any(g(b, "foundation", k) for k in ("vision", "positioning", "boilerplate", "pillars", "values", "voice", "direction")):
            sections.append("Brand Foundation")
        if logo_variants(b):
            sections.append("Logo")
        if g(b, "colors", "groups"):
            sections.append("Color")
        if g(b, "typography", "families"):
            sections.append("Typography")
        if any(g(b, "imagery", k) for k in ("photography", "iconography", "graphic_elements", "illustration")):
            sections.append("Imagery")
        if g(b, "editorial"):
            sections.append("Writing Style")
        sections.append("Using This Guide")
        self.toc_slide(sections)
        self.mission()
        self.foundation()
        self.logos()
        self.architecture()
        self.cobranding()
        self.colors()
        self.typography()
        self.imagery()
        self.editorial()
        self.governance()
        self.mission(closing=True)
        self.finish()
        self.prs.core_properties.title = f"{self.company} Brand Guidelines"
        self.prs.core_properties.author = self.company
        self.prs.save(out)
        return self.warnings


TILE_RATIO = misuse.TILE[1] / misuse.TILE[0]


def clear_space_diagram(logo, ratio, bg, accent, unit_label="x"):
    """Logo with exclusion zone drawn around it. ratio = x / logo height."""
    target_h = 240
    im = logo.copy()
    sc = target_h / im.height
    im = im.resize((int(im.width * sc), target_h), Image.LANCZOS)
    x = int(target_h * ratio)
    pad = x + 70
    Wd, Hd = im.width + pad * 2, im.height + pad * 2
    c = Image.new("RGBA", (Wd, Hd), hex_to_rgb(bg) + (255,))
    d = ImageDraw.Draw(c)
    ac = hex_to_rgb(accent)
    ox, oy = pad - x, pad - x
    zone = [ox, oy, ox + im.width + 2 * x, oy + im.height + 2 * x]
    zfill = tuple(int(v * 0.15 + 255 * 0.85) for v in ac) if rel_lum(hex_to_rgb(bg)) > 0.5 \
        else tuple(int(v * 0.25) for v in ac)
    d.rectangle(zone, fill=zfill + (255,))
    for k in range(zone[0], zone[2], 18):
        d.line([(k, zone[1]), (k + 9, zone[1])], fill=ac + (255,), width=3)
        d.line([(k, zone[3]), (k + 9, zone[3])], fill=ac + (255,), width=3)
    for k in range(zone[1], zone[3], 18):
        d.line([(zone[0], k), (zone[0], k + 9)], fill=ac + (255,), width=3)
        d.line([(zone[2], k), (zone[2], k + 9)], fill=ac + (255,), width=3)
    c.alpha_composite(im, (pad, pad))
    # x squares on the four sides
    sq = ac + (90,)
    ov = Image.new("RGBA", c.size, (0, 0, 0, 0))
    od = ImageDraw.Draw(ov)
    cx, cy = pad + im.width // 2, pad + im.height // 2
    for bx, by in [(cx - x // 2, oy), (cx - x // 2, pad + im.height),
                   (ox, cy - x // 2), (pad + im.width, cy - x // 2)]:
        od.rectangle([bx, by, bx + x, by + x], fill=sq, outline=ac + (255,), width=2)
    c = Image.alpha_composite(c, ov)
    d = ImageDraw.Draw(c)
    try:
        f = misuse._font(max(18, x // 2))
        for bx, by in [(cx - x // 2, oy), (ox, cy - x // 2)]:
            d.text((bx + x // 2, by + x // 2), unit_label, fill=ac + (255,), font=f, anchor="mm")
    except Exception:  # noqa
        pass
    return c.convert("RGB")


def gradient_image(stops, angle=90, size=(800, 300)):
    import math
    w, h = size
    cols = [hex_to_rgb(s) for s in stops]
    a = math.radians(float(angle or 90))
    dx, dy = math.sin(a), -math.cos(a)
    im = Image.new("RGB", size)
    px = im.load()
    proj = [((x - w / 2) * dx + (y - h / 2) * dy) for x in (0, w) for y in (0, h)]
    lo, hi = min(proj), max(proj)
    for y in range(h):
        for x in range(0, w):
            tt = (((x - w / 2) * dx + (y - h / 2) * dy) - lo) / (hi - lo)
            seg = min(int(tt * (len(cols) - 1)), len(cols) - 2)
            lt = tt * (len(cols) - 1) - seg
            c1, c2 = cols[seg], cols[seg + 1]
            px[x, y] = tuple(int(c1[k] + (c2[k] - c1[k]) * lt) for k in range(3))
    return im


def main():
    if len(sys.argv) < 3:
        print(__doc__)
        sys.exit(2)
    brand = load_brand(sys.argv[1])
    os.makedirs(os.path.dirname(os.path.abspath(sys.argv[2])), exist_ok=True)
    deck = Deck(brand, fallback_fonts="--fallback-fonts" in sys.argv)
    warns = deck.build(sys.argv[2])
    print(f"Wrote {sys.argv[2]} ({len(deck.prs.slides)} slides).")
    for w_ in warns:
        print("WARN:", w_)


if __name__ == "__main__":
    main()
