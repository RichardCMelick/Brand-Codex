#!/usr/bin/env python3
"""Validate a brand.json before building the guide.

Catches the errors professional guides ship with: HEX and RGB that disagree,
copy-pasted Pantone/CMYK values, malformed HEX, logo files that do not match
the background they are assigned to, and text/background pairs that fail
WCAG contrast. Also lists the sections that are still empty (gaps).

Usage: python validate_brand.py brand.json [--json]
Exit code 1 if any ERROR is found.
"""

import json
import os
import sys
from collections import defaultdict

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from brandlib import (load_brand, resolve, g, norm_hex, hex_to_rgb, parse_rgb,
                      contrast, wcag_label, all_swatches, color_by_ref,
                      load_logo, source_is_opaque, suggest_background,
                      visible_colors, logo_variants)

ERR, WARN, INFO = "ERROR", "WARN", "INFO"


def color_dist(a, b):
    return sum((x - y) ** 2 for x, y in zip(a, b)) ** 0.5


def main():
    if len(sys.argv) < 2:
        print(__doc__)
        sys.exit(2)
    brand = load_brand(sys.argv[1])
    as_json = "--json" in sys.argv
    issues = []

    def add(level, area, msg):
        issues.append({"level": level, "area": area, "msg": msg})

    # ---- meta ----
    for k in ("company", "version", "date"):
        if not g(brand, "meta", k):
            add(ERR if k == "company" else WARN, "meta",
                f"meta.{k} is missing. A guide without a version and date "
                "cannot be told apart from an outdated copy.")

    # ---- foundation ----
    if not g(brand, "foundation", "mission"):
        add(WARN, "foundation", "No mission statement.")
    if not g(brand, "foundation", "boilerplate"):
        add(INFO, "foundation", "No boilerplate. Press and partners will need one.")

    # ---- colors ----
    sw = all_swatches(brand)
    if not sw:
        add(ERR, "color", "No color swatches defined.")
    pms_map, cmyk_map = defaultdict(list), defaultdict(list)
    seen_hex = {}
    palette_rgb = []
    for grp, s in sw:
        name = s.get("name", "?")
        raw = s.get("hex")
        h = norm_hex(raw)
        if not h:
            add(ERR, "color", f"{grp} / {name}: HEX '{raw}' is malformed "
                "(needs 6 hex digits).")
            continue
        rgb_hex = hex_to_rgb(h)
        palette_rgb.append(rgb_hex)
        if h in seen_hex and seen_hex[h] != name:
            add(WARN, "color", f"{name} and {seen_hex[h]} share HEX {h}. "
                "Intentional alias or copy-paste error?")
        seen_hex[h] = name
        if s.get("rgb") is not None:
            rgb = parse_rgb(s.get("rgb"))
            if not rgb:
                add(ERR, "color", f"{name}: RGB '{s.get('rgb')}' is not three numbers.")
            elif color_dist(rgb, rgb_hex) > 3:
                add(ERR, "color", f"{name}: RGB {rgb} does not match HEX {h} "
                    f"(HEX converts to {rgb_hex}). Confirm which value is correct.")
        if s.get("pms"):
            pms_map[str(s["pms"]).strip().upper()].append((name, rgb_hex))
        if s.get("cmyk"):
            cmyk_map[str(s["cmyk"]).replace(" ", "")].append((name, rgb_hex))
        if not s.get("cmyk") and grp.lower().startswith("primary"):
            add(WARN, "color", f"{name}: no CMYK value. Print vendors will "
                "guess. Get it from the designer, do not compute it.")

    for label, mp in (("Pantone", pms_map), ("CMYK", cmyk_map)):
        for val, items in mp.items():
            if len(items) < 2:
                continue
            spread = max(color_dist(a[1], b[1]) for a in items for b in items)
            if spread > 40:
                add(ERR, "color", f"{label} '{val}' is assigned to visibly "
                    f"different colors: {', '.join(i[0] for i in items)}. "
                    "Likely copy-paste. Confirm print values per swatch.")

    # ---- modes / contrast ----
    modes = g(brand, "colors", "modes", default={}) or {}
    for mode in ("light", "dark"):
        m = modes.get(mode)
        if not m:
            add(WARN, "modes", f"No {mode} mode defined. Ask what backgrounds, "
                "text, and logo variant are used in " + mode + " mode.")
            continue
        bg = color_by_ref(brand, m.get("background"))
        if not bg:
            add(ERR, "modes", f"{mode}.background is missing or not a known color.")
            continue
        checks = [("text", 4.5, "body text"), ("text_muted", 4.5, "secondary text"),
                  ("link", 4.5, "links"), ("accent", 3.0, "UI accents/buttons")]
        for key, need, what in checks:
            c = color_by_ref(brand, m.get(key))
            if not c:
                if key in ("text",):
                    add(WARN, "modes", f"{mode}.{key} not set.")
                continue
            r = contrast(bg, c)
            if r < need:
                add(WARN, "contrast", f"{mode} mode: {key} {c} on background {bg} "
                    f"is {r:.2f}:1 ({wcag_label(r)}). WCAG needs {need}:1 for {what}.")
        for key in ("background", "surface", "text", "text_muted", "border",
                    "accent", "link"):
            if m.get(key) and not color_by_ref(brand, m.get(key)):
                add(ERR, "modes", f"{mode}.{key} = '{m.get(key)}' is not a HEX "
                    "or a swatch name.")

    # ---- logos ----
    vs = logo_variants(brand)
    if not vs:
        add(ERR, "logo", "No logo variants. Upload logo files first.")
    ids = set()
    for v in vs:
        vid = v.get("id", "?")
        ids.add(vid)
        p = resolve(brand, v.get("file"))
        if not p or not os.path.exists(p):
            add(ERR, "logo", f"{vid}: file not found ({v.get('file')}).")
            continue
        try:
            im = load_logo(p)
        except Exception as e:  # noqa
            add(ERR, "logo", f"{vid}: cannot open ({e}).")
            continue
        if not p.lower().endswith(".svg") and max(im.size) < 1000:
            add(WARN, "logo", f"{vid}: only {im.size[0]}x{im.size[1]} px after "
                "trimming. Supply SVG or a PNG at least 1000 px wide.")
        if source_is_opaque(p):
            add(WARN, "logo", f"{vid}: no transparent background. Brand Codex cut the flat "
                "background out automatically, but edges will be rough on colored "
                "backgrounds. Ask for a transparent PNG or SVG.")
        on = v.get("use_on")
        sug, cw, cb = suggest_background(im)
        if on == "light" and sug == "dark":
            add(WARN, "logo", f"{vid}: parts of this artwork reach only {cw:.1f}:1 "
                "contrast on white, but it is assigned to light backgrounds. "
                "Is this actually the knockout version?")
        if on == "dark" and sug == "light":
            add(WARN, "logo", f"{vid}: parts of this artwork reach only {cb:.1f}:1 "
                "contrast on dark, but it is assigned to dark backgrounds. "
                "Is this actually the light-mode version?")
        if palette_rgb and v.get("type", "full") != "mono":
            for c in visible_colors(im, 4):
                rgb = hex_to_rgb(c)
                if min(color_dist(rgb, p_) for p_ in palette_rgb) > 28 and \
                        not (min(rgb) > 235 or max(rgb) < 25):
                    add(WARN, "logo", f"{vid}: logo contains {c}, which is not "
                        "in the palette. Missing swatch or off-brand file?")
                    break
    pref = g(brand, "logos", "preferred")
    if vs and pref and pref not in ids:
        add(ERR, "logo", f"logos.preferred '{pref}' is not a variant id.")
    if vs and not g(brand, "logos", "clear_space"):
        add(WARN, "logo", "Clear space rule not defined.")
    if vs and not g(brand, "logos", "min_size"):
        add(WARN, "logo", "Minimum size not defined.")
    uses = {v.get("use_on") for v in vs}
    if vs and "dark" not in uses:
        add(WARN, "logo", "No logo variant for dark backgrounds.")
    if not (g(brand, "logos", "donts_custom") or g(brand, "logos", "where_not")):
        add(INFO, "logo", "No brand-specific don'ts recorded. Ask where the logo "
            "must never appear; generic don'ts alone miss the risky cases.")

    # ---- typography ----
    fams = g(brand, "typography", "families", default=[]) or []
    if not fams:
        add(ERR, "type", "No typefaces defined.")
    for f in fams:
        if not f.get("fallback"):
            add(WARN, "type", f"{f.get('name')}: no fallback font. Email, Google "
                "Slides, and Canva without upload need one.")
    names = {f.get("name", "").lower() for f in fams}
    for h in g(brand, "typography", "hierarchy", default=[]) or []:
        if h.get("family") and h["family"].lower() not in names:
            add(ERR, "type", f"Hierarchy level '{h.get('level')}' uses "
                f"'{h['family']}', which is not a declared typeface. "
                "(Contradictory font names are a common guide error.)")

    # ---- gaps ----
    sections = {
        "Voice and tone": g(brand, "foundation", "voice"),
        "Imagery / photography": g(brand, "imagery", "photography"),
        "Iconography": g(brand, "imagery", "iconography"),
        "Co-branding": g(brand, "cobranding"),
        "Editorial style": g(brand, "editorial"),
        "Contacts / approvals": g(brand, "governance", "contacts"),
        "Color usage ratio": g(brand, "colors", "ratio"),
    }
    for k, v in sections.items():
        if not v:
            add(INFO, "gaps", f"{k}: empty. The section will be omitted unless filled.")

    # proposed values
    prop = json.dumps(brand).count('"proposed": true')
    if prop:
        add(INFO, "proposed", f"{prop} value(s) marked proposed. They will carry "
            "a PROPOSED tag in the deck until approved.")

    if as_json:
        print(json.dumps(issues, indent=2))
    else:
        order = {ERR: 0, WARN: 1, INFO: 2}
        for i in sorted(issues, key=lambda x: order[x["level"]]):
            print(f"[{i['level']}] {i['area']}: {i['msg']}")
        n = defaultdict(int)
        for i in issues:
            n[i["level"]] += 1
        print(f"\n{n[ERR]} error(s), {n[WARN]} warning(s), {n[INFO]} note(s).")
    sys.exit(1 if any(i["level"] == ERR for i in issues) else 0)


if __name__ == "__main__":
    main()
