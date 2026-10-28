#!/usr/bin/env python3
"""Inspect uploaded logo files and draft the logos.variants list.

For each image or SVG: pixel size, transparency, dominant colors, whether
the artwork is light or dark (so which background it belongs on), and a
guess at its type from the aspect ratio. The guesses are a starting point
to CONFIRM with the user, never final answers.

Usage: python inspect_assets.py <folder-or-files...> [--out variants.json]
"""

import json
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from brandlib import (load_logo, source_is_opaque, visible_colors,
                      suggest_background)

EXTS = (".png", ".jpg", ".jpeg", ".webp", ".svg")


def guess_type(w, h):
    r = w / float(h)
    if r < 1.35:
        return "mark" if r < 1.15 else "stacked"
    return "full"


def main():
    args = [a for a in sys.argv[1:] if not a.startswith("--")]
    out = None
    if "--out" in sys.argv:
        out = sys.argv[sys.argv.index("--out") + 1]
        args = [a for a in args if a != out]
    files = []
    for a in args:
        if os.path.isdir(a):
            for f in sorted(os.listdir(a)):
                if f.lower().endswith(EXTS):
                    files.append(os.path.join(a, f))
        elif a.lower().endswith(EXTS):
            files.append(a)
    drafts = []
    for p in files:
        try:
            im = load_logo(p)
        except Exception as e:  # noqa
            print(f"- {p}: cannot open ({e})")
            continue
        w, h = im.size
        cols = visible_colors(im, 4)
        transparent = not source_is_opaque(p)
        use_on, cw, cb = suggest_background(im)
        mono = len({c for c in cols}) <= 1
        t = guess_type(w, h)
        if mono:
            t_label = "mono"
        else:
            t_label = t
        base = os.path.splitext(os.path.basename(p))[0]
        print(f"- {os.path.basename(p)}: {w}x{h}px trimmed, "
              f"{'transparent' if transparent else 'OPAQUE background (auto-cut for preview; ask for a transparent file)'}, "
              f"colors {', '.join(cols)}, weakest-element contrast "
              f"{cw:.1f}:1 on white / {cb:.1f}:1 on black "
              f"-> likely for {use_on} backgrounds, looks like: {t_label}")
        drafts.append({
            "id": base.lower().replace(" ", "-"),
            "name": base,
            "file": p,
            "type": t_label,
            "use_on": use_on,
            "usage": "",
        })
    if out:
        with open(out, "w") as f:
            json.dump(drafts, f, indent=2)
        print(f"\nDraft variants written to {out}. Confirm each with the user.")


if __name__ == "__main__":
    main()
