# brand.json reference

`brand.json` is the single source of truth. Every script reads it; nothing
is typed twice. Paths are relative to the folder that holds brand.json.
Leave a field out rather than filling it with placeholder text: empty
sections are skipped in the deck.

Colors may be referenced anywhere by HEX (`"#1F5EFF"`) or by swatch name
(`"Halyard Blue"`). Add `"proposed": true` to any swatch, logo variant, or
mode that Claude suggested and the user has not yet approved; the deck tags
it PROPOSED.

A complete working example is in `assets/example/brand.json`.

```jsonc
{
  "meta": {
    "company": "Acme",                    // required, exact spelling
    "guide_title": "Brand Guidelines",
    "subtitle": "optional line under the title",
    "version": "1.0",                     // strongly recommended
    "date": "October 2026",
    "purpose": "Who this guide is for and how strict it is.",
    "audience": "Employees, agencies, partners.",
    "cover_image": "images/cover.jpg"     // optional hero, darkened behind the title
  },
  "foundation": {
    "mission": "", "vision": "", "tagline": "", "positioning": "", "boilerplate": "",
    "direction": "Where the company is going.", "direction_public": true,
    "pillars": [{"name": "", "description": "", "points": [""]}],
    "values": [{"name": "", "description": ""}],
    "voice": {
      "summary": "",
      "traits": [{"we_are": "Direct", "we_are_not": "Blunt"}],
      "tone_notes": [""]
    }
  },
  "audiences": [{"name": "", "titles": "", "needs": "", "tone": [""]}],
  "logos": {
    "description": "", "anatomy": [""],
    "preferred": "light",                 // id of the preferred variant
    "variants": [{
      "id": "light", "name": "Full color", "file": "logos/acme.svg",
      "type": "full | stacked | mark | wordmark | mono",
      "use_on": "light | dark | color | photo",
      "background": "optional color for its tile",
      "usage": "When to use it."
    }],
    "backgrounds": [{"variant": "light", "background": "White", "status": "preferred | approved | minimal", "note": ""}],
    "clear_space": {"ratio": 0.5, "unit_description": "half the mark height",
                    "rule": "sentence shown in the guide", "small_size_rule": ""},
    "min_size": {"print": "1 in", "digital": "120 px", "mark_only": "16 px"},
    "trademark": "® required on first use; dropped below minimum size",
    "donts_auto": ["stretch", "rotate", "recolor", "effects", "outline", "busy",
                   "wrong_mode", "crowd", "low_contrast", "sentence", "squash"],
    "donts_custom": ["text rule", {"text": "rule with example", "image": "dont/x.png"}],
    "where_not": ["contexts where the logo must never appear"],
    "lockup_rules": [""], "lockup_examples": ["images/lockup.png"]
  },
  "architecture": {"intro": "", "tiers": [{"name": "Master brand", "rule": "", "examples": [""]}]},
  "cobranding": {
    "intro": "", "separator": "line | plus | dotted | none", "min_gap": "0.5 in / 48 px",
    "rules": [{"title": "Order", "text": ""}], "donts": [""],
    "dont_images": [{"image": "dont/cobrand.png"}]
  },
  "colors": {
    "intro": "",
    "groups": [{"name": "Primary", "usage": "",
                "swatches": [{"name": "", "hex": "#000000", "rgb": "0, 0, 0",
                              "cmyk": "0, 0, 0, 100", "pms": "Black 6 C", "role": ""}]}],
    "modes": {
      "light": {"background": "", "surface": "", "text": "", "text_muted": "",
                "border": "", "accent": "", "link": "", "logo_variant": "light"},
      "dark":  {"...": "same keys"}
    },
    "modes_intro": "",
    "ratio": [{"color": "White", "percent": 60, "label": "", "use": ""}],
    "gradients": [{"name": "", "stops": ["#...", "#..."], "angle": 135}],
    "gradient_rules": "",
    "donts": [""]
  },
  "typography": {
    "intro": "",
    "families": [{"name": "Inter", "role": "headline | body | mono | accent",
                  "weights": ["Regular", "Bold"], "usage": "", "source": "",
                  "fallback": "Arial", "reserved": "optional: e.g. logo lockups only"}],
    "hierarchy": [{"level": "H1", "family": "Inter", "weight": "Semi Bold", "size": "48px",
                   "line_height": "1.1", "tracking": "0", "case": "sentence", "sample": ""}],
    "hierarchy_intro": "", "rules": [""], "donts": [""]
  },
  "imagery": {
    "photography":  {"summary": "", "principles": [{"title": "", "text": ""}], "dont": [""],
                     "do_images": [""], "dont_images": [""], "examples": [""]},
    "illustration": {"...": "same shape"},
    "iconography":  {"...": "same shape"},
    "graphic_elements": {"...": "same shape"}
  },
  "editorial": {
    "intro": "", "base_style_guide": "the AP Stylebook",
    "name_usage": [""], "terms": [{"use": "", "avoid": "", "note": ""}],
    "formatting": [""], "donts": [""]
  },
  "governance": {
    "asset_location": "", "approval_note": "",
    "contacts": [{"role": "", "name": "", "email": ""}],
    "checklist": ["optional; otherwise generated"],
    "version_history": [{"version": "", "date": "", "notes": ""}]
  }
}
```

## Misuse generators (`logos.donts_auto`)

Each key renders the brand's own logo misused, on a tile with a red X:

| Key | Shows |
|---|---|
| stretch / squash | Distorted proportions |
| rotate | Logo at an angle |
| recolor | Hue-shifted or tinted logo |
| effects | Drop shadow |
| outline | Stroke around the logo |
| busy | Logo on a noisy pattern made from brand colors |
| wrong_mode | Light-mode logo on the dark background (needs a dark variant) |
| low_contrast | Logo on a palette color with similar luminance |
| crowd | Text and blocks inside the clear space |
| sentence | Logo used as a word inside a sentence |

Generators work best on a transparent PNG or SVG. On an opaque JPG the
results look like a box being distorted; ask for a better file.
