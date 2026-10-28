---
name: brand-codex
description: Guided interview that assembles a company's existing brand into a publishable brand guidelines deck (PowerPoint) plus a handoff kit for Canva, Claude Design, and other AI tools. Use whenever someone wants to create, document, compile, rebuild, update, or showcase brand guidelines, a brand book, style guide, logo usage rules, or brand standards, or wants to load their brand into Canva or Claude Design, even if they only upload a logo or an old brand PDF and say "help me with our brand guide." Not for inventing a new brand identity from scratch.
license: MIT. Complete terms in LICENSE.txt
---

# Brand Codex

Brand Codex turns a brand that already exists (in someone's head, in a
folder of logo files, or in an outdated PDF) into a clear, public-ready
brand guideline. It does **not** design the brand. It records decisions,
asks for the ones nobody has written down, catches errors, and presents the
result well.

Deliverables:

1. **Brand guidelines deck** (`.pptx`, 16:9), built from the user's real
   logo files and exact values.
2. **Handoff kit**: `brand-guidelines.md`, `design-tokens.json`,
   `tokens.css`, `canva-brand-kit.md`, `claude-design-setup.md`, exported
   logos, and one zip of everything.
3. **brand.json**: the single source of truth, so the guide can be
   regenerated when anything changes.

Read `references/anatomy.md` once before starting. It explains what a
complete brand guideline contains, why each section exists, and the errors
that professionally produced guides ship with. It is the model for
everything below.

## Principles

- **Record, do not invent.** Never make up a color value, CMYK or Pantone
  conversion, clear-space rule, font, or boilerplate. If something is not
  defined, say so and offer two choices: leave it out, or let Claude
  propose a value that is tagged `"proposed": true` (the deck labels it
  PROPOSED) until the user approves it.
- **Every section answers three questions:** what it is, how to use it, and
  how not to use it. The "how not to" answers are what prevent real damage,
  so every interview module ends with a "what does not work" batch.
- **One source of truth.** All values go into `brand.json`. Scripts read it;
  nothing is retyped into slides. This is how the deck avoids the
  mismatched values and contradictory font names found in real guides.
- **Validate before building.** Run the validator after each major module
  and before every build. Report errors plainly and ask which value is
  right. Do not silently "fix" a user's brand.
- **Respect the user's time.** Ask three to five questions per turn, skip
  anything already answered, and read extracted values back for
  confirmation instead of asking from scratch.

## Setup

```bash
pip install python-pptx Pillow cairosvg   # cairosvg only needed for SVG logos
```
LibreOffice (`soffice`) is optional but needed to render the deck for
visual QA. Scripts live in `scripts/` next to this file; run them with the
skill directory's full path.

Create a working folder (for example `brand-work/`) holding `brand.json`,
a `logos/` folder, and any images. Copy uploaded files there; never modify
the user's originals.

## Workflow

Track these as a task list if the environment has one.

### 1. Kickoff (Module 0 in `references/interview.md`)

Explain in two sentences what will happen and what they will get. Then ask
for: an existing guide, all logo files, the guide's audience and
strictness, and font files or font names.

- **Existing guide uploaded:** extract everything into `brand.json` first
  (see "Reading an existing guide" at the end of the interview reference),
  run the validator, present what was found and what is missing, and only
  interview for gaps and conflicts. This is usually 70% faster.
- **No guide:** start `brand.json` from `assets/brand.template.json` and
  interview module by module.

### 2. Logo inventory

Run:
```bash
python scripts/inspect_assets.py brand-work/logos --out brand-work/variants.draft.json
```
It reports each file's size, transparency, colors, which background it
seems made for (by measuring contrast against white and black), and its
likely type. Present this as a short table and ask the user to confirm or
correct each row. Ask for SVG or transparent PNG replacements for any file
flagged opaque or small, because the deck and the misuse examples are built
from these files.

### 3. Interview

Work through `references/interview.md` modules 1 to 10 in order, skipping
modules that do not apply (co-branding with no partners, architecture with
one brand). Write answers into `brand.json` as you go; the field reference
is `references/brand-schema.md`.

Points that need extra care:

- **Company direction and goals** (Module 1): ask, and ask whether it may
  appear in a public guide.
- **Logo misuse** (Module 4): confirm which standard don'ts apply, then dig
  for brand-specific ones: where the logo must never appear, how it has
  been misused before, and which part of it people tend to break. Ask for
  screenshots of past misuse to show as examples.
- **Light and dark mode** (Module 6): collect the full token set for each
  mode (background, surface, text, muted text, border, accent, link, logo
  version). Then run the validator and show the WCAG contrast results. A
  failing pair is the user's decision to keep, change, or restrict.
- **Print values:** if only HEX is known, CMYK and Pantone stay blank and
  the validator flags them. Do not compute them; screen-to-print conversion
  depends on paper and press, and a wrong value is worse than a blank.
- **Co-branding separator:** real guides disagree on whether a divider is
  required or forbidden. Always ask.

### 4. Validate

```bash
python scripts/validate_brand.py brand-work/brand.json
```
Fix every ERROR with the user. Discuss WARNs briefly (one line each, grouped).
INFO lines list empty sections; those sections will be left out of the deck,
which is fine for a lean guide.

### 5. Build the deck

```bash
python scripts/build_deck.py brand-work/brand.json brand-work/out/<company>-brand-guidelines.pptx
```

The deck structure follows `references/anatomy.md`: cover, about this
guide, contents, mission, foundation (positioning, boilerplate, pillars,
direction, voice, audiences), logo (primary, variations, clear space and
minimum size, background pairings, misuse examples generated from the real
logo, brand-specific rules, lockups), architecture, co-branding, color
(palettes with HEX/RGB/CMYK/Pantone, gradients, light and dark mode mockups,
proportion, computed contrast table, color don'ts), typography (specimens,
hierarchy, rules), imagery, writing style, checklist, contacts, and the
mission again as a closing bookend.

Fonts: the deck names the brand fonts. If the people who will open it may
not have them installed, build with `--fallback-fonts` (uses each family's
fallback) or tell the user that PowerPoint will substitute until the fonts
are installed.

### 6. Visual QA (do not skip)

Render and look at every slide:
```bash
soffice --headless --convert-to pdf <deck>.pptx
pdftoppm -r 50 -jpeg <deck>.pdf slide
```
Check for: text overflowing boxes, titles colliding with body text, logo
tiles where the logo is invisible on its background (usually a mislabeled
`use_on`), empty panels, misuse tiles that do not read as mistakes, and
typos. Fix the data or the generator and rebuild. Proofread all user
supplied text for spelling; flag suspected typos to the user rather than
changing their wording.

### 7. Export the handoff kit

```bash
python scripts/export_handoff.py brand-work/brand.json brand-work/out/handoff
```
Produces the Canva setup guide, Claude Design setup guide, the complete
markdown guide (written to be pasted into any AI tool), W3C design tokens,
CSS variables with dark theme, exported logos, and `brand-handoff.zip`.

### 8. Deliver

Send the `.pptx`, `brand-handoff.zip`, and `brand.json`. In the reply:

- One line on what the deck contains (slide count and sections).
- Anything still PROPOSED or missing, as a short list the user can act on.
- How to update later: edit `brand.json` (or ask Claude to) and rerun the
  build and export scripts.

Keep `brand.json` with the deliverables; it is what makes the next version
a five-minute job instead of a redesign.

## Editing an existing Brand Codex guide

If the user returns with a `brand.json`, load it, make the requested
changes, bump `meta.version` and `meta.date` (ask first), add a
`governance.version_history` entry, validate, rebuild, and re-export.

## Files

| Path | Purpose |
|---|---|
| `references/anatomy.md` | What a brand guideline contains and why; failure patterns |
| `references/interview.md` | Question bank by module, including "what does not work" probes |
| `references/brand-schema.md` | Every brand.json field, and the misuse generator keys |
| `assets/brand.template.json` | Empty starting structure |
| `assets/example/` | Complete fictional example (Halyard) with logos, for testing |
| `scripts/inspect_assets.py` | Analyze uploaded logo files, draft the variants list |
| `scripts/validate_brand.py` | Value consistency, contrast, file checks, and gaps |
| `scripts/build_deck.py` | brand.json to .pptx |
| `scripts/misuse.py` | Generates incorrect-usage examples from the real logo |
| `scripts/export_handoff.py` | Canva, Claude Design, tokens, markdown, zip |
