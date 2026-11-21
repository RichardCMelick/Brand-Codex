# Changelog

## Unreleased

Deck design pass (Impeccable principles):

- Removed default PowerPoint drop shadows, pill-chip eyebrows, section
  numbers, boxed gray cards, and side-bar accents
- Type scale raised throughout (body 12-15 pt, captions 10 pt minimum,
  titles 34 pt, display up to 72 pt)
- Lists, rules, and tables now sit on hairline rows instead of cards
- Visual areas sit in a softly rounded inset panel; tiles, swatches, and
  images use consistent rounded corners
- Short palettes get one dominant swatch; section dividers and the cover
  are set low with generous space; the contents page shows page numbers
- Status marks (check, cross) are drawn, not font glyphs

## 1.0.0 (October 28, 2026)

First release.

- Guided interview in ten modules, each ending with "what does not work"
- Existing-guide extraction: reads an old brand guide and only asks about gaps
- Logo inventory script that detects transparency, colors, likely
  background, and logo type
- Validator for HEX/RGB mismatches, duplicated Pantone and CMYK codes,
  malformed HEX, undeclared fonts, mislabeled logo files, and WCAG contrast
- PowerPoint generator with logo misuse examples rendered from the real
  logo, clear space diagrams, light and dark mode mockups, and a computed
  contrast table
- Handoff kit: Canva Brand Kit setup, Claude Design setup, W3C design
  tokens, CSS variables, complete markdown guide, and exported logos
- MIT license
- PROPOSED labeling for any value Claude suggests that the user has not
  approved
