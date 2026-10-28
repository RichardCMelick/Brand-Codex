# Interview Question Bank

Ask in small batches (three to five questions per turn), one module at a
time. Use multiple-choice questions where the answer space is known, and
free text where it is not. Before each module, say in one line what it
covers and why. Skip any question already answered by an uploaded guide or
earlier answer; instead, read the recorded value back and ask the user to
confirm it.

Every module ends with a **"what does not work"** batch. Do not skip it.
The negative rules are the part of a guide that prevents real damage, and
they are the part people forget to write down.

---

## Module 0: Kickoff and existing materials

1. Do you have an existing brand guide, style guide, or brand book? Upload
   it (PDF, deck, doc, or a link you can paste). *If yes, extract every
   value you can into brand.json first, run the validator, and then only ask
   about gaps and conflicts.*
2. Upload every logo file you have: all versions, colors, and formats. SVG
   is best, then transparent PNG at least 1000 px wide. Include the mark on
   its own, favicons, and old versions you want to retire (label them).
3. Who will read this guide? (Employees / agencies and freelancers /
   partners and resellers / press / AI design tools / all of these)
4. How strict should it read? (Strict rules, no exceptions / Firm rules with
   room for creativity / Lightweight starter guide)
5. Do you have brand fonts as files, and are they licensed for sharing with
   agencies? (Yes / Yes but internal only / They are Google or open-source
   fonts / Not sure)
6. Anything else to include: photography samples, icon sets, templates,
   mascot or illustration files, an existing deck to match?

## Module 1: Company, direction, and foundation

1. Company name exactly as it should be written (capitalization, spacing,
   legal suffix or not).
2. Mission: one sentence on why you exist.
3. Vision: what the world looks like if you succeed. (Optional.)
4. Where is the company going in the next one to three years? What are the
   goals that the brand has to support? *Ask whether this can appear in a
   public guide; record `direction_public`.*
5. Positioning: who you serve, what you do, and what sets you apart.
6. Approved boilerplate ("About us" paragraph). If none exists, flag it as a
   gap; do not write one unless the user asks, and mark it proposed if so.
7. Brand pillars or values: up to four, each with a line of proof.
8. Tagline (external) and mantra (internal), if any.

## Module 2: Voice and audience

1. Describe the brand's voice in three to five words.
2. For each, what is the overdone version you never want? ("Confident, not
   arrogant.") This becomes the we-are / we-are-not table.
3. Primary audiences and how tone changes for each (an executive buyer vs. a
   hands-on practitioner, for example).
4. **What does not work:** words, phrases, or tactics that are off-brand
   (fear-based selling, hype, jargon, competitor bashing, slang, emojis).

## Module 3: Logo system

Run `scripts/inspect_assets.py` on the uploaded files first and present
what it found: each file, which background it seems built for, and its
apparent type. Ask the user to confirm or correct, rather than asking them
to describe files you can already see.

1. Which version is primary/preferred?
2. For each file: name, which backgrounds it goes on (light / dark / brand
   color / photography), and when to use it.
3. Is there a mark-only version? When may it be used alone (favicon,
   avatar, app icon, small spaces)?
4. Is there a stacked version? When is it preferred over horizontal?
5. Clear space rule: what unit defines it (for example, the height of a
   letter or the mark) and how many of those units? If unknown, ask whether
   to record "not defined" or propose a rule marked PROPOSED.
6. Minimum sizes for print and digital, per version.
7. Trademark symbol: required, where, and when it can be dropped?
8. Are there product, department, or event lockups? Who may create them?
9. Is there a parent brand, sub-brands, or product brands with their own
   logos? (Leads to the architecture section.)

## Module 4: Logo "what does not work"

Show the standard list and ask which apply (multi-select). Default is all:

- Stretch or squash / rotate / recolor / add effects (shadow, glow) /
  outline / busy backgrounds / wrong version for the background / crowding
  clear space / logo used as a word in a sentence / redraw or retype /
  separate mark from wordmark / add taglines to the logo / outdated versions

Then dig for the brand-specific ones. These are the most valuable answers in
the whole interview:

1. Where should the logo **never** appear? Think about: competitor events
   or materials, certain merchandise, political or controversial content,
   specific imagery (for example, the animal or object the brand is named
   after, breach-news photos), partner products without an agreement, user
   generated content, social avatars of non-official accounts.
2. Has the logo been misused before? How? (Each answer becomes a rule.
   Ask for a screenshot to show as a "don't" example.)
3. Is any part of the logo special that people tend to break? (A line that
   only appears in the full lockup, a color that must not change, a symbol
   that must always be present.)
4. Any old logos or names that must no longer be used?

## Module 5: Co-branding and partners (skip if no partners)

1. Do partners, resellers, sponsors, or customers display your logo?
2. Order: whose logo goes first in whose materials?
3. Size relationship: equal visual weight, or does one lead?
4. Separator: vertical rule, plus sign, dotted rule, or none? *Real guides
   disagree on this, so never assume.*
5. Minimum gap between logos.
6. Mode matching: must both logos be in the same mode (both knockout)?
7. "Powered by" or certification badges: who may use them, and is approval
   required?
8. **What does not work:** grayscale conversion, combined lockups, uneven
   sizing, recreating your branding in partner materials, anything else.

## Module 6: Color

Ask for exact values; never derive print values. If the user only has HEX,
record CMYK/Pantone as missing and tell them to get print values from their
designer or printer.

1. Primary colors: name, HEX, RGB, CMYK, Pantone.
2. Secondary or accent colors, and what each is for (CTAs, charts,
   highlights, product color).
3. Neutrals: grays and off-whites, and their roles (backgrounds, borders,
   body text).
4. Functional colors: success, warning, error, info, severity levels.
5. Gradients: which colors, which angles, and which directions are
   forbidden.
6. Roughly what proportion of a typical layout does each color take?
   (For example 60% white, 30% dark, 10% brand color.)

### Light mode and dark mode

Ask these as two parallel sets. If the brand has no dark mode, record that.

1. Page background color.
2. Card or surface color (if different).
3. Primary text and secondary (muted) text colors.
4. Border/divider color.
5. Accent / primary button color, and the text color on that button.
6. Link color.
7. Which logo version is used.
8. Are there colors that only exist in one mode (for example, an accessible
   "light-mode blue" and "dark-mode blue")?

After answering, run the validator and **show the contrast results**. If a
pair fails WCAG, say so plainly and ask whether to keep it, adjust it, or
restrict it to large text. Do not silently fix it.

### Color "what does not work"

1. Colors that must never be used together.
2. Colors that must never be used for text, or as backgrounds.
3. Tints, shades, or opacity changes: allowed or not?
4. Competitor colors to avoid.

## Module 7: Typography

1. Typeface for headlines/display, and for body. Any third (code, data,
   accent)?
2. Weights allowed for each.
3. Where to get them (Google Fonts, Adobe Fonts, licensed foundry) and
   whether they can be shared.
4. Fallback font when the brand font is unavailable (email, Google Slides,
   Canva without upload, PowerPoint on other machines).
5. Type scale: for eyebrow, H1 to H4, lead, body, caption, quote: family,
   weight, size, line height, letter spacing, and case.
6. Case rule: sentence case, title case, or uppercase for headlines?
7. Is any font reserved (for the logo or lockups only)?
8. **What does not work:** fonts to avoid, minimum sizes, weights not to
   mix, alignment (justified text), all caps body, etc.

## Module 8: Imagery

1. Photography: subject, light, color treatment, composition. Upload three
   to six approved examples.
2. Photos to avoid (and upload a couple of "don't" examples if possible).
3. Illustration or mascot: when it appears, and when it must not.
4. Icons: style (line, solid, duotone), grid, stroke, colors, size range,
   and where the library lives.
5. Graphic devices: patterns, shapes, textures, and their rules.

## Module 9: Writing style (optional but recommended)

1. How company and product names are written (capitalization, trademark
   on first use, never prefixing products with the company name).
2. Preferred terms and the terms they replace.
3. Formatting: Oxford comma, ampersands, dates, times, phone numbers,
   numbers, currency.
4. Base style guide for everything else (AP, Chicago, or a house guide).
5. Claims policy: superlatives, statistics, citations.

## Module 10: Governance

1. Who approves exceptions, and how to reach them?
2. Where do people download logos, fonts, and templates?
3. Version number and date for this edition.
4. Anything to add to the final pre-publish checklist?

---

## Reading an existing guide

When the user uploads a guide, extract into brand.json in this order:
meta, foundation, logos (note each variant shown and its background),
colors (every swatch with every value printed), modes, typography,
imagery, co-branding, editorial, and every "do not" example. Then:

1. Run `scripts/validate_brand.py`. Professionally produced guides routinely
   contain mismatched HEX/RGB values, copy-pasted Pantone codes, and
   contradictory font names. Report each finding and ask which value is
   correct. Do not pick one silently.
2. Tell the user which sections were missing from their guide and ask only
   about those.
3. Ask for logo source files even if the guide shows the logo: images
   cropped from a PDF are not good enough to publish.
