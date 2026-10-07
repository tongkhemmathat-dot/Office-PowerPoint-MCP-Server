---
name: ppt-thai-sarabun
description: Build Thai PowerPoint decks and system diagrams in the TH Sarabun New font (Thai government / enterprise document style) with the Office-PowerPoint-MCP-Server tools. Use when the user wants Thai slides, "ฟอนต์ TH Sarabun / สารบรรณ", or a deck that must look like a Thai official document; also load ppt-slides-and-diagrams for the general workflow.
---

# Thai slides in TH Sarabun New

Draft. Extends `ppt-slides-and-diagrams`; everything there still applies. This skill only changes the font and the sizes that follow from it.

## Fonts
- Default: **TH Sarabun New** (family name exactly `TH Sarabun New`). Same design as TH SarabunPSK.
- Official letters (หนังสือราชการ / สารบรรณ): **TH SarabunIT๙** (family name `TH SarabunIT๙`). Use only if the user asks for it.
- Not the same as Google's `Sarabun` (wider, Tahoma-like). If the user says "th sarabun", they mean TH Sarabun New.
- Check it is installed (Windows): `(New-Object System.Drawing.Text.InstalledFontCollection).Families | ? Name -like '*Sarabun*'`. `render_slide` silently falls back to another font if it is missing, so confirm the rendered Thai looks like Sarabun (loops on the heads of letters).
- The font must also exist on the machine that opens the file; otherwise PowerPoint substitutes another font and the layout shifts. For sharing, install the font there or send a PDF.

## Set the font in the template, not per element
Use `templates/SE_Template_Sarabun.pptx` (`create_presentation_from_template`). It sets the theme's Latin, complex-script (`cs`, which PowerPoint uses for Thai) and Thai fonts to TH Sarabun New, so shapes, text boxes, `add_icon` labels, tables and placeholders all inherit it.
- `add_icon` labels and shape text have no font parameter, so a per-element `font_name` is not available there anyway.
- Do not rely on `font_name` for Thai text: it sets the Latin font only (untested for Thai; the theme route is the verified one).
- Rebuild the template with other settings: `python scripts/build_se_template.py --font "TH Sarabun New" --scale 1.3 --out templates/SE_Template_Sarabun.pptx`.

## Sizes (measured against Tahoma at the same pt)
TH Sarabun New is narrow and small: width ≈ 0.67×, x-height ≈ 0.62×, Thai glyph height ≈ 0.70× of Tahoma. The same pt size looks about one third smaller.

| Use | Tahoma (old diagrams) | TH Sarabun New |
|---|---|---|
| Slide title | 26 | 34 (template default) |
| Body bullets | 20 | 24–28 (template: 26) |
| Table body | 11–12 | 16 |
| Table header | 11–13 | 16–18 |
| Card / box titles | 10–11 | 16–18 |
| Spec rows, node labels | 9–10 | 14–16 |
| Notes, footnotes | 8–9 | 13–14 |
| Absolute minimum anywhere | 8 | 13 |

Rule of thumb: Tahoma size × 1.4, round up. At ×1.4 a line is about as wide as before (0.94×), but every line is about 1.4× taller, so give text boxes ~1.3–1.4× more height and expect fewer items per slide.
- `add_icon`: pass `label_font_size` (default 9 is far too small, use 13–14), `label_height` ≥ 0.6 for two lines.
- `add_table`: `header_font_size` / `body_font_size` 16, row height ≥ 0.5 in; 10 rows at 16 pt fill a slide.
- Spec cards, KPI strips and legends designed at 9–10 pt need re-laying out, not just a bigger number.

## Workflow
1. `create_presentation_from_template` with `SE_Template_Sarabun.pptx`; layouts 0-6 as in the default template (5 = Title Only (Diagram / Table), 6 = Blank).
2. Lay out with the larger text from the start (fewer words, wider boxes); do not copy a Tahoma layout and only raise the sizes.
3. `render_slide(return_image=true)` after each slide: check clipped text, wrapped rows in aligned columns (label / value columns must stay on one line each), labels touching icons or borders.
4. Widen boxes before shrinking text. Never go below the minimums above.

## Thai text notes
- Thai has no spaces between words; PowerPoint breaks lines itself, so a narrow box can split a phrase. Widen the box or shorten the phrase; don't insert manual line breaks inside a phrase.
- Mixed Thai/English in one line is fine: both come from the same font.
- Use พ.ศ. for years in Thai documents unless told otherwise; keep Arabic digits for specs and numbers (Thai digits only if asked).
- Keep product names and units in English (ESXi, vCPU, GB, HTTPS 443).

## Gotchas
- Bold: use `bold: true` (TH Sarabun New Bold exists). Italic works but reads poorly in slides.
- Tables: `format_table_cell` may override fonts; keep cell fonts untouched so they inherit the theme font.
- Anything built earlier in Tahoma (e.g. the ArcGIS/HCI decks) does not change automatically; rebuild from the Sarabun template with scaled sizes.
