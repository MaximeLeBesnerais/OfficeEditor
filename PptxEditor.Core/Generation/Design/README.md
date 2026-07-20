# Design token sets (Phase 5, P8)

Pre-mined design tokens for the slide generator, conforming to plan.md §3.1
(`palette`, `fonts`, `shape`, `metrics`). Seeded by `BrandProfileExtractor`
(theme-level) plus a shape-level frequency mine of `a:srgbClr` / `a:latin`
across slide XML of the reference decks in `examples/REF/PPTX/`.

## Provenance

### `pres-pro.tokens.json` — source: `pres-pro.pptx`
All palette/font values from the deck's custom theme (first master):
- primary `#CEBA80` = theme accent1 (also resolved slide background)
- accent `#169C9A` = theme hlink slot (teal)
- ink `#2C3932` = theme dk2
- paper `#FFFFFF` = theme lt1
- muted `#7A7567` = theme accent3
- fonts: `Aptos Light` (major) / `Aptos` (minor)

### `aetherlink.tokens.json` — source: `AetherLink-Glass-Shareholder-Overview.pptx`
Deck uses the stock Office theme, so values come from shape-level mining:
- primary `#003E7E` (105 uses), accent `#1E7BC6` (39 uses)
- ink `#000000` (33 uses), muted `#343434` (27 uses), paper `#FFFFFF` (24 uses)
- font: `Public Sans` (51 uses; `Public Sans Bold` 77 uses — bold is a run
  property, the family is the token)

## Hand-tuning (metrics/radii only, per P8 acceptance)

Palette and font values are exactly as extracted — nothing invented. The
following were tuned by hand:
- `shape.cornerRadius` / `cardStyle`: not extractable as a single value;
  chosen to match each deck's visual character (0 editorial for pres-pro,
  8 corporate for AetherLink). All `flat`
  (v1 `cardStyle` enum: `flat | outline | shadow`, plan §8.2).
- `metrics`: extractor reports master level-0 sizes (54/18 pres-pro; 44/32
  stock) which overstate generated-slide needs; tuned to plan §3.1 scale
  (title 30–40pt, body 14pt, margin 43–48pt, gutter 18pt).

## Omitted source deck

`FusionFest-Architecture-Overview.pptx` was mined but folded into the set
family rather than shipped as a fourth set: it is another blue deck
(`#005FC2`, `Koho` / `Neo Tech Bold`) with a thin palette (2 mined colors),
redundant next to `aetherlink.tokens.json`. It corroborates the corporate
blue direction.
