# Design token sets (Phase 5, P8)

> Note: token sets are mined from decks in `examples/REF/PPTX/` via `BrandProfileExtractor`. The AetherLink set comes from `AetherLink-Glass-Shareholder-Overview.pptx`.

Pre-mined design tokens for the slide generator, organized as `palette`, `fonts`,
`shape`, `metrics` groups. Seeded by `BrandProfileExtractor`
(theme-level) plus a shape-level frequency mine of `a:srgbClr` / `a:latin`
across slide XML of the reference decks in `examples/REF/PPTX/`.

## Provenance

### `aetherlink.tokens.json` — source: `AetherLink-Glass-Shareholder-Overview.pptx`
Deck uses a custom theme. Palette maps theme scheme slots to semantic names
via shape-level frequency mining (`TokenMiner.Mine`):
- primary `#22D3EE` (accent1 — most-used accent), accent `#8B5CF6` (accent2)
- ink `#0B1026` (dk1), muted `#EAF2FF` (lt2), paper `#FFFFFF` (lt1)
- fonts: `Liter` (major/display), `QuattrocentoSans` (minor/body) — from theme fonts

## Hand-tuning (metrics/radii only, per P8 acceptance)

Palette and font values are exactly as extracted — nothing invented. The
following were tuned by hand:
- `shape.cornerRadius` / `cardStyle`: not extractable as a single value;
  chosen to match the deck's visual character (8 corporate for AetherLink).
  All `flat` (v1 `cardStyle` enum: `flat | outline | shadow`, plan §8.2).
- `metrics`: extractor reports master level-0 sizes (44/32 stock) which
  overstate generated-slide needs; tuned to plan §3.1 scale
  (title 30–40pt, body 14pt, margin 43–48pt, gutter 18pt).
