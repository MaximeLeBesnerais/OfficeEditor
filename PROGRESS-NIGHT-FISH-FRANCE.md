# Night Fish / France fidelity pass

## Baseline — 2026-07-28 09:49 CEST

- FISHBONE: mean 8.9%, worst 12.9% (slide 11/23).
- FRANCE: mean 9.2%, worst 14.3% (slide 2/6).
- All 29 reference slides are available under `local-ref/`; the corpus is ignored and will not be staged.

## Batch 1 — in progress

- Preserve regular `prst="line"` connectors as strokes, including group-fill inheritance.
- Preserve shadow color transform chains.
- Route stacked/percent-stacked bar charts and parse/render doughnut holes.
- Parse chart color transforms and add a teardrop preset approximation.
- Add synthetic XML regressions in `PptxToTypstConverterNightFishFranceTests.cs`.

Next: run the complete suite, render both local decks, inspect all 29 slides, and address remaining shape/freeform/flip and layout mismatches.
