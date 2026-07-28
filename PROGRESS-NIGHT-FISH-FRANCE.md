# Night Fish / France fidelity pass

## Baseline — 2026-07-28 09:49 CEST

- FISHBONE: mean 8.9%, worst 12.9% (slide 11/23).
- FRANCE: mean 9.2%, worst 14.3% (slide 2/6).
- All 29 reference slides are available under `local-ref/`; the corpus is ignored and will not be staged.

## Completed batches — 2026-07-28

- Preserved regular `prst="line"` connectors as strokes, including group-fill inheritance.
- Preserved shadow color transform chains and color-transform-aware chart fills.
- Routed stacked/percent-stacked bar charts and parsed/rendered doughnut holes.
- Added chart titles, teardrop geometry, SmartArt custom-path connectors, and xfrm flips.
- Added synthetic XML regressions in `PptxToTypstConverterNightFishFranceTests.cs`.

## Verification

- Full `dotnet build` and `dotnet test`: clean (0 warnings/errors; 1,924 tests passed).
- FISHBONE: mean 8.6%, median 8.4%, p90 9.9%; 21/23 below 10%.
- FRANCE: mean 8.5%, median 8.2%, p90 14.1%; 5/6 below 10%.
- Rendered and visually inspected all 29 reference slides. The remaining highest non-font differences are the France slide 2 chart layout, France slide 3 pin silhouette approximation, and residual fishbone text/layout drift.
