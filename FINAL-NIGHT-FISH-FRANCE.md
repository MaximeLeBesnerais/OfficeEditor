# Night Fish / France final report

## Result at deadline

Committed WIP on `wip/night-fish-france`; no merge or push performed.

### Fixes

- Regular line connectors now emit stroked `#line` primitives and inherit group fills.
- Stacked/percent-stacked bars, doughnut holes, chart titles, and transformed chart colors render natively.
- Teardrops, SmartArt custom connector paths, xfrm flips, and shadow transform chains are preserved.
- Added `PptxToTypstConverterNightFishFranceTests.cs` with synthetic chart XML regressions.

### Verification

- Full build/test clean: 0 warnings/errors; 1,924 tests passed.
- FISHBONE (23 slides): mean 8.6%, median 8.4%, p90 9.9%.
- FRANCE (6 slides): mean 8.5%, median 8.2%, p90 14.1%.
- All 29 slides were rendered and visually inspected. Local reference files were not staged.

Remaining visual variance is concentrated in France slide 2 chart layout, France slide 3 pin silhouettes, and residual fishbone text/layout drift.
