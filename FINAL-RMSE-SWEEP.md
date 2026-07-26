# FINAL — SmartArt corpus fidelity campaign (RMSE sweep)

Point-in-time report, branch `dev` @ `7fa251c`. Corpus: `local-ref/smartarts/smartarts.pptx` (164 slides, showeet, local-only) vs official PowerPoint PDF. Method: per-slide RMSE at 2000×1125 (ImageMagick `compare -metric RMSE`), refs via `pdftocairo -r 144`.

## Headline

| Metric | Campaign start | Final |
|---|---|---|
| Mean RMSE | 0.120 | **0.0915** |
| Median | 0.106 | 0.0904 |
| P90 | 0.173 | 0.114 |
| Slides < 0.15 | 142/164 | **163/164** |
| Slides < 0.10 | 56/164 | 114/164 |

Only slide 31 remains above 0.15 (0.168), at its measured residual floor: rotated-text geometry + colorsDef/cached-gradient trade-off (see `FINAL-REPORT-SA-B5A.md` §31).

## Batches (all merged, each with PROGRESS-*/FINAL-REPORT-* at root)

1. **B1 — shared root causes** (`29c6b54`): rotation-aware `ComputeBoundingBox`, shade/satMod + gamma tints, roundRect default adj, bbox pollution skip, presets `flowChartManualOperation`/`homePlate`/`quadArrow`/`blockArc`. Fixed slides 25/30/33/130/131 + 7 bonus.
2. **B2 — presets** (`a91bbfd`): `leftRightRibbon`, `upArrowCallout`, `pie`, `pieWedge` → slides 79/67/87/88/132 under 0.15.
3. **B3 — regression repairs** (`7681d2e`): `nonIsoscelesTrapezoid` + `trapezoid` adj, `wedgeRectCallout` (degenerate-adj rect fallback), dual-fit bbox selection → slides 134/135/49/149/152.
4. **B4 — geometry + text** (`3a6f5ab`): `round2DiagRect`, chevron adj honoring, rotated diagram text (txXfrm/shape rot), fontRef→tx1 fallback (62 slides improved) → 113/58/56/15.
5. **B5 — colors + fonts** (`8d9d395`): `SmartArtColorsDefResolver` (PowerPoint-relayout fill precedence), per-run font chains + system font wiring in convert-pptx, PowerPoint line-pitch model, outerShdw approximation. Corpus mean −0.011 from fonts alone.

## Environment note

Numbers depend on **Carlito + Open Sans** being installed (OFL; installed on the campaign machine by the B5 agent — the corpus declares Open Sans). CI/teammates need both fonts for comparable RMSE.

## Test state

1455 + 280 + 40 + 20 + 36 green (Api has 2 known failures only while the temporary local `smartarts` demo-catalog entry sits uncommitted in `OfficeEditor.Api/Services/DemoDeckService.cs`).

## Remaining known gaps (future batches)

- Slide 31 floor: rotated-text ink geometry + cached-vs-colorsDef gradient residuals
- `outerShdw` is an offset-copy approximation, not a blur
- Drop shadows/effects generally: no `effectDag`, no blur
- Serif fallback anywhere fonts are missing (see environment note)
