# Night Pencils — Progress

Branch: `wip/night-pencils` · 2026-07-28

## Baseline

Measured with `scripts/rmse.py --force` against the supplied PDF renders:

| Deck | Slides | Mean | Worst |
|---|---:|---:|---:|
| Pencil-Standard (4:3) | 22 | 5.4% | s22 10.7% |
| Pencil-Wide (16:9) | 22 | 6.0% | s22 12.9% |

All 44 slides were enumerated. Slides 1–21 are in the expected 4–7% raster/font
noise band; slide 22 is the only recurring structural outlier and was visually
inspected in both twins.

## Fix batch 1

- Layout/master override matching now scopes bounds across shapes, pictures,
  connectors, graphic frames, and groups instead of only `p:sp` shapes.
- Zero-width/zero-height `p:cxnSp` connectors now emit Typst `line` primitives
  with their stroke, preserving vertical and horizontal rules.
- Paragraph-level `a:hlinkClick` is applied to every paragraph run (theme link
  color when available plus underline), with an OuterXml fallback for SDK
  child-model differences.
- Shadow copies preserve multi-subpath geometry, so holes remain holes.
- Added synthetic regression coverage in
  `DocxEditor.Tests/Unit/PptxToTypstConverterNightPencilTests.cs`.

## Verification

- Targeted night-pencil tests: 3 passed.
- `PptxEditor.Core` build: 0 warnings, 0 errors.
- Full build/test and post-fix 44-slide RMSE remain to be run before final.
