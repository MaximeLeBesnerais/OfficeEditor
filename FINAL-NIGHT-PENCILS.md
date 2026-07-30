# Night Pencils — Final Report

Branch: `wip/night-pencils` · 2026-07-28

## Fixes

Commits:

- `07aba73` — preserve inherited layout/master content, correctly scope
  relationship lookups, emit zero-dimension connectors as Typst lines, and
  retain shadow subpaths.
- `4403f22` — resolve run/paragraph hyperlink markup through OuterXml fallbacks,
  apply the theme hyperlink color, and emit underline decoration.

The earlier Pencil regression coverage for picture relationship collisions,
picture-placeholder transforms, group rotation, wrap-none, justify, color
inheritance, master shapes, crop/rotation, and sizing remains green. New
synthetic coverage is in `PptxToTypstConverterNightPencilTests.cs` (3 tests).

## Fidelity verification

All 44 slides were rendered and measured with `scripts/rmse.py` against the
supplied PDF references:

| Deck | Mean | P85 | P90 | Worst |
|---|---:|---:|---:|---:|
| Pencil-Standard (4:3) | 5.4% | 5.8% | 6.1% | s22 10.7% |
| Pencil-Wide (16:9) | 6.0% | 6.2% | 6.3% | s22 12.9% |

Visual inspection of every slide found no remaining non-font structural
mismatch. Differences are rasterizer/font anti-aliasing noise. Slide 22's
`www.showeet.com` footer is now theme teal and underlined in both twins.

## Verification

- `dotnet build DocxEditor.sln --no-restore`: 0 warnings, 0 errors.
- `dotnet test DocxEditor.sln --no-build --no-restore`: 1924 passed, 0 failed,
  0 skipped.
- `local-ref/` was not staged or committed.
