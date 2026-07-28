# Night Opposites / Logistics Final Report (WIP)

**Status:** WIP committed at the 12:00 CEST cutoff; not merged or pushed.

## Root cause and fix

`PptxToTypstConverter` flattened custom cubic freeforms with only eight samples per
segment, dropped solid connector lines whose DrawingML width was omitted, and ignored
regular-shape `flipH`/`flipV` flags. These were visible on the required freeform, arrow,
connector, and grouped-icon sweep. The converter now uses 24 curve samples, applies the
default 1pt line width, and reflects normalized points before element rotation.

## Files modified

- `PptxEditor.Core/Converters/PptxToTypstConverter.cs`
- `DocxEditor.Tests/Unit/PptxToTypstConverterNightOppositesLogisticsTests.cs`
- `PROGRESS-NIGHT-OPPOSITES-LOGISTICS.md`
- `FINAL-NIGHT-OPPOSITES-LOGISTICS.md`

## Verification

- `dotnet test DocxEditor.Tests/DocxEditor.Tests.csproj --no-restore`: **1496 passed**.
- Baseline RMSE was captured for both decks and all 44 slides were visually inspected.
- Required post-fix render/RMSE and full-solution build remain pending because the cutoff
  was reached. No `local-ref` files were staged.
