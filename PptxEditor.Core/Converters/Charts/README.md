# Charts Converter

Parses OOXML `c:chartSpace` parts inside `graphicFrame` slide elements and
decomposes them into existing Typst primitives (rects, text, lines, wedge
polygons) so chart content is rendered natively instead of falling back to a
placeholder frame.

## Entry point

`PptxToTypstConverter.Charts.cs` — the `ConvertChartGraphicFrame` partial
method on `PptxToTypstConverter`. It resolves the `c:chart` relationship to
its chart part XML, hands it to `ChartPartParser`, then dispatches bar/column
charts to `BarChartElementBuilder` (clustered, stacked and percent-stacked)
and pie/doughnut charts to `PieChartElementBuilder` (a doughnut keeps the
`c:holeSize` hole transparent). Unsupported chart kinds (line), unreadable or
unparseable chart parts, and charts with no plottable data keep the visible
placeholder frame and raise a warning.

## Chart type support

| Chart kind | Grouping | Supported |
|---|---|---|
| Bar/Column | Clustered | Yes — one rect per point |
| Bar/Column | Stacked / PercentStacked | Yes — series stacked per category; percent-stacked normalises each category to 100% |
| Line | — | Placeholder + warning |
| Pie | — | Yes (wedge polygons, honours `c:firstSliceAng`) |
| Doughnut | — | Yes (pie wedges with the `c:holeSize` hole kept transparent) |

## Files

- **`ChartModel.cs`** — Immutable model of a parsed chart (kind, direction,
  grouping, series, categories, axes, legend, data labels, pie/doughnut
  geometry such as `firstSliceAng`/`holeSize` and per-point fills). All values
  come from cached `strCache`/`numCache` points; cell references (`c:f`) are
  never evaluated.
- **`ChartPartParser.cs`** — `XDocument`-based parser that walks
  `c:chartSpace` XML into a `ChartModel` (`pieChart` and `doughnutChart` both
  map to `ChartKind.Pie`). Uses `XDocument` (not the OpenXML
  SDK) so synthetic chart XML can be unit-tested without a package.
- **`ChartAxisScale.cs`** — Auto value-axis scaling with PowerPoint-style
  nice rounding. Pure static math shared by bar/column builders.
- **`BarChartElementBuilder.cs`** — Decomposes a clustered, stacked or
  percent-stacked bar/column `ChartModel` into rects (bars), gridlines, axis
  lines, legend swatches and text labels inside the `graphicFrame` rect.
- **`PieChartElementBuilder.cs`** — Decomposes a pie or doughnut `ChartModel`
  into arc-approximated wedge polygons (two even-odd contours when a doughnut
  hole is present) plus title/legend text inside the `graphicFrame` rect.
- **`PptxToTypstConverter.Charts.cs`** — Partial class on the converter:
  resolves chart parts, parses them, and dispatches bar/column charts to
  `BarChartElementBuilder` and pie/doughnut charts to `PieChartElementBuilder`.

## Tests

Chart conversion tests live in `DocxEditor.Tests/Unit/Charts/`
(`ChartPartParserTests`, `BarChartElementBuilderTests`,
`PieChartElementBuilderTests`, `ChartAxisScaleTests`,
`ChartConversionIntegrationTests`), plus
`DocxEditor.Tests/Unit/PptxToTypstConverterChartConnectorTests.cs`.
