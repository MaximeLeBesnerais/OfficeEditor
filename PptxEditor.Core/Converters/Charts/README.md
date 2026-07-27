# Charts Converter

Parses OOXML `c:chartSpace` parts inside `graphicFrame` slide elements and
decomposes them into existing Typst primitives (rects, text, lines) so chart
content is rendered natively instead of falling back to a placeholder frame.

## Entry point

`PptxToTypstConverter.Charts.cs` — the `ConvertChartGraphicFrame` partial
method on `PptxToTypstConverter`. It resolves the `c:chart` relationship to
its chart part XML, hands it to `ChartPartParser`, and dispatches supported
chart types to `BarChartElementBuilder`. Unsupported chart kinds (line, pie,
doughnut, stacked) get a visible placeholder + warning.

## Chart type support

| Chart kind | Grouping | Supported |
|---|---|---|
| Bar/Column | Clustered | Yes |
| Bar/Column | Stacked / PercentStacked | Placeholder |
| Line | — | Placeholder |
| Pie | — | Yes (wedge polygons) |
| Doughnut | — | Rendered as pie (hole ignored) |

## Files

- **`ChartModel.cs`** — Immutable model of a parsed chart (kind, direction,
  grouping, series, categories, axes, legend, data labels). All values come
  from cached `strCache`/`numCache` points; cell references (`c:f`) are
  never evaluated.
- **`ChartPartParser.cs`** — `XDocument`-based parser that walks
  `c:chartSpace` XML into a `ChartModel`. Uses `XDocument` (not the OpenXML
  SDK) so synthetic chart XML can be unit-tested without a package.
- **`ChartAxisScale.cs`** — Auto value-axis scaling with PowerPoint-style
  nice rounding. Pure static math shared by bar/column builders.
- **`BarChartElementBuilder.cs`** — Decomposes a clustered bar/column
  `ChartModel` into rects (bars), gridlines, axis lines, legend swatches and
  text labels inside the `graphicFrame` rect.
- **`PptxToTypstConverter.Charts.cs`** — Partial class on the converter:
  resolves chart parts, parses them, dispatches clustered bar/column charts
  to the element builder.

## Tests

Chart conversion tests live in `DocxEditor.Tests/` under the PPTX converter
test suite (`*Chart*Tests.cs` patterns).
