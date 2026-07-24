# Chart Conversion — D1 Report (Bar/Column charts)

Branch: `feat/chart-conversion` · Worktree: `DocxEditor-feat-charts`

## 1. Inventory — chart usage across `examples/REF/PPTX/*.pptx`

| Deck | Chart parts | Location | Type | Grouping | Series × cats | Legend | Axes | Data labels |
|---|---|---|---|---|---|---|---|---|
| `sales_acceleration_deck.pptx` | `ppt/charts/chart1.xml` | slide 9, frame 60,110 → 600×330 pt | `barChart` `barDir="bar"` (horizontal) | `clustered` (gap 60, overlap −20) | 2 × 5 | `r` | catAx (l) + valAx (b, gridlines `#D7D7D7`) | `showVal`, 9 pt Verdana `#1B1B1B` |
| `AetherLink-Glass-Shareholder-Overview.pptx` | `ppt/charts/chart1.xml` | slide 9 | `barChart` `barDir="col"` (vertical) | `clustered` (gap 70, overlap −10) | 2 × 4 | present, **no `legendPos`** (default = right) | catAx (b, line `#9FB3D9`) + valAx (l, gridlines `#2A3560`) | `showVal`, 9 pt QuattrocentoSans `#EAF2FF` |
| `northwind-demo.pptx` | — | — | — | — | — | — | — | — |
| `northwind-investor-40.pptx` | — | — | — | — | — | — | — | — |
| `northwind-launch-review.pptx` | — | — | — | — | — | — | — | — |

**No `lineChart` or `pieChart` exists anywhere in the REF corpus.** Both real charts are
clustered bar charts with explicit per-series `solidFill` (`srgbClr`), cached string/number
caches (`c:strCache`/`c:numCache`), no explicit axis min/max (auto), and `General` number
format. All computed values are cached in the XML — the parser reads caches only, never
resolves `c:f` references (their target is an embedded XLSX we don't need).

### Reference-render observations (drives emission math)

From the official PDFs (slide 9 of each deck):

- **Horizontal bars plot categories bottom-to-top** (first `c:pt` at the bottom), series
  within a slot bottom-to-top in series order, and the **legend is reversed** to match.
- **Auto value-axis max extends past the data max when data labels are shown**:
  data max 100 → axis 120 (unit 20); data max 48 → axis 60 (unit 10).
  Approximation used: `unit = NiceCeil(max/5)`, `axisMax = unit · ⌈(max + 0.5·unit)/unit⌉`
  when labels are shown (both reference cases reproduced exactly), else `unit · ⌈max/unit⌉`.
- Gridlines at every major tick; axis lines drawn even without explicit `a:ln`
  (fallback colour = axis label colour).

## 2. Architecture

```
PptxToTypstConverter.ConvertGraphicFrame            (existing; +~10-line chart hook)
  └─ PptxToTypstConverter.Charts.cs                 (NEW partial — part resolution, fallback policy)
       ├─ ChartPartParser.Parse(xml, schemeResolver) → ChartModel
       │     (System.Xml.Linq; caches only; colours resolved at parse time)
       └─ BarChartElementBuilder.Build(model, frameRect) → List<TypstElement>
             ├─ ChartAxisScale (nice-round auto axis math, pure static)
             └─ decomposes to EXISTING primitives: rect (bars, gridlines, axis lines,
                legend swatches) + Text (category/value/tick/legend labels)
```

`SourceGeneration.cs` and `TypstModels.cs` need **zero changes** — every chart visual is a
`TypstElement` of type `Shape` (rect) or `Text`, emitted through the existing `#place`
pipeline.

### Files

| File | Role |
|---|---|
| `PptxEditor.Core/Converters/Charts/ChartModel.cs` | Immutable chart model (kind, direction, grouping, series, axes, legend, labels) |
| `PptxEditor.Core/Converters/Charts/ChartPartParser.cs` | `c:chartSpace` XML → `ChartModel` (caches only) |
| `PptxEditor.Core/Converters/Charts/ChartAxisScale.cs` | Auto axis min/max/unit with nice-rounding |
| `PptxEditor.Core/Converters/Charts/BarChartElementBuilder.cs` | Layout + decomposition to Typst primitives |
| `PptxEditor.Core/Converters/Charts/PptxToTypstConverter.Charts.cs` | Converter partial: part resolution + routing + fallback |
| `PptxToTypstConverter.cs` | ~10-line hook routing chart graphicFrames to the pipeline |
| `DocxEditor.Tests/Unit/Charts/*.cs` | Parser / scale / builder / integration tests |

### Fallback policy (unchanged for unsupported content)

Line/pie/stacked charts and unresolvable/unparseable chart parts still render the existing
visible placeholder and warn; the warning now names the detected chart type
(e.g. `Chart 'X' (line chart) is not supported yet …`).

## 3. Extension points for round 2 (line / pie)

- `ChartModel.Kind` already carries `Line`/`Pie`/`Other`; the parser already recognises
  `lineChart`/`pieChart` roots (they fall back with an accurate warning today).
- **Line charts**: reuse `ChartAxisScale` + the plot-area/category-slot math; emit polyline
  segments via the existing `polygon` shape (stroke only) plus `ellipse` markers — no new
  Typst element types needed. Add a `LineChartElementBuilder` next to
  `BarChartElementBuilder` and route on `Kind == Line` in `PptxToTypstConverter.Charts.cs`.
- **Pie charts**: need per-slice percentage math (values are already cached; `showPercent`
  flag parsed via `ChartDataLabels`). Typst emission options: approximate slices with
  `polygon` point fans (many-point arc tessellation) — parser and legend plumbing from D1
  carry over unchanged.
- **Stacked bars**: `BarGrouping` is parsed (`Stacked`/`PercentStacked`); the builder
  currently declines them (fallback + warning). Stacked support = cumulative offsets in the
  existing bar loop + optional percent axis normalisation.
- **Chart titles**: `c:title`/`autoTitleDeleted` not yet modelled; add to `ChartModel` and
  emit one centred `Text` above the plot area (auto-title rule: shown only for
  single-series charts when `autoTitleDeleted=0`).
- **Colour transforms** (`lumMod`/`lumOff`/`tint`/`shade`) on chart colours are parsed as
  the base colour only; apply transforms at parse time if a deck needs them.

## 4. Self-check RMSE (sales_acceleration_deck slide 9)

(filled in after implementation — before/after numbers)
