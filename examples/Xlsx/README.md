# XLSX Examples

Every example shows **what you start with → the code you run → what you get**.

---

## 1. Basic Workbook

**Input:** None (created from scratch)

**Code:**
```csharp
using var builder = WorkbookBuilder.Create("01-basic.xlsx");

var sales = builder.AddWorksheet("Sales");

sales.AddHeaderRow(new List<string> { "Product", "Q1", "Q2", "Q3", "Q4" });
sales.AddDataRow(new List<string> { "Widget", "100", "150", "200", "250" }, 2);
sales.AddDataRow(new List<string> { "Gadget", "80", "120", "160", "200" }, 3);
sales.AddDataRow(new List<string> { "Tool", "60", "90", "120", "150" }, 4);

builder.Save();
```

**Output:** `output/xlsx/01-basic.xlsx` — Worksheet "Sales"

| | A | B | C | D | E |
|---|---|---|---|---|---|
| 1 | Product | Q1 | Q2 | Q3 | Q4 |
| 2 | Widget | 100 | 150 | 200 | 250 |
| 3 | Gadget | 80 | 120 | 160 | 200 |
| 4 | Tool | 60 | 90 | 120 | 150 |

---

## 2. Formulas & Calculations

**Input:** None (created from scratch)

**Code:**
```csharp
using var builder = WorkbookBuilder.Create("02-formulas.xlsx");

var sheet = builder.AddWorksheet("Budget");

sheet.AddHeaderRow(new List<string> { "Category", "Q1", "Q2", "Q3", "Q4", "Total" });
sheet.AddDataRow(new List<string> { "Revenue", "10000", "12000", "11000", "15000" }, 2);
sheet.AddDataRow(new List<string> { "Expenses", "6000", "7000", "6500", "8000" }, 3);
sheet.AddDataRow(new List<string> { "Profit", "4000", "5000", "4500", "7000" }, 4);
sheet.AddCell("F2", "=SUM(B2:E2)", true);
sheet.AddCell("F3", "=SUM(B3:E3)", true);
sheet.AddCell("F4", "=SUM(B4:E4)", true);

sheet.AddFormulaRow(new List<string> {
    "TOTAL", "=SUM(B2:B4)", "=SUM(C2:C4)", "=SUM(D2:D4)", "=SUM(E2:E4)", "=SUM(F2:F4)"
}, 5);

sheet.AddFormulaRow(new List<string> {
    "AVERAGE", "=AVERAGE(B2:B4)", "=AVERAGE(C2:C4)", "=AVERAGE(D2:D4)", "=AVERAGE(E2:E4)", "=AVERAGE(F2:F4)"
}, 6);

builder.Save();
```

**Output:** `output/xlsx/02-formulas.xlsx` — Worksheet "Budget"

| | A | B | C | D | E | F |
|---|---|---|---|---|---|---|
| 1 | Category | Q1 | Q2 | Q3 | Q4 | Total |
| 2 | Revenue | 10000 | 12000 | 11000 | 15000 | |
| 3 | Expenses | 6000 | 7000 | 6500 | 8000 | |
| 4 | Profit | 4000 | 5000 | 4500 | 7000 | |
| 5 | **TOTAL** | `=SUM(B2:B4)` | `=SUM(C2:C4)` | `=SUM(D2:D4)` | `=SUM(E2:E4)` | `=SUM(F2:F4)` |
| 6 | **AVERAGE** | `=AVERAGE(B2:B4)` | ... | ... | ... | ... |

> Excel calculates: TOTAL row = 20000, 24000, 22000, 30000, 96000

---

## 3. Multiple Worksheets (Cross-References)

**Input:** None (created from scratch)

**Code:**
```csharp
using var builder = WorkbookBuilder.Create("03-multi-sheet.xlsx");

// Sheet 1: Sales Data
var sales = builder.AddWorksheet("Sales Data");
sales.AddHeaderRow(new List<string> { "Month", "Revenue", "Expenses", "Profit" });
sales.AddDataRow(new List<string> { "Jan", "10000", "6000", "4000" }, 2);
sales.AddDataRow(new List<string> { "Feb", "12000", "7000", "5000" }, 3);
sales.AddDataRow(new List<string> { "Mar", "11000", "6500", "4500" }, 4);

// Sheet 2: Summary (references Sheet 1)
var summary = builder.AddWorksheet("Summary");
summary.AddHeaderRow(new List<string> { "Metric", "Value" });
summary.AddDataRow(new List<string> { "Total Revenue", "" }, 2);
summary.AddCell("B2", "=SUM('Sales Data'!B2:B4)", true);
summary.AddDataRow(new List<string> { "Total Expenses", "" }, 3);
summary.AddCell("B3", "=SUM('Sales Data'!C2:C4)", true);
summary.AddDataRow(new List<string> { "Total Profit", "" }, 4);
summary.AddCell("B4", "=SUM('Sales Data'!D2:D4)", true);
summary.AddDataRow(new List<string> { "Average Profit", "" }, 5);
summary.AddCell("B5", "=AVERAGE('Sales Data'!D2:D4)", true);

builder.Save();
```

**Output:** `output/xlsx/03-multi-sheet.xlsx`

**Worksheet "Sales Data":**

| | A | B | C | D |
|---|---|---|---|---|
| 1 | Month | Revenue | Expenses | Profit |
| 2 | Jan | 10000 | 6000 | 4000 |
| 3 | Feb | 12000 | 7000 | 5000 |
| 4 | Mar | 11000 | 6500 | 4500 |

**Worksheet "Summary":**

| | A | B |
|---|---|---|
| 1 | Metric | Value |
| 2 | Total Revenue | `=SUM('Sales Data'!B2:B4)` → 33000 |
| 3 | Total Expenses | `=SUM('Sales Data'!C2:C4)` → 19500 |
| 4 | Total Profit | `=SUM('Sales Data'!D2:D4)` → 13500 |
| 5 | Average Profit | `=AVERAGE('Sales Data'!D2:D4)` → 4500 |

---

## 4. Variable Detection

**Input:** Template with variables

```csharp
using var builder = WorkbookBuilder.Create("04-template.xlsx");
var sheet = builder.AddWorksheet("Report");

sheet.AddHeaderRow(new List<string> { "Field", "Value" });
sheet.AddDataRow(new List<string> { "Company", "{{companyName}}" }, 2);
sheet.AddDataRow(new List<string> { "Report Date", "{{reportDate}}" }, 3);
sheet.AddDataRow(new List<string> { "Prepared By", "{{preparedBy}}" }, 4);
sheet.AddDataRow(new List<string> { "Total Sales", "{{totalSales}}" }, 5);

builder.Save();
```

**Operation:** Detect variables
```csharp
using var detector = WorkbookBuilder.Open("04-template.xlsx");
var variables = detector.DetectVariables();
// Returns: companyName, reportDate, preparedBy, totalSales
```

**Output:** `output/xlsx/04-template.xlsx` (unchanged)

Console:
```
Found 4 variables:
  - companyName
  - reportDate
  - preparedBy
  - totalSales
```

---

## 5. Mail Merge (Variable Replacement)

**Input:** Template with variables

```csharp
using var builder = WorkbookBuilder.Create("05-template.xlsx");
var sheet = builder.AddWorksheet("Invoice");

sheet.AddHeaderRow(new List<string> { "Description", "Quantity", "Price", "Total" });
sheet.AddDataRow(new List<string> { "{{item1}}", "{{qty1}}", "{{price1}}", "=B2*C2" }, 2);
sheet.AddDataRow(new List<string> { "{{item2}}", "{{qty2}}", "{{price2}}", "=B3*C3" }, 3);
sheet.AddDataRow(new List<string> { "{{item3}}", "{{qty3}}", "{{price3}}", "=B4*C4" }, 4);
sheet.AddDataRow(new List<string> { "TOTAL", "", "", "=SUM(D2:D4)" }, 5);

builder.Save();
```

**Operation:** Replace variables
```csharp
using var builder = WorkbookBuilder.Open("05-template.xlsx");
builder.MergeVariables(new Dictionary<string, string>
{
    ["item1"] = "Consulting Services", ["qty1"] = "10", ["price1"] = "150",
    ["item2"] = "Software License",   ["qty2"] = "1",  ["price2"] = "500",
    ["item3"] = "Training",           ["qty3"] = "5",  ["price3"] = "100"
});
builder.Save("05-processed.xlsx");
```

**Output:** `output/xlsx/05-processed.xlsx` — Worksheet "Invoice"

| | A | B | C | D |
|---|---|---|---|---|
| 1 | Description | Quantity | Price | Total |
| 2 | Consulting Services | 10 | 150 | `=B2*C2` → 1500 |
| 3 | Software License | 1 | 500 | `=B3*C3` → 500 |
| 4 | Training | 5 | 100 | `=B4*C4` → 500 |
| 5 | **TOTAL** | | | `=SUM(D2:D4)` → 2500 |

> **Note:** The original template `05-template.xlsx` is preserved.
>
> **Inline strings:** templates may store text in inline-string cells (`<is><t>`)
> as well as shared strings. `DetectVariables`/`MergeVariables` read and rewrite both
> cell kinds in place, so a merged value keeps the cell's type instead of leaving a
> stale inline string beside a new `<v>`.

---

## 6. JSON → XLSX (rich-report vocabulary + generator)

The JSON instruction/generation engine (`XlsxEditor.Core/Instructions/`) parses, loudly validates, plans, and executes a v1 rich vocabulary: **typed cells, named styles, column widths, row heights, merges, freeze panes, autofilter, tables, formulas with styles/number formats, workbook metadata, and variable resolution.** It is exercised end-to-end by the test suite against [`instructions/rich-report.json`](instructions/rich-report.json) — a three-sheet report covering every feature.

**Input:** `instructions/rich-report.json` (abridged below — see the file for the full set: styles, columns, rowHeights, merges, freezePanes, autoFilter, tables, typed cells, legacy style ids, variables).

```json
{
  "version": "1.0",
  "description": "Comprehensive three-sheet rich report ...",
  "metadata": {
    "title": "Northwind Labs Quarterly Report",
    "author": "OfficeEditor Generator",
    "subject": "Q3 2026 Performance",
    "category": "Finance",
    "keywords": "northwind, quarterly, revenue",
    "comments": "Generated by XlsxGenerator from a declarative instruction set."
  },
  "styles": [
    { "name": "header", "font": { "bold": true, "color": "FFFFFF", "size": 11 },
      "fill": { "color": "4472C4", "pattern": "solid" },
      "alignment": { "horizontal": "center", "vertical": "center" } },
    { "name": "money", "numberFormat": "$#,##0.00" },
    { "name": "pct",   "numberFormat": "0.0%" },
    { "name": "title", "font": { "bold": true, "size": 14 }, "alignment": { "horizontal": "left" } }
  ],
  "variables": { "reportPeriod": "Q3 2026", "preparedBy": "Analytics Team" },
  "worksheets": [
    {
      "name": "Summary",
      "startRow": 2,
      "headerStyle": "header",
      "columns": [
        { "name": "Region", "width": 16, "type": "string" },
        { "name": "Revenue", "width": 14, "style": "money", "type": "number" },
        { "name": "Growth", "width": 10, "style": "pct", "type": "number" },
        { "name": "Active", "width": 10, "type": "boolean" },
        { "name": "As Of", "width": 12, "type": "date" }
      ],
      "headers": ["Region", "Revenue", "Growth", "Active", "As Of"],
      "rows": [
        ["North America", "1250000", "0.082", "true", "2026-09-30"],
        ["Europe", "980000", "0.051", "true", "2026-09-30"]
      ],
      "rowHeights": [ { "row": 1, "height": 20 }, { "row": 2, "height": 24 } ],
      "merges": ["A1:E1"],
      "cells": [
        { "address": "A1", "value": "{{reportPeriod}} — {{preparedBy}}", "style": "title" },
        { "address": "F1", "value": "Total Revenue", "type": "string", "style": "title" },
        { "address": "F2", "formula": "=SUM(B3:B6)", "style": "money", "numberFormat": "$#,##0.00" }
      ],
      "freezePanes": { "cell": "A3" },
      "tables": [ { "name": "SummaryTable", "range": "A2:E6" } ]
    }
  ]
}
```

**Generate from the library:**
```csharp
using XlsxEditor.Core.Instructions;

// Result carries bytes + diagnostics; errors block, warnings do not.
var result = XlsxGenerator.GenerateFromFile("instructions/rich-report.json");
File.WriteAllBytes("rich-report.xlsx", result.Bytes!);

// Atomic file output: a failure never leaves a partial/corrupt workbook behind.
XlsxGenerator.GenerateToFile(json, "rich-report.xlsx");

// Throwing convenience: bytes or the first error as an XlsxException.
var bytes = XlsxGenerator.GenerateBytes(json);
```

`XlsxGenerator` accepts JSON text, a stream, a file path, or an already-parsed `XlsxInstructionSet` (`XlsxInstructionParser.ParseFromFile`). `XlsxGenerateOptions.ValidatePackage` runs the OpenXML SDK validator over the finished package. Generation is **plan-first**: every set is validated and every variable/type/address/range/style is resolved before anything is mutated, so a rejected set can never leave a partially-written workbook. The minimal v1 budget sample [`instructions/sample.json`](instructions/sample.json) also still parses and executes end-to-end.

### Vocabulary (v1)

| Level | Keys |
|---|---|
| Root | `version` (`"1.0"`), `description`, `metadata`, `styles`, `worksheets`, `variables` |
| `metadata` | `title`, `subject`, `author`, `category`, `keywords`, `comments` → workbook core properties |
| `styles[]` | `name`, `font` (`bold`/`italic`/`color`/`size`), `fill` (`color` + `pattern`), `border` (`left`/`right`/`top`/`bottom`, each a `style` + optional `color`), `alignment` (`horizontal`/`vertical`/`wrapText`), `numberFormat` |
| `worksheets[]` | `name`, `headers`, `headerStyle`, `startRow`, `columns[]`, `rows[]`, `rowHeights[]`, `cells[]`, `merges[]`, `freezePanes`, `autoFilter`, `tables[]` |
| `columns[]` | `name`, `width` (Excel units, 1–255), `style`, `type` |
| `cells[]` | `address` + `value` XOR `formula` (formula starts with `=`), `type`, `numberFormat`, `style` |
| `rowHeights[]` | `row` + `height` (points) |
| `freezePanes` | `cell` (e.g. `"A3"`) or `row`/`column` |
| `tables[]` | `name` (unique workbook-wide) + `range` |

Cell `type` is `auto`, `string`, `number`, `boolean`, `date`, or `datetime`; `auto` infers from the resolved value (numbers parse, `true`/`false` become booleans, ISO-like values become date/datetime, everything else is text). Explicit `date` requires `yyyy-MM-dd` and `datetime` requires `yyyy-MM-ddTHH:mm:ss` (optionally `.fff`); typed values that fail to parse fail the whole set with a descriptive `XlsxException`. Date/datetime cells are stored as Excel serial numbers under a date number format (`yyyy-mm-dd`, or `yyyy-mm-dd h:mm:ss` for datetimes) so they render as dates in Excel. A `style` reference is a named style or a legacy numeric style-id string (e.g. `"0"`). Style colors accept a 6-digit RGB, 8-digit ARGB hex string, or a common color name; fills accept any real Excel pattern type (`solid`, `gray125`, …); borders take up to four edges (`left`/`right`/`top`/`bottom`), each a border style (`thin`, `medium`, `thick`, `double`, `dashed`, `dotted`, …) plus an optional color. Formula cells are written without a cached value and the workbook sets `fullCalcOnLoad`, so Excel/LibreOffice recalculates the results when the file opens. Variables `{{name}}` resolve at planning time in values and formulas; unresolved placeholders and variable cycles are errors.

Validation is loud: unknown keys, invalid sheet names, out-of-bounds addresses (columns A–XFD, rows 1–1,048,576), cells with both/neither `value` and `formula`, over-limit cell text (32,767 chars) and formulas (8,192 chars), and unknown types/alignments are all rejected with path-qualified errors. Duplicate writes to one cell warn (last write wins).

### CLI

The wired command is the unified `generate` surface (extension-selected, like DOCX/PPTX):

```bash
officeeditor generate examples/Xlsx/instructions/rich-report.json --output rich-report.xlsx
```

The same instruction set also **renders straight to PDF or PNG** through the Typst pipeline — see [Rendering & Conversion](#7-rendering--conversion-xlsx--pdfpngsvg) below:

```bash
officeeditor generate examples/Xlsx/instructions/rich-report.json --output rich-report.pdf
officeeditor generate examples/Xlsx/instructions/rich-report.json --output rich-report.png   # directory of page-NNN.png
```

`officeeditor create <out.xlsx> --instructions <file.json>` is still **not implemented** — `create` currently makes a blank sheet, and XLSX `edit` is unimplemented (the unified `edit` command is a stub). Drive the vocabulary through `XlsxGenerator` / `XlsxInstructionExecutor` from the library when you need it from code.

### Explicit limitations

- **Images** — there is no `images[]` vocabulary; pictures are fluent-API only and not in JSON.
- **Row replication** — there is no `repeat`/`foreach` loop; the roadmap's hardest item remains.
- **Number formats** — applied via named styles or per-cell `numberFormat`; there is no separate `numFmt` table authoring.
- **Charts, defined names, data validation, panes beyond freeze** — out of the v1 vocabulary.
- **`create --instructions` / `edit`** — not wired; use `officeeditor generate <input.json> --output <out.xlsx>` or the library.
- **Rendering evaluates no formulas** — the renderer (see [Rendering & Conversion](#7-rendering--conversion-xlsx--pdfpngsvg) below) draws the workbook's cached `<v>` values only. OfficeEditor-written formula cells carry no cached value (the workbook sets `fullCalcOnLoad`), so they render as empty cells until the file is opened in Excel/LibreOffice; third-party workbooks that save cached results (Excel/LibreOffice do) render those.

---

## 7. Rendering & Conversion (XLSX → PDF/PNG/SVG)

The renderer (`XlsxEditor.Core/Rendering/`) converts spreadsheets to **PDF / PNG / SVG** using the repository's own Typst pipeline: `XlsxReader` → `XlsxToTypstConverter` (native Typst `table` emitter) → `TypstCompilerService` (TypstBridge primary, typst CLI safety net). **No external office suite is used in product code** — LibreOffice appears only as a test-only oracle in `visual-diff --suite xlsx`. It covers typed cells, number formats (`$#,##0.00`, `0.0%`, dates, `[Red]` negatives), fills/borders/fonts/alignment, merges, explicit column widths & row heights, freeze-pane header repetition, and auto-pagination. See the [XLSX roadmap](../docs/roadmap-xlsx.md) "Out of scope" note for the current scope and limits.

### `tools/convert-xlsx` — all five formats

Input is an `.xlsx` file **or** a `.json` instruction set; `--format` defaults to `pdf`:

```bash
# PDF — writes a single file
dotnet run --project tools/convert-xlsx -- rich-report.json rich-report.pdf --format pdf

# PNG / SVG — each writes a directory of page-NNN.ext pages
dotnet run --project tools/convert-xlsx -- rich-report.xlsx out/pages --format png --ppi 150
dotnet run --project tools/convert-xlsx -- rich-report.xlsx out/pages --format svg

# Typst source
dotnet run --project tools/convert-xlsx -- rich-report.xlsx rich-report.typ --format typ

# .xlsx → JSON instruction set (best-effort round-trip back into the v1 vocabulary)
dotnet run --project tools/convert-xlsx -- rich-report.xlsx rich-report.json --format json
```

`--font-path <path>` adds an extra font directory; `--ppi <n>` sets PNG resolution (default 150).

### Unified CLI and API

- `officeeditor generate <input.json> --output out.pdf` — PDF (single file); `--output out.png` — a directory of `page-NNN.png` pages.
- `POST /api/convert` with an `.xlsx` source and `targetFormat=pdf|png|svg` renders through the same pipeline (`ConvertXlsxToRenderAsync`); PNG/SVG return the first page.

### Visual regression — `visual-diff --suite xlsx`

The suite diffs **our** Typst-rendered PDF against a **LibreOffice-generated reference PDF** — LibreOffice is a test-only oracle, never product code:

```bash
dotnet run --project tools/visual-diff -- --suite xlsx
dotnet run --project tools/visual-diff -- --suite xlsx --generate   # build missing generated PDFs via tools/convert-xlsx
```

### Fixtures

- [`instructions/rich-report.json`](instructions/rich-report.json) — the three-sheet rich report above (also a rendering fixture).
- [`instructions/complex-dashboard.json`](instructions/complex-dashboard.json) — financial dashboard exercising borders, fill patterns, cross-sheet formulas, every cell type, and layout features; the renderer's stress fixture.
- Reference PDFs (LibreOffice-rendered, committed): `examples/REF/XLSX/rich-report.pdf` and `examples/REF/XLSX/complex-dashboard.pdf`.

---

## 8. Layout & Merges

Explicit column widths and row heights, and display-only cell merges, via the fluent API.

**Code:**
```csharp
using var builder = WorkbookBuilder.Create("07-layout.xlsx");
var sheet = builder.AddWorksheet("Report");

sheet.AddHeaderRow(new List<string> { "Metric", "Q1", "Q2" });
sheet.AddDataRow(new List<string> { "Revenue", "100", "150" }, 2);

sheet.SetColumnWidth("A", 24);      // IWorksheetBuilder SetColumnWidth(string column, double width)
sheet.SetColumnWidth("B", 12.5);
sheet.SetRowHeight(1, 30);          // IWorksheetBuilder SetRowHeight(int rowIndex, double height)

sheet.MergeCells("A1:A1");          // rejected: single cell
sheet.MergeCells("A1:C1");          // header spans the table
sheet.UnmergeCells("A1:C1");        // restores it (no-op if not merged)

sheet.GetMergeRanges();             // List<string> in A1 notation, document order
sheet.GetColumnWidth("B");          // double? -> 12.5 (null when undefined)
sheet.GetRowHeight(1);              // double? -> 30 (null when undefined)
builder.Save();
```

Widths are in Excel column-width units (characters of the default font, max 255); heights are in points (max 409.5). Ranges are normalized and bounded to Excel's real sheet limits (columns A–XFD, rows 1–1,048,576) before mutation; malformed, single-cell, reversed, duplicate and overlapping merges throw an `XlsxException` and leave the worksheet untouched.

---

## 9. Read Back

Reopen a workbook and inspect cells, rows, ranges, and dimensions.

**Code:**
```csharp
using var builder = WorkbookBuilder.Open("03-multi-sheet.xlsx");
var summary = builder.GetWorksheet("Summary");

summary.GetCellValue("B2");              // string? -> cached value, shared-string resolved
summary.GetCellFormula("B2");            // string? -> "=SUM('Sales Data'!B2:B4)"
summary.CellExists("B2");                // bool
summary.GetCellInfo("B2");               // CellInfo? (Reference/Value/Formula/DataType)
summary.GetRange("A1", "B5");            // List<CellInfo>
summary.GetRows();                       // List<RowInfo>
summary.GetRow(2);                       // RowInfo?
summary.GetDimensions();                 // (firstRow, lastRow, firstCol, lastCol)
builder.Save();
```

Missing cells read as `null` (not exceptions); `CellExists` distinguishes an absent cell from an empty one.

---

## Running Examples

```bash
cd examples
dotnet run
```

The runner currently executes sections 1–5. Sections 6–9 document tested library APIs but are not yet called from `examples/Program.cs`.
