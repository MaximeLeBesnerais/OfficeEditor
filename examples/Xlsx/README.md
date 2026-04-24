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

> Excel calculates: TOTAL row = 20000, 26000, 22000, 30000, 98000

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

---

## JSON Instructions

**Input:** `instructions/sample.json`

```json
{
  "version": "1.0",
  "worksheets": [
    {
      "name": "Revenue",
      "headers": ["Month", "Product A", "Product B", "Product C", "Total"],
      "rows": [
        ["January", "{{janA}}", "{{janB}}", "{{janC}}", "=SUM(B2:D2)"],
        ["February", "{{febA}}", "{{febB}}", "{{febC}}", "=SUM(B3:D3)"],
        ["March", "{{marA}}", "{{marB}}", "{{marC}}", "=SUM(B4:D4)"]
      ]
    },
    {
      "name": "Summary",
      "cells": [
        { "address": "A1", "value": "Total Revenue Q1" },
        { "address": "B1", "formula": "=SUM(Revenue!E2:E4)" }
      ]
    }
  ],
  "variables": {
    "janA": "10000", "janB": "8000", "janC": "6000",
    "febA": "12000", "febB": "9500", "febC": "7000",
    "marA": "11000", "marB": "9000", "marC": "7500"
  }
}
```

**Output:** Workbook generated from the instruction set.

---

## Running Examples

```bash
cd examples
dotnet run
```
