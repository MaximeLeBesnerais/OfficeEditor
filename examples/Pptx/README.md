# PPTX Examples

Every example shows **what you start with → the code you run → what you get**.

---

## 1. Basic Presentation

**Input:** None (created from scratch)

**Code:**
```csharp
using var builder = PresentationBuilder.Create("01-basic.pptx");

// Slide 1: Title
builder.AddSlide();
builder.CurrentSlide
    .AddTitle("OfficeEditor Presentation")
    .AddSubtitle("Created with Fluent C# API");

// Slide 2: Bullet list
builder.AddSlide();
builder.CurrentSlide
    .AddTitle("Key Features")
    .AddBulletList(new[] {
        "Easy slide creation",
        "Multiple content types",
        "Variable detection",
        "Typst export"
    });

// Slide 3: Text blocks
builder.AddSlide();
builder.CurrentSlide
    .AddTitle("Getting Started")
    .AddText("1. Install the NuGet package")
    .AddText("2. Create a PresentationBuilder")
    .AddText("3. Add slides and content")
    .AddText("4. Save or export");

builder.Save();
```

**Output:** `output/pptx/01-basic.pptx` (3 slides)

| Slide | Content |
|-------|---------|
| 1 | **OfficeEditor Presentation** / Created with Fluent C# API |
| 2 | **Key Features** / • Easy slide creation • Multiple content types ... |
| 3 | **Getting Started** / 1. Install... 2. Create... |

---

## 2. Advanced Content (Tables & Lists)

**Input:** None (created from scratch)

**Code:**
```csharp
using var builder = PresentationBuilder.Create("02-advanced.pptx");

builder.AddSlide();
builder.CurrentSlide
    .AddTitle("Q4 Sales Report")
    .AddSubtitle("2024 Performance Review");

builder.AddSlide();
builder.CurrentSlide
    .AddTitle("Sales by Region")
    .AddTable(new List<List<string>>
    {
        new() { "Region", "Q1", "Q2", "Q3", "Q4", "Total" },
        new() { "North", "$100K", "$120K", "$110K", "$140K", "$470K" },
        new() { "South", "$80K", "$95K", "$105K", "$130K", "$410K" },
        new() { "East", "$90K", "$100K", "$120K", "$150K", "$460K" },
        new() { "West", "$110K", "$125K", "$135K", "$160K", "$530K" }
    });

builder.AddSlide();
builder.CurrentSlide
    .AddTitle("Top Priorities")
    .AddNumberedList(new[] {
        "Expand to new markets",
        "Improve customer retention",
        "Launch new product line",
        "Optimize operations"
    });

builder.Save();
```

**Output:** `output/pptx/02-advanced.pptx` (3 slides)

| Slide | Content |
|-------|---------|
| 1 | **Q4 Sales Report** / 2024 Performance Review |
| 2 | **Sales by Region** / 5×6 table with revenue data |
| 3 | **Top Priorities** / 1. Expand... 2. Improve... |

---

## 3. Variable Detection & Mail Merge

**Input:** Template with variables

```csharp
using var builder = PresentationBuilder.Create("03-template.pptx");

builder.AddSlide();
builder.CurrentSlide
    .AddTitle("Welcome {{companyName}}")
    .AddSubtitle("Presented by {{presenterName}}");

builder.AddSlide();
builder.CurrentSlide
    .AddTitle("Project: {{projectName}}")
    .AddText("Budget: ${{budget}}")
    .AddText("Timeline: {{timeline}}")
    .AddText("Status: {{status}}");

builder.Save();
```

**Operation 1:** Detect variables
```csharp
using var detector = PresentationBuilder.Open("03-template.pptx");
var variables = detector.DetectVariables();
// Returns: companyName, presenterName, projectName, budget, timeline, status
```

**Operation 2:** Merge
```csharp
using var builder = PresentationBuilder.Open("03-template.pptx");
builder.MergeVariables(new Dictionary<string, string>
{
    ["companyName"] = "TechCorp Inc.",
    ["presenterName"] = "Jane Smith",
    ["projectName"] = "Cloud Migration",
    ["budget"] = "500,000",
    ["timeline"] = "6 months",
    ["status"] = "On Track"
});
builder.Save("03-merged.pptx");
```

**Output:** `output/pptx/03-merged.pptx`

| Slide | Content |
|-------|---------|
| 1 | **Welcome TechCorp Inc.** / Presented by Jane Smith |
| 2 | **Project: Cloud Migration** / Budget: $500,000 / Timeline: 6 months / Status: On Track |

---

## 4. PPTX → PDF (Typst Integration)

**Input:** PowerPoint file

```csharp
// First create the PPTX
using (var builder = PresentationBuilder.Create("04-export.pptx"))
{
    builder.AddSlide();
    builder.CurrentSlide
        .AddTitle("PDF Export Demo")
        .AddSubtitle("Using Typst Integration");

    builder.AddSlide();
    builder.CurrentSlide
        .AddTitle("Features")
        .AddBulletList(new[] {
            "High-fidelity PDF output",
            "Preserves formatting",
            "Embeds fonts and images",
            "Fast compilation"
        });

    builder.Save();
}
```

**Operation:** Export to PDF
```csharp
using var builder = PresentationBuilder.Open("04-export.pptx");
byte[] pdfBytes = builder.ExportToPdf();
await File.WriteAllBytesAsync("04-export.pdf", pdfBytes);
```

**Output:** `output/pptx/04-export.pdf` (2-page PDF)

> Converts PPTX slides → Typst source → compiles to PDF through `TypstCompilerService` with TypstBridge as the primary backend.

> **Font fidelity note:** for closest results, install the same fonts used by the
> source presentation, such as Microsoft Aptos for modern Office decks. Installed
> fonts improve matching, but exact PowerPoint parity is not guaranteed because
> Typst, PowerPoint, and platform font engines can differ in font metrics, line
> wrapping, hinting, and layout behavior.

---

## 5. PPTX → Typst Source

**Input:** Same PowerPoint as above

**Operation:** Export Typst code
```csharp
using var builder = PresentationBuilder.Open("05-typst-source.pptx");
string typstSource = builder.ExportToTypst();
await File.WriteAllTextAsync("05-source.typ", typstSource);
```

**Output:** `output/pptx/05-source.typ`

```typst
// Generated by OfficeEditor from PPTX

#set text(font: ("Arial", "Helvetica", "Liberation Sans"))

#set page(width: 720.00pt, height: 405.00pt, margin: 0pt)

#place(top + left, dx: 0.00pt, dy: 0.00pt)[#text(size: 18.00pt)[Typst Source Export]]
#place(top + left, dx: 0.00pt, dy: 56.69pt)[#text(size: 18.00pt)[View the generated code]]

#set page(width: 720.00pt, height: 405.00pt, margin: 0pt)

#pagebreak(to: "odd")

#place(top + left, dx: 0.00pt, dy: 0.00pt)[#text(size: 18.00pt)[How it works]]
#place(top + left, dx: 0.00pt, dy: 113.39pt)[#text(size: 18.00pt)[1. Extract PPTX content]]
...
```

---

## 6. Slide Thumbnails (PNG)

**Input:** PowerPoint file

**Operation:** Export each slide as PNG
```csharp
using var builder = PresentationBuilder.Open("presentation.pptx");
var thumbnails = builder.ExportThumbnails(new ThumbnailOptions { Ppi = 150 });

for (int i = 0; i < thumbnails.Length; i++)
{
    await File.WriteAllBytesAsync($"slide_{i + 1}.png", thumbnails[i]);
}
```

**Output:** `slide_1.png`, `slide_2.png`, ... (one PNG per slide at 150 PPI)

---

## JSON Edit Instructions

PPTX edit instructions target an existing deck. Use deck anatomy to obtain the 1-based slide number and element id, then parse an `operations` array with `PptxJsonInstructionParser`. From-scratch generation is a separate `version: "2.0"` vocabulary demonstrated by `demo/demo-deck.json`.

**Input:** `instructions/sample.json`

```json
{
  "operations": [
    { "type": "replaceText", "slide": 1, "elementId": 2, "text": "Q4 Review" },
    {
      "type": "replaceTable",
      "slide": 2,
      "elementId": 5,
      "rows": [["Quarter", "Revenue"], ["Q4", "$150K"]]
    },
    { "type": "duplicateSlide", "slide": 2, "position": 3 }
  ]
}
```

```csharp
using var builder = PresentationBuilder.Open("template.pptx");
var set = new PptxJsonInstructionParser().Parse(
    await File.ReadAllTextAsync("instructions/sample.json"));
var result = new PptxInstructionEngine().Apply(builder, set);
if (result.FailedOps.Count > 0) { /* inspect operation errors */ }
builder.Save("edited.pptx");
```

Supported operations are `replaceText`, `replaceImage`, `replaceTable`, `moveSlide`, `duplicateSlide`, and `deleteSlide`.

**Output:** `edited.pptx`, with the requested operations applied to the existing deck.

---

## Running Examples

```bash
cd examples
dotnet run
```

The runner executes PPTX sections 1–5. Section 6 and the JSON edit-instruction section are tested library snippets but are not yet called from `examples/Program.cs`.

**Prerequisites for Typst features:**
- TypstBridge native runtime asset for your RID (built/copied by the project where supported)
- Optional `typst` CLI in `PATH` for the external CLI fallback
