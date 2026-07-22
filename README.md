# OfficeEditor

A modern .NET 9 suite for creating and editing Office documents (DOCX, PPTX, XLSX) via instruction sets (JSON/YAML) or fluent C# APIs. Features markdown conversion, variable detection, publipostage (mail merge), and rich content blocks across all formats.

## Features

- **Multi-Format Support** - Word (DOCX), PowerPoint (PPTX), Excel (XLSX)
- **Create & Edit** - From scratch or existing documents
- **Fluent C# API** - Chain methods for intuitive document building
- **JSON/YAML Instructions** - Declarative document manipulation
- **Markdown Support** - Convert markdown to styled documents
- **Variable Detection** - Find `{{variables}}` in templates
- **Publipostage** - Batch generate documents from templates
- **Rich Content** - Tables, lists, headings, images, charts
- **Style Preservation** - Maintains existing document styles
- **Template Logic** - Conditionals and loops in templates
- **Typst Integration** - Export PPTX to PDF via Typst (PPTX only)

## Quick Start

### CLI

> **Note:** The `officeeditor` command is available from source via `dotnet run --project OfficeEditor.Cli -- [command]`. It is not yet published as a standalone dotnet tool.

```bash
# Create documents (auto-detects format from extension)
officeeditor create output.docx --text "Hello World"
officeeditor create output.pptx --title "My Presentation"
officeeditor create output.xlsx --sheet "Sales"

# Detect variables in templates
officeeditor detect template.docx
officeeditor detect template.pptx
officeeditor detect template.xlsx

# Merge template with data
officeeditor merge template.docx data.json output.docx
officeeditor merge template.pptx data.json output.pptx
officeeditor merge template.xlsx data.json output.xlsx

# Edit with instructions
officeeditor edit document.docx --instructions instructions.json
```

### C# API - Word (DOCX)

```csharp
using DocxEditor.Core.Builders;

// Create document
using var builder = DocumentBuilder.Create("output.docx");
builder.AddParagraph("Hello World", "Heading1");
builder.AddMarkdown("""
    # Title
    
    This is **bold** and *italic* text.
    
    - Item 1
    - Item 2
    """);
builder.Save();

// Edit existing document
using var builder = DocumentBuilder.Open("template.docx");
builder.ReplaceText("{{name}}", "John Doe");
builder.MergeVariables(new Dictionary<string, string>
{
    ["clientName"] = "Acme Corp",
    ["date"] = "2024-01-01"
});
builder.Save();
```

### C# API - PowerPoint (PPTX)

```csharp
using PptxEditor.Core.Builders;

// Create presentation
using var builder = PresentationBuilder.Create("output.pptx");

// Title slide
builder.AddSlide();
builder.CurrentSlide
    .AddTitle("My Presentation")
    .AddSubtitle("By John Doe");

// Content slide
builder.AddSlide();
builder.CurrentSlide
    .AddTitle("Agenda")
    .AddBulletList(new[] { "Item 1", "Item 2", "Item 3" });

// Slide with table
builder.AddSlide();
builder.CurrentSlide
    .AddTitle("Sales Data")
    .AddTable(new List<List<string>>
    {
        new() { "Product", "Q1", "Q2", "Q3" },
        new() { "Widget", "100", "200", "150" }
    });

builder.Save();

// Variable detection
var variables = builder.DetectVariables();

// Mail merge
builder.MergeVariables(new Dictionary<string, string>
{
    ["company"] = "Acme Corp"
});

// Export to PDF via Typst
var pdfBytes = builder.ExportToPdf();
await File.WriteAllBytesAsync("output.pdf", pdfBytes);

// Export slide thumbnails as PNG
var thumbnails = builder.ExportThumbnails(new ThumbnailOptions { Ppi = 150 });
for (int i = 0; i < thumbnails.Length; i++)
{
    await File.WriteAllBytesAsync($"slide_{i + 1}.png", thumbnails[i]);
}

// Export to Typst source
var typstSource = builder.ExportToTypst();
await File.WriteAllTextAsync("presentation.typ", typstSource);
```

### C# API - Excel (XLSX)

```csharp
using XlsxEditor.Core.Builders;

// Create workbook
using var builder = WorkbookBuilder.Create("output.xlsx");

// Add worksheet with data
var worksheet = builder.AddWorksheet("Sales");
worksheet
    .AddHeaderRow(new List<string> { "Product", "Q1", "Q2", "Q3", "Total" })
    .AddDataRow(new List<string> { "Widget", "100", "200", "150" }, 2)
    .AddDataRow(new List<string> { "Gadget", "50", "75", "100" }, 3)
    .AddFormulaRow(new List<string> { "", "=SUM(B2:B3)", "=SUM(C2:C3)", "=SUM(D2:D3)" }, 4);

// Add another worksheet
var summary = builder.AddWorksheet("Summary");
summary
    .AddCell("A1", "Total Sales")
    .AddCell("B1", "=SUM(Sales!B2:D10)", true);

builder.Save();

// Variable detection
var variables = builder.DetectVariables();

// Mail merge
builder.MergeVariables(new Dictionary<string, string>
{
    ["reportDate"] = "2024-01-01"
});

### C# API - In-memory / service usage

All builders support pathless creation and `byte[]`/`Stream` I/O, so you can generate documents in memory for web APIs, Azure Functions, or services. The existing file-path APIs remain unchanged.

Create a DOCX and return bytes:

```csharp
using DocxEditor.Core.Builders;
using OfficeEditor.Core.Services;

using var builder = DocumentBuilder.Create();
builder.AddParagraph("Hello");
BinaryOfficeDocument doc = builder.ToBinaryDocument();
return doc.Bytes;
```

Create a PPTX in memory:

```csharp
using PptxEditor.Core.Builders;
using OfficeEditor.Core.Services;

using var builder = PresentationBuilder.Create();
builder.AddSlide();
builder.CurrentSlide.AddTitle("Hello");
BinaryOfficeDocument doc = builder.ToBinaryDocument();
```

Create an XLSX in memory:

```csharp
using XlsxEditor.Core.Builders;
using OfficeEditor.Core.Services;

using var builder = WorkbookBuilder.Create();
builder.AddWorksheet("Sheet1").AddCell("A1", "Hello");
BinaryOfficeDocument doc = builder.ToBinaryDocument();
```

Return a generated file from an ASP.NET Core minimal API:

```csharp
app.MapGet("/report.docx", () =>
{
    using var builder = DocumentBuilder.Create();
    builder.AddParagraph("Hello");
    BinaryOfficeDocument doc = builder.ToBinaryDocument();
    return Results.File(doc.Bytes, doc.ContentType, $"report{doc.FileExtension}");
});
```

Roundtrip an existing document from bytes:

```csharp
byte[] templateBytes = await File.ReadAllBytesAsync("template.docx");

using var builder = DocumentBuilder.Open(templateBytes);
builder.ReplaceText("{{name}}", "John Doe");
BinaryOfficeDocument doc = builder.ToBinaryDocument();

await File.WriteAllBytesAsync("output.docx", doc.Bytes);
```

> **Stream ownership:** Each builder manages its own internal buffer. When you pass a stream to `Save(Stream)`, the builder writes to it and leaves it open so the caller can continue using it.

## Installation

> **Pre-1.0 Notice:** The packages are **not yet published on NuGet**. Build from source for now.

```bash
# Clone and build from source
git clone https://github.com/MaximeLeBesnerais/OfficeEditor.git
cd OfficeEditor
dotnet build

# Run the CLI from source (not yet available as a dotnet tool)
dotnet run --project OfficeEditor.Cli -- --help
```

**Planned NuGet packages** (coming soon):

| Package | Description |
|---------|-------------|
| `MaximeLB.OfficeEditor.Core` | Shared abstractions, services, variables |
| `MaximeLB.DocxEditor.Core` | Word (DOCX) creation & editing |
| `MaximeLB.PptxEditor.Core` | PowerPoint (PPTX) creation, editing & Typst export |
| `MaximeLB.XlsxEditor.Core` | Excel (XLSX) creation & editing |
| `MaximeLB.OfficeEditor.Cli` | Unified CLI tool |
| `MaximeLB.TypstBridge.Managed` | In-process Typst compiler bridge |

## Visual PDF Diffs

Use the visual diff helper to compare reference PDFs against generated PDFs as rasterized page images:

```bash
dotnet run --project tools/visual-diff -- --suite docx
```

The default report is written to `examples/output/visual-diff/docx/index.html`. See [`tools/visual-diff/README.md`](tools/visual-diff/README.md) for setup, arbitrary PDF pair usage, and how to read RMSE metrics.

## Project Structure

```
OfficeEditor/
├── OfficeEditor.Core/          # Shared abstractions
│   ├── Models/                 # VariableInfo, StyleMapping
│   ├── Services/               # TypstCompilerService
│   ├── Variables/              # VariableDetector, VariableReplacer, TemplateEngine
│   └── Exceptions/             # OfficeEditorException
├── DocxEditor.Core/            # Word (DOCX)
│   ├── Builders/               # DocumentBuilder fluent API
│   ├── Content/                # Content block rendering
│   ├── Markdown/               # Markdown parser
│   ├── Models/                 # ContentBlocks, Instructions
│   ├── Variables/              # DocxVariableDetector, DocxVariableReplacer
│   └── Instructions/           # Instruction engine
├── PptxEditor.Core/            # PowerPoint (PPTX)
│   ├── Builders/               # PresentationBuilder, SlideBuilder
│   ├── Converters/             # PptxToTypstConverter
│   ├── Models/                 # TypstModels, SlideAnatomy
│   └── Variables/              # PptxVariableDetector, PptxVariableReplacer
├── XlsxEditor.Core/            # Excel (XLSX)
│   ├── Builders/               # WorkbookBuilder, WorksheetBuilder
│   └── Variables/              # XlsxVariableDetector, XlsxVariableReplacer
├── OfficeEditor.Cli/           # Unified CLI (all formats)
├── DocxEditor.Cli/             # Word-specific CLI
└── DocxEditor.Tests/           # Unit tests (all formats)
```

## Core Concepts

### Variable Detection

Detect variables in templates across all formats:

```csharp
// Word
using var docx = DocumentBuilder.Open("template.docx");
var variables = docx.DetectVariables();

// PowerPoint
using var pptx = PresentationBuilder.Open("template.pptx");
var variables = pptx.DetectVariables();

// Excel
using var xlsx = WorkbookBuilder.Open("template.xlsx");
var variables = xlsx.DetectVariables();
```

All return `List<VariableInfo>`:
```csharp
public record VariableInfo
{
    public required string Name { get; init; }
    public required string FullMatch { get; init; }
    public required string Location { get; init; }
    public string? DefaultValue { get; init; }
}
```

### Publipostage (Mail Merge)

Replace variables with data across all formats:

```csharp
var data = new Dictionary<string, string>
{
    ["name"] = "John Doe",
    ["company"] = "Acme Corp",
    ["date"] = "2024-01-01"
};

// Works for DOCX, PPTX, XLSX
builder.MergeVariables(data);
```

### Typst Integration (PPTX → PDF)

Export PowerPoint presentations to PDF with high fidelity using Typst:

```csharp
using PptxEditor.Core.Builders;

// Open or create presentation
using var builder = PresentationBuilder.Open("presentation.pptx");

// Export to PDF
byte[] pdfBytes = builder.ExportToPdf();
await File.WriteAllBytesAsync("output.pdf", pdfBytes);

// Export slide thumbnails (PNG per slide)
var thumbnails = builder.ExportThumbnails(new ThumbnailOptions 
{ 
    Ppi = 150  // Resolution in pixels per inch
});

for (int i = 0; i < thumbnails.Length; i++)
{
    await File.WriteAllBytesAsync($"slide_{i + 1}.png", thumbnails[i]);
}

// Export to Typst source code
string typstSource = builder.ExportToTypst();
await File.WriteAllTextAsync("presentation.typ", typstSource);
```

**Features:**
- Converts PPTX slides to Typst pages
- Preserves text formatting (bold, italic, color, font size)
- Extracts and embeds images
- Handles tables with borders and cell formatting
- Extracts embedded fonts from PPTX for accurate rendering
- Compiles through `TypstCompilerService` using TypstBridge first
- Automatic font fallback if fonts are missing
- Positioning via Typst's `#place` function

> **Font fidelity note:** installing Microsoft Office fonts such as Aptos can improve
> PPTX rendering fidelity, and the converter will use installed fonts when it can
> resolve them. Exact PowerPoint/PDF parity is still not guaranteed: PowerPoint,
> Typst, and platform font engines can differ in font metrics, line breaking,
> hinting, and layout behavior even when the same font family is installed.

### Template Logic

Use conditionals and loops in templates:

```markdown
{{#if amount > 100}}
High value order: {{amount}}
{{/if}}

{{#ifnot isActive}}
Account is inactive
{{/ifnot}}

{{#each items}}
- {{name}}: ${{price}}
{{/each}}
```

```csharp
var data = new Dictionary<string, object>
{
    ["amount"] = 150,
    ["isActive"] = false,
    ["items"] = new List<Dictionary<string, object>>
    {
        new() { ["name"] = "Widget", ["price"] = 25 },
        new() { ["name"] = "Gadget", ["price"] = 50 }
    }
};

// Format-specific template engines
var engine = new DocxTemplateEngine();  // or PptxTemplateEngine, XlsxTemplateEngine
engine.Process(document, data);
```

## CLI Commands

> **Note:** The `officeeditor` and `docxeditor` commands are available from source (via `dotnet run --project OfficeEditor.Cli` / `dotnet run --project DocxEditor.Cli`). They are not yet published as standalone dotnet tools.

### Unified CLI (`officeeditor`)

| Command | Description |
|---------|-------------|
| `create` | Create new document (auto-detects format) |
| `edit` | Edit with instructions |
| `detect` | List variables in template |
| `merge` | Merge template with data |

```bash
# Create documents
officeeditor create report.docx --text "Annual Report" --style Heading1
officeeditor create slides.pptx --title "Q4 Review"
officeeditor create data.xlsx --sheet "Sales"

# Detect variables
officeeditor detect template.docx
officeeditor detect template.pptx
officeeditor detect template.xlsx

# Merge with data
officeeditor merge template.docx data.json output.docx
officeeditor merge template.pptx data.json output.pptx
officeeditor merge template.xlsx data.json output.xlsx
```

### Format-Specific CLIs

```bash
# Word
docxeditor create output.docx --text "Hello"
docxeditor edit input.docx --instructions ops.json
docxeditor markdown input.md output.docx

# PowerPoint (via unified CLI)
officeeditor create output.pptx --title "My Presentation"

# Excel (via unified CLI)
officeeditor create output.xlsx --sheet "Sheet1"
```

## Architecture

### Shared Components (OfficeEditor.Core)

- **VariableInfo** - Immutable record for variable metadata
- **StyleMapping** - Maps markdown elements to document styles
- **VariableDetector** - Generic regex-based variable scanning
- **VariableReplacer** - Generic text replacement engine
- **TemplateEngine** - Format-agnostic conditionals and loops
- **OfficeEditorException** - Base exception type

### Format-Specific Implementations

Each format has its own Core project that references OfficeEditor.Core:

- **DocxEditor.Core** - WordprocessingDocument, Body, Paragraph, Run
- **PptxEditor.Core** - PresentationDocument, Slide, ShapeTree
- **XlsxEditor.Core** - SpreadsheetDocument, Worksheet, SheetData

### Design Patterns

- **Instruction Pattern** - All operations modeled as immutable instruction objects
- **Builder Pattern** - Fluent API with method chaining
- **Strategy Pattern** - Format-specific variable detection/replacement
- **Repository Pattern** - Abstract document storage

## Dependencies

- **.NET 9**
- **DocumentFormat.OpenXml** - Microsoft OpenXML SDK
- **Markdig** - Markdown parser (DOCX only)
- **YamlDotNet** - YAML parser
- **Spectre.Console** - CLI output (optional)
- **TypstBridge.Managed** - Primary in-process Typst compiler bridge for PPTX PDF/SVG/PNG export

## Testing

```bash
dotnet test
```

1,218 tests across 4 test projects (xUnit), 0 failures, covering:
- Document creation and manipulation (DOCX, PPTX, XLSX)
- Content block rendering
- Markdown conversion
- Variable detection and replacement
- Serialization
- Slide management (PPTX)
- Worksheet operations (XLSX)
- **Typst integration (PPTX)**
  - PDF export
  - Thumbnail generation
  - Typst source generation
  - Font extraction and fallback

### Coverage

Measured with coverlet + ReportGenerator on the main test suite (DocxEditor.Tests, which exercises all Core assemblies): **89.2% line coverage** (15,643 / 17,524 lines), **74.8% branch coverage**, **93.7% method coverage**. Per-assembly line coverage: PptxEditor.Core 90.3%, XlsxEditor.Core 92.3%, DocxEditor.Core 89.5%, OfficeEditor.Core 81.4%, TypstBridge.Managed 58% (native interop layer).

```bash
dotnet test --collect:"XPlat Code Coverage" --results-directory coverage
reportgenerator "-reports:coverage/*/coverage.cobertura.xml" "-targetdir:coverage/report" "-reporttypes:TextSummary"
```

(`reportgenerator` is a dotnet global tool: `dotnet-reportgenerator-globaltool`.)

## Performance

PPTX render pipeline (PptxEditor → Typst → PNG/PDF) benchmarked against headless LibreOffice
(`tools/pptx-benchmark`, median of 5 warm runs per deck, Apple Silicon, .NET 9 — single machine,
treat as orders of magnitude):

| Deck | Slides | OfficeEditor warm (total) | Per slide | LibreOffice warm (total) | Per slide | Speedup |
|---|---|---|---|---|---|---|
| .pptx | 16 | 189.5 ms | 11.8 ms | 2,661.5 ms | 166.3 ms | ~14× |

Preview path = open + whole-deck PNG @150ppi. Product target: <500 ms per slide — comfortably met.
The LibreOffice leg (`soffice --headless --convert-to pdf`) includes full process start and profile
cost. Cold starts (fresh process, JIT + backend probes): ~650–800 ms per deck.

Reproduce: `dotnet run --project tools/pptx-benchmark` (full methodology and limitations in
`tools/pptx-benchmark/README.md`; latest full report: `examples/output/benchmark/report.md`,
gitignored).

## Typst Compilation Backend

`TypstCompilerService` uses TypstBridge as the primary in-process backend for PPTX exports. TypstBridge supports PDF, SVG, PNG, multi-page outputs, working-directory assets, explicit font paths, PNG PPI, and diagnostics.

The external `typst` CLI remains a safety net when native compilation is unavailable. Install the `typst` binary if you need CLI fallback support in your environment.

## Status & Limitations

**Version:** 0.1.0 (pre-1.0). The public API may change between versions.

**Platform Support:** pure managed .NET 9 — no platform-specific code, so the suite is expected to run anywhere .NET runs.
- **macOS (arm64)** — Current development platform; tests, benchmarks and coverage actively run here.
- **linux-x64** — Previously tested.
- **Windows** — Expected to work, not yet verified.

**Rendering (PPTX → Typst → PDF/PNG/SVG):**
- Text, images, shapes, and tables render with good fidelity.
- Charts, SmartArt, and animations have **partial support** — complex charts may render as simplified representations or be omitted entirely.
- **Font fidelity** varies by platform. Installing Microsoft Office fonts (e.g., Aptos, Calibri) improves visual accuracy, but exact parity with PowerPoint is not guaranteed due to differences in font metrics, line breaking, and layout engines between PowerPoint and Typst.

**Packages:**
- Not yet published on NuGet. See [Installation](#installation) for source-build instructions.
- The `officeeditor` CLI is not yet published as a dotnet tool. Use `dotnet run --project OfficeEditor.Cli` instead.

## License

MIT - Copyright 2026 Maxime Le Besnerais

## Contributing

1. Fork the repository
2. Create a feature branch (`git checkout -b feature/amazing-feature`)
3. Commit changes (`git commit -m 'feat: add amazing feature'`)
4. Push to branch (`git push origin feature/amazing-feature`)
5. Open a Pull Request

## Author

**Maxime Le Besnerais**
