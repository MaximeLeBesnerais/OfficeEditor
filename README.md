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
```

## Installation

```bash
# Core libraries (pick what you need)
dotnet add package OfficeEditor.Core
dotnet add package DocxEditor.Core
dotnet add package PptxEditor.Core
dotnet add package XlsxEditor.Core

# CLI tools
dotnet tool install OfficeEditor.Cli
dotnet tool install DocxEditor.Cli
```

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
- Automatic font fallback if fonts are missing
- Positioning via Typst's `#place` function

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
- **typstsharp** - Typst compiler for PDF export (PPTX only)

## Testing

```bash
dotnet test
```

80+ unit tests covering:
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

## TODO

### Build Native Typst Wrapper from Source

**Status:** Currently using CLI fallback (`typst` binary in PATH). Native library loading works on Windows and some Linux distros, but fails on hardened kernels (e.g., Artix/Arch) due to the bundled `libtypst_core.so` missing the `-z noexecstack` linker flag.

**Plan:** Build our own .NET-native wrapper around Typst's C API (`libtypst`) compiled from source with proper flags:
```bash
# Rust build with correct stack flags
RUSTFLAGS="-C link-arg=-z -C link-arg=noexecstack" cargo build --release
```

**Why this matters:**
- CLI fallback adds ~500ms-2s per compilation (process spawn overhead)
- Native wrapper would be near-instant (in-process)
- Removes the external `typst` dependency for end users
- Works on all Linux distros regardless of kernel hardening

**Current workaround:** `TypstCompilerService` automatically falls back to the `typst` CLI when the native library fails to load. Install `typst` via your package manager and everything works.

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
