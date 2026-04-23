# OfficeEditor Examples

This folder contains practical examples demonstrating how to use the OfficeEditor libraries for creating and manipulating DOCX, PPTX, and XLSX files.

## Structure

```
examples/
├── Docx/               # Word document examples
│   ├── README.md       # This file
│   ├── Program.cs      # DOCX examples runner
│   ├── sample.md       # Sample markdown file
│   └── instructions/   # JSON instruction examples
├── Pptx/               # PowerPoint examples
│   ├── README.md
│   ├── Program.cs
│   └── instructions/
├── Xlsx/               # Excel examples
│   ├── README.md
│   ├── Program.cs
│   └── instructions/
├── Shared/             # Shared resources
│   ├── data.json       # Sample data for templates
│   └── images/         # Sample images
└── Program.cs          # Main entry point (runs all examples)
```

## Quick Start

### Run All Examples

```bash
cd examples
dotnet run
```

### Run Specific Examples

```bash
# DOCX only
dotnet run --project Docx

# PPTX only
dotnet run --project Pptx

# XLSX only
dotnet run --project Xlsx
```

## What You'll Learn

### DOCX Examples
- Creating documents from scratch
- Converting Markdown to DOCX
- Using JSON instructions for document manipulation
- Variable detection and replacement
- Template processing with conditionals and loops

### PPTX Examples
- Creating presentations with multiple slides
- Adding various content types (text, tables, images, charts)
- **Typst Integration**: Export to PDF
- **Typst Integration**: Generate slide thumbnails
- **Typst Integration**: Export to Typst source code
- Variable detection in presentations

### XLSX Examples
- Creating workbooks with multiple worksheets
- Adding data, formulas, and formatting
- Variable detection in spreadsheets
- Template processing

## Prerequisites

- .NET 9 SDK
- OfficeEditor packages (referenced in example projects)

## Output

All generated files are saved to `examples/output/` directory.

## License

Same as main project - MIT License
