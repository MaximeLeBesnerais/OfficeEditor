# OfficeEditor Suite - Phase 3 Implementation Plan

## Overview

Transform DocxEditor into a full Office document editing suite supporting **Word (DOCX)**, **PowerPoint (PPTX)**, and **Excel (XLSX)**. Each format gets its own dedicated project with shared components extracted into a common base.

## Architecture

```
OfficeEditor/
├── OfficeEditor.Core/          # Shared abstractions (NEW)
│   ├── Models/
│   │   ├── VariableInfo.cs
│   │   └── StyleMapping.cs
│   ├── Variables/
│   │   ├── VariableDetector.cs
│   │   ├── VariableReplacer.cs
│   │   └── TemplateEngine.cs
│   ├── Serialization/
│   │   ├── JsonParser.cs
│   │   └── YamlParser.cs
│   └── Exceptions/
│       └── OfficeEditorException.cs
├── DocxEditor.Core/            # Word (refactored to use shared)
├── PptxEditor.Core/            # PowerPoint (NEW)
├── XlsxEditor.Core/            # Excel (NEW)
├── OfficeEditor.Cli/           # Unified CLI (NEW)
├── docxeditor/                 # Word-specific CLI (symlink/wrapper)
├── pptxeditor/                 # PowerPoint-specific CLI
├── xlsxeditor/                 # Excel-specific CLI
└── OfficeEditor.Tests/         # Unified tests
    ├── Docx/
    ├── Pptx/
    └── Xlsx/
```

## Key Decisions

1. **Naming**: Rebrand to `OfficeEditor` (repo rename later)
2. **CLI**: 4 CLIs total - 1 unified (`officeeditor`) + 3 format-specific
3. **Markdown for PPTX**: Yes, map `# Title` → title slide, `## Slide` → content slides
4. **Charts**: Simplified `ChartBuilder` API for V1, TODO for raw OpenXML escape hatch
5. **Images**: Format-specific implementations (different APIs for DOCX vs PPTX vs XLSX)
6. **Priority**: PPTX first, then XLSX
7. **Breaking changes**: Allowed - refactor existing code freely

## Shared Components (OfficeEditor.Core)

### Extracted from existing DocxEditor.Core:
- `VariableInfo` → shared model
- `StyleMapping` → shared model  
- `VariableDetector` → generic regex scanning
- `VariableReplacer` → generic text replacement
- `TemplateEngine` → format-agnostic conditionals/loops
- `JsonInstructionParser` / `YamlInstructionParser` → generic parsing utilities
- `DocxEditorException` → rename to `OfficeEditorException`

### New shared abstractions:
```csharp
public interface IOfficeDocumentBuilder : IDisposable
{
    void Save(string? path = null);
}

public interface IVariableSupport
{
    List<VariableInfo> DetectVariables();
    void MergeVariables(Dictionary<string, string> data);
}

public interface IInstructionSupport
{
    void ExecuteInstructions(DocumentInstructions instructions);
}
```

## PowerPoint (PPTX) Design

### Builder API:
```csharp
using var builder = PresentationBuilder.Create("output.pptx");

// Slide with explicit layout
builder.AddSlide("Title Slide")
    .AddTitle("My Presentation")
    .AddSubtitle("By John Doe");

// Auto-detected layout (title + content = "Title and Content")
builder.AddSlide()
    .AddTitle("Agenda")
    .AddBulletList(new[] { "Item 1", "Item 2", "Item 3" });

// With chart
builder.AddSlide()
    .AddTitle("Sales Data")
    .AddChart(ChartType.Bar, new Dictionary<string, int> {
        ["Q1"] = 100, ["Q2"] = 200, ["Q3"] = 150
    });

builder.Save();
```

### Content Blocks:
```csharp
public abstract record SlideBlock;
public record TitleBlock : SlideBlock { public string Text { get; init; } }
public record BodyTextBlock : SlideBlock { public string Text { get; init; } }
public record BulletListBlock : SlideBlock { public List<string> Items { get; init; } }
public record ImageBlock : SlideBlock { public string Path { get; init; } }
public record ChartBlock : SlideBlock { 
    public ChartType Type { get; init; }
    public Dictionary<string, object> Data { get; init; }
}
public record TableBlock : SlideBlock { public List<List<string>> Rows { get; init; } }
```

### Variable Detection:
- Scan all text in shapes across all slides
- Same `{{var|default}}` syntax
- Location: `slide:3:shape:Title 1`

### Markdown-to-PPTX:
```markdown
# My Presentation
## By John Doe

## Agenda
- Item 1
- Item 2

## Sales Data
| Q1 | Q2 | Q3 |
|----|----|----|
| 100| 200| 150|
```
Maps to:
- `# My Presentation` → Title slide
- `## Agenda` → New content slide with title
- `- Item 1` → Bullet points
- `## Sales Data` → New slide with table

## Excel (XLSX) Design

### Builder API:
```csharp
using var builder = WorkbookBuilder.Create("output.xlsx");

// Add worksheet
builder.AddWorksheet("Sales")
    .AddHeaderRow(new[] { "Product", "Q1", "Q2", "Q3", "Total" })
    .AddDataRow(new[] { "Widget", "100", "200", "150" })
    .AddFormulaRow(new[] { "", "=SUM(B2:D2)" })
    .AddTable("A1:D10", "SalesTable")
    .AddChart(ChartType.Line, "A1:D10");

// Another worksheet
builder.AddWorksheet("Summary")
    .AddCell("A1", "Total Sales")
    .AddCell("B1", "=SUM(Sales!B2:D10)", "Currency");

builder.Save();
```

### Content Blocks:
```csharp
public abstract record CellBlock;
public record ValueBlock : CellBlock { public string Value { get; init; } }
public record FormulaBlock : CellBlock { public string Formula { get; init; } }
public record HeaderBlock : CellBlock { public string Text { get; init; } }
public record TableBlock : CellBlock { public List<List<string>> Rows { get; init; } }
public record ChartBlock : CellBlock { public ChartType Type { get; init; } }
```

### Variable Detection:
- Scan all cell values (not formulas)
- Same `{{var|default}}` syntax
- Location: `sheet:Sales:cell:A1`

## ChartBuilder (Simplified)

For V1, wrap the verbose OpenXML chart API:

```csharp
public class ChartBuilder
{
    public static ChartPart CreateBarChart(ChartPart chartPart, string title, Dictionary<string, int> data)
    {
        // ~50 lines instead of 200+
        // Handles: ChartSpace, PlotArea, BarChart, Axes, Legend
    }
    
    public static ChartPart CreateLineChart(ChartPart chartPart, string title, Dictionary<string, int> data)
    {
        // Similar simplification
    }
    
    public static ChartPart CreatePieChart(ChartPart chartPart, string title, Dictionary<string, int> data)
    {
        // Similar simplification
    }
}
```

TODO for V2: Raw OpenXML escape hatch for complex charts.

## CLI Design

### Unified CLI (`officeeditor`):
```bash
# Auto-detects format from extension
officeeditor create output.pptx --title "My Presentation"
officeeditor edit input.xlsx --instructions ops.json
officeeditor detect template.docx
officeeditor merge template.pptx data.json output.pptx
```

### Format-specific CLIs:
```bash
# Word
docxeditor create output.docx --text "Hello"
docxeditor edit input.docx --instructions ops.json

# PowerPoint  
pptxeditor create output.pptx --title "My Presentation"
pptxeditor edit input.pptx --instructions ops.json

# Excel
xlsxeditor create output.xlsx --sheet "Sheet1"
xlsxeditor edit input.xlsx --instructions ops.json
```

Format detection: Extension first, content sniffing as fallback.

## Implementation Order

### Phase 3.1: Shared Foundation
1. Create `OfficeEditor.Core` project
2. Move shared models, variables, serialization, exceptions
3. Refactor `DocxEditor.Core` to reference shared project
4. Update all existing tests to pass

### Phase 3.2: PowerPoint Foundation  
1. Create `PptxEditor.Core` project
2. Implement `PresentationBuilder` (Create/Open/Save)
3. Implement slide management (add/remove/reorder)
4. Implement basic shapes (text box, placeholder)
5. Add slide layout support
6. Add theme support
7. Write tests

### Phase 3.3: PowerPoint Content
1. Implement slide content blocks
2. Add text formatting
3. Add bullet/numbered lists
4. Add table support
5. Add image support
6. Add simplified ChartBuilder
7. Write tests

### Phase 3.4: PowerPoint Variables & Templates
1. Implement variable detection
2. Implement variable replacement
3. Add template logic
4. Add publipostage
5. Write tests

### Phase 3.5: Excel Foundation
1. Create `XlsxEditor.Core` project
2. Implement `WorkbookBuilder`
3. Implement worksheet management
4. Implement cell operations
5. Write tests

### Phase 3.6: Excel Content
1. Implement cell content blocks
2. Add table support
3. Add chart support
4. Add formula support
5. Add styling
6. Write tests

### Phase 3.7: Excel Variables & Templates
1. Implement variable detection
2. Implement variable replacement
3. Add template logic
4. Add publipostage
5. Write tests

### Phase 3.8: CLI
1. Create unified `OfficeEditor.Cli`
2. Create format-specific wrappers
3. Add format detection
4. Write tests

### Phase 3.9: Documentation
1. Update README for all formats
2. Add PPTX examples
3. Add XLSX examples
4. Update architecture docs

## Commit Strategy

1. `feat: create OfficeEditor.Core shared project`
2. `refactor: extract shared components from DocxEditor.Core`
3. `feat: create PptxEditor.Core project structure`
4. `feat: implement PresentationBuilder`
5. `feat: add PPTX slide management`
6. `feat: add PPTX content blocks`
7. `feat: add PPTX variable detection`
8. `feat: add simplified ChartBuilder`
9. `feat: create XlsxEditor.Core project structure`
10. `feat: implement WorkbookBuilder`
11. `feat: add XLSX cell operations`
12. `feat: add XLSX variable detection`
13. `feat: create unified OfficeEditor.Cli`
14. `test: add comprehensive tests for all formats`
15. `docs: update README and examples`

## Open Questions

None remaining - all decisions made.

Ready to proceed with implementation!
