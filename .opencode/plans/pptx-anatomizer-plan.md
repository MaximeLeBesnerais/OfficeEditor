# PPTX Anatomizer - Implementation Plan

## Overview

Build a PPTX anatomizer that scans a presentation to detect all editable areas (text shapes, tables, images) and allows targeted replacement by element ID/name.

## Architecture

### New Models

```csharp
// PptxEditor.Core/Models/SlideElement.cs
public record SlideElement
{
    public required string Type { get; init; } // "Text", "Table", "Image"
    public required uint Id { get; init; }
    public required string Name { get; init; }
    public required string Location { get; init; } // "slide:1:element:Title"
    public string? Text { get; init; }
    public List<List<string>>? TableData { get; init; }
    public string? ImagePath { get; init; } // For extracted images
}

public record SlideAnatomy
{
    public int SlideIndex { get; init; }
    public List<SlideElement> Elements { get; init; } = new();
}
```

### New Service: PptxAnatomizer

```csharp
// PptxEditor.Core/Services/PptxAnatomizer.cs
public class PptxAnatomizer
{
    public List<SlideAnatomy> Analyze(PresentationDocument document);
    public void ReplaceElement(SlidePart slidePart, uint elementId, SlideElement newElement);
    public void ReplaceTextInElement(SlidePart slidePart, uint elementId, string newText);
    public void ReplaceTableData(SlidePart slidePart, uint elementId, List<List<string>> newData);
    public void ReplaceImage(SlidePart slidePart, uint elementId, string newImagePath);
}
```

## Implementation Details

### 1. Element Detection

Iterate `ShapeTree.ChildElements` and handle all types:

- **`P.Shape`** → Type: "Text", extract text from `TextBody`
- **`P.GraphicFrame`** → Check `GraphicData` for `A.Table`, Type: "Table", extract rows/cells
- **`P.Picture`** → Type: "Image", extract image part via `Blip.Embed`
- **`P.GroupShape`** → Recurse into child elements

For each element, read:
- `NonVisualDrawingProperties.Id` → uint Id
- `NonVisualDrawingProperties.Name` → string Name

### 2. Element Replacement

Find element by ID in ShapeTree:
- **Text**: Replace `TextBody` content preserving structure
- **Table**: Replace inner `A.Table` rows/cells
- **Image**: Replace `ImagePart` data, keep same relationship ID

### 3. API Integration

Add to `IPresentationBuilder`:
```csharp
List<SlideAnatomy> Analyze();
void ReplaceElement(uint elementId, string newText);
void ReplaceTable(uint elementId, List<List<string>> newData);
void ReplaceImage(uint elementId, string newImagePath);
```

## Tests

1. **Analyze** detects text shapes, tables, images
2. **ReplaceText** updates text preserving formatting
3. **ReplaceTable** updates table data
4. **ReplaceImage** updates image binary
5. **Element lookup by ID** works correctly

## Files to Create/Modify

**New:**
- `PptxEditor.Core/Models/SlideElement.cs`
- `PptxEditor.Core/Services/PptxAnatomizer.cs`
- `DocxEditor.Tests/Unit/PptxAnatomizerTests.cs`

**Modify:**
- `PptxEditor.Core/Builders/PresentationBuilder.cs` (add Analyze/Replace methods)
- `PptxEditor.Core/Builders/SlideBuilder.cs` (add Replace methods)

## Out of Scope

- PDF export (requires external renderer)
- Thumbnail generation (requires external renderer)
- DOCX anatomizer (future phase)
- SmartArt detection (complex, future phase)

## Commit Strategy

1. `feat: add SlideElement and SlideAnatomy models`
2. `feat: implement PptxAnatomizer with element detection`
3. `feat: add element replacement methods`
4. `feat: integrate anatomizer into PresentationBuilder`
5. `test: add PptxAnatomizer tests`
