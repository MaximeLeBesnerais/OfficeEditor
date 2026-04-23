# PPTX Examples

This folder contains examples for creating and editing PowerPoint presentations (.pptx), including the new Typst integration features.

## Examples

### 1. Basic Presentation Creation
Create a presentation with multiple slides and content types.

### 2. Advanced Slide Content
Add tables, images, charts, and formatted text.

### 3. Variable Detection & Mail Merge
Find and replace variables in presentations.

### 4. **Typst Export to PDF**
Export presentations to high-fidelity PDF using Typst.

### 5. **Typst Slide Thumbnails**
Generate PNG thumbnails of each slide.

### 6. **Typst Source Export**
Export presentation to Typst source code.

## Running Examples

```bash
cd examples/Pptx
dotnet run
```

Output files are saved to `examples/output/pptx/`.

## Typst Integration Notes

The Typst integration requires:
- `typstsharp` NuGet package (included)
- Native `libtypst_core.so` binary (auto-copied on Linux)

**Features:**
- Converts PPTX slides to Typst pages
- Preserves text formatting (bold, italic, color, font size)
- Extracts and embeds images
- Handles tables with borders and cell formatting
- Extracts embedded fonts for accurate rendering
- Automatic font fallback
