# DocxEditor - Phase 2 Plan: Rich Content, Markdown & Publipostage

## Overview
Extend DocxEditor to support rich content insertion (headers, lists, tables), markdown-to-docx conversion with custom style mapping, variable detection, and publipostage (mail merge).

## User Requirements

1. **Custom markdown extensions**: `:::tip`, `:::warning`, `:::note` etc. map to Word styles (e.g., `Tip`, `Warning`, `Note`) in the document
2. **Variable detection**: Full document scan (body + headers + footers) for `{{var}}` patterns
3. **Default values**: Support `{{name|Guest}}` syntax
4. **Conditional blocks**: Full logic with comparisons and loops

## New Features

### 1. Rich Content Blocks
Replace simple text with structured content blocks that map to docx elements:
- **Paragraphs** with inline formatting (bold, italic, underline, strikethrough, code)
- **Headings** (levels 1-6)
- **Bullet lists** and **numbered lists**
- **Tables** with rows/cells
- **Blockquotes**
- **Code blocks**
- **Horizontal rules**

### 2. Markdown-to-Docx Conversion
Parse markdown text and convert to proper docx structure:
```markdown
# Title (maps to Heading1)
## Subtitle (maps to Heading2)

Normal paragraph with **bold** and *italic* text.

- Bullet item 1
- Bullet item 2

1. Numbered item 1
2. Numbered item 2

> Blockquote text

| Column 1 | Column 2 |
|----------|----------|
| Cell 1   | Cell 2   |

:::tip
This is a tip that maps to "Tip" style in Word
:::

:::warning
This is a warning that maps to "Warning" style in Word
:::
```

**Style Mapping Configuration:**
```json
{
  "styleMap": {
    "heading1": "CustomTitle",
    "heading2": "CustomSubtitle",
    "paragraph": "Normal",
    "blockquote": "Quote",
    "code": "Code",
    "tip": "Tip",
    "warning": "Warning",
    "note": "Note"
  }
}
```

### 3. Variable Detection
Scan entire document (body + headers + footers) for `{{variableName}}` or `{{variableName|defaultValue}}` patterns:
```csharp
var detector = new VariableDetector();
var variables = detector.Scan(document);
// Returns: [
//   { name: "clientName", fullMatch: "{{clientName}}", location: "header", defaultValue: null },
//   { name: "guestName", fullMatch: "{{guestName|Guest}}", location: "body", defaultValue: "Guest" }
// ]
```

### 4. Publipostage (Mail Merge)
Replace variables with data from JSON:
```csharp
// Single record
var data = new Dictionary<string, string> {
    ["clientName"] = "John Doe",
    ["company"] = "Acme Corp"
};
document.Merge(data);

// Batch generation
var records = JsonSerializer.Deserialize<List<Dictionary<string, string>>>(json);
document.MergeBatch(records, "output_{index}.docx");
```

### 5. Conditional Blocks
Full template logic with comparisons and loops:
```markdown
{{#if amount > 100}}
High value order
{{/if}}

{{#ifnot isActive}}
Account is inactive
{{/ifnot}}

{{#each items}}
- {{name}}: {{price}}
{{/each}}
```

## Architecture Changes

### New Models
```
DocxEditor.Core/Models/
├── ContentBlocks.cs          # Rich content block types
├── MarkdownDocument.cs       # Markdown AST nodes
├── VariableInfo.cs           # Variable detection results
├── StyleMapping.cs           # Custom style mappings
└── TemplateLogic.cs          # Conditional/loop nodes
```

### New Services
```
DocxEditor.Core/
├── Markdown/
│   ├── MarkdownParser.cs      # Parse markdown to AST
│   ├── MarkdownToDocxConverter.cs  # Convert AST to docx elements
│   └── StyleMapper.cs         # Map markdown elements to docx styles
├── Variables/
│   ├── VariableDetector.cs    # Scan for {{var}} patterns
│   ├── VariableReplacer.cs    # Replace variables with data
│   └── TemplateEngine.cs      # Process conditionals and loops
├── Content/
│   └── ContentBlockBuilder.cs # Build rich content blocks
└── Extensions/
    └── CustomMarkdownExtensions.cs  # :::tip, :::warning etc.
```

### Updated Builders
```csharp
public interface IDocumentBuilder : IDisposable
{
    // Existing methods...
    
    // Rich content
    IDocumentBuilder AddRichContent(List<ContentBlock> blocks);
    IDocumentBuilder ReplaceWithRichContent(string targetText, List<ContentBlock> blocks);
    
    // Markdown
    IDocumentBuilder AddMarkdown(string markdown, StyleMapping? styleMap = null);
    IDocumentBuilder ReplaceWithMarkdown(string targetText, string markdown, StyleMapping? styleMap = null);
    
    // Variables
    List<VariableInfo> DetectVariables();
    IDocumentBuilder MergeVariables(Dictionary<string, string> data);
    
    // Publipostage
    void MergeBatch(List<Dictionary<string, string>> records, string outputPattern);
}
```

## Implementation Phases

### Phase 2.1: Content Blocks
- [ ] Define ContentBlock types (Paragraph, Heading, List, Table, etc.)
- [ ] Implement ContentBlockBuilder for fluent construction
- [ ] Add AddRichContent and ReplaceWithRichContent methods
- [ ] Support inline formatting (bold, italic, underline, strikethrough, code)
- [ ] Tests for content block generation

### Phase 2.2: Markdown Parser
- [ ] Integrate Markdig (or custom parser) for markdown parsing
- [ ] Create markdown AST visitor pattern
- [ ] Implement MarkdownToDocxConverter
- [ ] Support all standard markdown elements
- [ ] Tests for markdown conversion

### Phase 2.3: Custom Markdown Extensions
- [ ] Implement :::tip, :::warning, :::note syntax
- [ ] Map custom blocks to Word styles
- [ ] Support user-defined custom extensions
- [ ] Tests for custom extensions

### Phase 2.4: Style Mapping
- [ ] Create StyleMapping configuration class
- [ ] Load style mappings from JSON/YAML
- [ ] Map markdown elements to custom docx styles
- [ ] Validate style existence in target document
- [ ] Tests for style mapping

### Phase 2.5: Variable Detection
- [ ] Implement VariableDetector scanning body + headers + footers
- [ ] Support custom variable patterns (default: `{{var}}`)
- [ ] Support default values (`{{var|default}}`)
- [ ] Return variable locations (body/header/footer/paragraph index)
- [ ] Tests for variable detection

### Phase 2.6: Publipostage
- [ ] Implement VariableReplacer for single record
- [ ] Add MergeBatch for multiple records
- [ ] Support output filename patterns (`output_{index}.docx`, `output_{clientName}.docx`)
- [ ] Handle missing variables (use default, empty string, or error)
- [ ] Tests for mail merge

### Phase 2.7: Template Logic (Conditionals & Loops)
- [ ] Implement TemplateEngine for {{#if}}, {{#ifnot}}
- [ ] Add comparison operators (>, <, ==, !=)
- [ ] Add {{#each}} loops for arrays
- [ ] Support nested conditionals
- [ ] Tests for template logic

### Phase 2.8: CLI Updates
- [ ] Add `detect` command to list variables in a document
- [ ] Add `merge` command for publipostage
- [ ] Add `markdown` command to create document from markdown file
- [ ] Add `--style-map` parameter for custom style mappings
- [ ] Update instruction schema to support rich content and markdown

### Phase 2.9: Integration
- [ ] End-to-end tests: markdown → docx with custom styles
- [ ] End-to-end tests: template + data → multiple documents
- [ ] End-to-end tests: conditional blocks in templates
- [ ] Performance tests for large batch generation
- [ ] Update documentation and examples

## New Instruction Types

```json
{
  "operations": [
    {
      "type": "addRichContent",
      "blocks": [
        {
          "type": "heading",
          "level": 1,
          "text": "Document Title",
          "style": "CustomTitle"
        },
        {
          "type": "paragraph",
          "text": "Hello {{clientName}}",
          "inline": [
            { "type": "bold", "text": "Important" },
            { "type": "text", "text": " notice" }
          ]
        },
        {
          "type": "list",
          "ordered": false,
          "items": ["Item 1", "Item 2", "Item 3"]
        }
      ]
    },
    {
      "type": "addMarkdown",
      "markdown": "# Title\n\nSome **bold** text\n\n:::tip\nThis is a tip\n:::",
      "styleMap": {
        "heading1": "CustomTitle",
        "tip": "TipStyle"
      }
    },
    {
      "type": "mergeVariables",
      "data": {
        "clientName": "John Doe",
        "company": "Acme Corp"
      }
    }
  ]
}
```

## CLI Commands

```bash
# Detect variables in document
docxeditor detect template.docx
# Output: Found 3 variables: clientName, company, date

# Merge single record
docxeditor merge template.docx data.json output.docx

# Batch merge
docxeditor merge template.docx batch_data.json "output_{clientName}.docx"

# Create from markdown
docxeditor markdown input.md output.docx --style-map styles.json

# Edit with rich content instructions
docxeditor edit template.docx --instructions rich_content.json
```

## Dependencies to Add
- **Markdig** (or similar) - Markdown parser
- Optional: **System.Text.RegularExpressions** for variable detection (already available)

## Commit Strategy
1. `feat: add content block models and builder`
2. `feat: implement markdown parser and converter`
3. `feat: add custom markdown extensions (:::tip, :::warning)`
4. `feat: add style mapping configuration`
5. `feat: implement variable detection with default values`
6. `feat: add publipostage batch generation`
7. `feat: implement template logic (conditionals and loops)`
8. `feat: update CLI with detect, merge, and markdown commands`
9. `test: add integration tests for rich content and mail merge`
10. `docs: add examples and usage documentation`
