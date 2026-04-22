# DocxEditor

A modern .NET 9 library and CLI tool for creating and editing DOCX files via instruction sets (JSON/YAML) or a fluent C# API. Features markdown-to-docx conversion, variable detection, publipostage (mail merge), and rich content blocks.

## Features

- **Create & Edit DOCX** - From scratch or existing documents
- **Fluent C# API** - Chain methods for intuitive document building
- **JSON/YAML Instructions** - Declarative document manipulation
- **Markdown Support** - Convert markdown to styled Word documents
- **Variable Detection** - Find `{{variables}}` in templates
- **Publipostage** - Batch generate documents from templates
- **Rich Content** - Tables, lists, headings, blockquotes, code blocks
- **Style Preservation** - Maintains existing document styles
- **Template Logic** - Conditionals and loops in templates

## Quick Start

### CLI

```bash
# Create from markdown
docxeditor markdown input.md output.docx --style-map styles.json

# Detect variables in template
docxeditor detect template.docx

# Merge template with data
docxeditor merge template.docx data.json output.docx

# Edit with instructions
docxeditor edit document.docx --instructions instructions.json
```

### C# API

```csharp
using DocxEditor.Core.Builders;
using DocxEditor.Core.Models;

// Create document from markdown
using var builder = DocumentBuilder.Create("output.docx");
builder.AddMarkdown("""
    # Hello World
    
    This is **bold** and *italic* text.
    
    - Item 1
    - Item 2
    """);
builder.Save();

// Edit existing document
using var builder = DocumentBuilder.Open("template.docx");
builder.ReplaceText("{{name}}", "John Doe");
builder.AddParagraph("New paragraph", "Heading1");
builder.Save();

// Variable detection
var variables = builder.DetectVariables();
// Returns: [{ Name: "name", DefaultValue: null, Location: "body" }]

// Mail merge
builder.MergeVariables(new Dictionary<string, string>
{
    ["clientName"] = "Acme Corp",
    ["date"] = "2024-01-01"
});
```

## Installation

```bash
dotnet add package DocxEditor.Core
```

## Project Structure

```
DocxEditor/
├── DocxEditor.Core/           # Core library
│   ├── Builders/              # DocumentBuilder fluent API
│   ├── Content/               # Content block rendering
│   ├── Markdown/              # Markdown parser
│   ├── Models/                # Data models
│   ├── Serialization/         # JSON/YAML parsers
│   ├── Variables/             # Variable detection & replacement
│   └── Instructions/          # Instruction engine
├── DocxEditor.Cli/            # CLI application
└── DocxEditor.Tests/          # Unit tests
```

## Core Concepts

### Document Builder

The `DocumentBuilder` is the main entry point for document manipulation:

```csharp
public interface IDocumentBuilder : IDisposable
{
    // Basic operations
    IDocumentBuilder AddParagraph(string text, string? style = null);
    IDocumentBuilder InsertAfter(string targetText, string text, string? style = null);
    IDocumentBuilder InsertBefore(string targetText, string text, string? style = null);
    IDocumentBuilder ReplaceText(string find, string replace);
    IDocumentBuilder ReplaceParagraph(string targetText, string newText, string? style = null);
    IDocumentBuilder DeleteParagraph(string targetText);
    IDocumentBuilder ApplyStyle(string styleId);
    
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
    void MergeBatch(List<Dictionary<string, string>> records, string outputPattern, string? templatePath = null);
    
    void Save(string? path = null);
}
```

### Content Blocks

Build structured content programmatically:

```csharp
var blocks = new ContentBlockBuilder()
    .AddHeading(1, "Document Title", "CustomTitle")
    .AddParagraph("Introduction paragraph")
    .AddList(false, new[] { "Bullet 1", "Bullet 2" })
    .AddTable(new[] {
        new[] { "Header 1", "Header 2" },
        new[] { "Cell 1", "Cell 2" }
    })
    .AddCode("var x = 1;", "csharp")
    .AddBlockquote("Important quote")
    .Build();

builder.AddRichContent(blocks);
```

### Markdown Conversion

Convert markdown with custom style mapping:

```markdown
# Title (Heading 1)
## Subtitle (Heading 2)

Normal paragraph with **bold** and *italic* text.

- Bullet item
- Another item

1. Numbered item
2. Another item

> Blockquote

| Column 1 | Column 2 |
|----------|----------|
| Cell 1   | Cell 2   |

:::tip
This maps to "Tip" style
:::

:::warning
This maps to "Warning" style
:::
```

```csharp
var styleMap = new StyleMapping
{
    StyleMap = new Dictionary<string, string>
    {
        ["heading1"] = "CustomTitle",
        ["heading2"] = "CustomSubtitle",
        ["paragraph"] = "Normal",
        ["tip"] = "Tip",
        ["warning"] = "Warning"
    }
};

builder.AddMarkdown(markdown, styleMap);
```

### Variable Detection

Detect variables in templates:

```csharp
// Template contains: "Hello {{name}} and {{company|Unknown}}"
var variables = builder.DetectVariables();
// Returns:
// [
//   { Name: "name", FullMatch: "{{name}}", Location: "body" },
//   { Name: "company", FullMatch: "{{company|Unknown}}", Location: "body", DefaultValue: "Unknown" }
// ]
```

### Publipostage (Mail Merge)

Replace variables with data:

```csharp
// Single document
builder.MergeVariables(new Dictionary<string, string>
{
    ["name"] = "John Doe",
    ["company"] = "Acme Corp"
});

// Batch generation
var records = new List<Dictionary<string, string>>
{
    new() { ["name"] = "Alice", ["company"] = "Corp A" },
    new() { ["name"] = "Bob", ["company"] = "Corp B" }
};

builder.MergeBatch(records, "output_{name}.docx", "template.docx");
// Creates: output_Alice.docx, output_Bob.docx
```

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

var engine = new TemplateEngine();
engine.Process(document, data);
```

### Instruction Sets

Execute batch operations via JSON/YAML:

```json
{
  "operations": [
    { "type": "addParagraph", "text": "Hello World", "style": "Heading1" },
    { "type": "replaceText", "find": "{{NAME}}", "replace": "John" },
    { "type": "insertAfter", "target": "Introduction", "content": { "text": "New section" } }
  ]
}
```

```csharp
var parser = new JsonInstructionParser();
var instructions = parser.Parse(json);

var engine = new InstructionEngine();
engine.Execute(builder, instructions);
```

## CLI Commands

| Command | Description |
|---------|-------------|
| `create` | Create new document |
| `edit` | Edit with instructions |
| `template` | Create from template |
| `validate` | Validate instruction file |
| `detect` | List variables in template |
| `merge` | Merge template with data |
| `markdown` | Convert markdown to docx |

```bash
# Examples
docxeditor create output.docx --text "Hello" --style Heading1
docxeditor edit doc.docx --instructions ops.json
docxeditor template tpl.docx out.docx --instructions ops.yaml
docxeditor detect template.docx
docxeditor merge template.docx data.json "output_{client}.docx"
docxeditor markdown input.md output.docx --style-map styles.json
```

## Architecture

### Instruction Pattern
All document operations are modeled as immutable instruction objects executed by the `InstructionEngine`.

### Builder Pattern
Fluent API for composing operations with method chaining.

### Style Preservation
When editing existing documents:
- Loads and caches all existing styles
- Never mutates original style definitions
- References styles by ID when adding content
- Preserves document defaults

## Dependencies

- **.NET 9**
- **DocumentFormat.OpenXml** - Microsoft OpenXML SDK
- **Markdig** - Markdown parser
- **YamlDotNet** - YAML parser
- **Spectre.Console** - CLI output (optional)

## Testing

```bash
dotnet test
```

25+ unit tests covering:
- Document creation and manipulation
- Content block rendering
- Markdown conversion
- Variable detection and replacement
- Serialization

## License

MIT

## Contributing

1. Fork the repository
2. Create a feature branch (`git checkout -b feature/amazing-feature`)
3. Commit changes (`git commit -m 'feat: add amazing feature'`)
4. Push to branch (`git push origin feature/amazing-feature`)
5. Open a Pull Request

## Author

**Maxime** - maxime.le-besnerais@epitech.eu
