# DocxEditor - Development Plan

## Overview
A .NET 9 document editor that can produce DOCX files from scratch using instruction sets, or edit existing documents to add/change content while preserving existing styles and structure.

## Architecture

```
DocxEditor/
├── DocxEditor.Core/           # Core library (class lib)
│   ├── Models/                # Instruction models, DTOs
│   ├── Instructions/          # Instruction execution engine
│   ├── Builders/              # Fluent API builders
│   ├── Serialization/         # JSON/YAML instruction parsers
│   └── DocumentEngine/        # OpenXML wrapper & operations
├── DocxEditor.Cli/            # CLI tool
│   └── Commands/              # CLI commands (create, edit, etc.)
└── DocxEditor.Tests/          # Unit/integration tests
```

## Technology Stack
- **.NET 9** (9.0.313)
- **DocumentFormat.OpenXml** - Microsoft's official OpenXML SDK
- **System.Text.Json** - JSON instruction parsing
- **YamlDotNet** - YAML instruction parsing
- **Spectre.Console** - Rich CLI output

## Commit Strategy

### Branch Strategy
- `main` - Production-ready code
- `feature/*` - Feature branches
- `fix/*` - Bug fix branches

### Commit Conventions
- `feat: ` - New features
- `fix: ` - Bug fixes
- `docs: ` - Documentation changes
- `test: ` - Test additions/changes
- `refactor: ` - Code refactoring
- `chore: ` - Build/tooling changes

### Commit Checkpoints
1. `chore: initialize solution structure`
2. `feat: add core instruction models`
3. `feat: implement document builder fluent API`
4. `feat: add JSON/YAML serialization`
5. `feat: implement core document operations`
6. `feat: add style preservation engine`
7. `feat: implement CLI tool`
8. `test: add unit and integration tests`

## Implementation Phases

### Phase 1: Foundation
- [ ] Create solution and project structure
- [ ] Add DocumentFormat.OpenXml package
- [ ] Add YamlDotNet package
- [ ] Add Spectre.Console package
- [ ] Create basic folder structure

### Phase 2: Core Models
- [ ] Define instruction base classes and interfaces
- [ ] Create operation types (Create, AddParagraph, ReplaceText, etc.)
- [ ] Define document models (Paragraph, Run, Style, etc.)
- [ ] Create DTOs for serialization

### Phase 3: Fluent API
- [ ] Design IDocumentBuilder interface
- [ ] Implement DocumentBuilder class
- [ ] Add paragraph building methods
- [ ] Add text replacement methods
- [ ] Add insertion methods (before/after)
- [ ] Add style application methods

### Phase 4: Instruction Engine
- [ ] Create InstructionEngine class
- [ ] Implement instruction execution pipeline
- [ ] Add error handling and validation
- [ ] Support for batch operations

### Phase 5: Serialization
- [ ] JSON instruction parser
- [ ] YAML instruction parser
- [ ] Validation of instruction files
- [ ] Error reporting for invalid instructions

### Phase 6: Document Operations
- [ ] Create document from scratch
- [ ] Open existing document
- [ ] Add paragraph at end
- [ ] Insert paragraph at position
- [ ] Replace text in paragraph
- [ ] Replace entire paragraph
- [ ] Delete paragraph
- [ ] Apply styles
- [ ] Add tables
- [ ] Add images

### Phase 7: Style Preservation
- [ ] Load and cache existing styles
- [ ] Reference styles by ID
- [ ] Preserve document defaults
- [ ] Allow style overrides at element level
- [ ] Clone styles from template documents

### Phase 8: CLI Tool
- [ ] Create CLI project
- [ ] Add create command
- [ ] Add edit command
- [ ] Add template command
- [ ] Add validate command
- [ ] Rich console output

### Phase 9: Testing
- [ ] Unit tests for builders
- [ ] Unit tests for instruction engine
- [ ] Integration tests for document operations
- [ ] Snapshot tests for generated XML
- [ ] CLI command tests

## Core Operations

| Operation | Description |
|-----------|-------------|
| `Create` | New blank DOCX from scratch |
| `AddParagraph` | Append paragraph at end |
| `InsertParagraph` | Insert at specific position |
| `ReplaceText` | Find/replace text within paragraphs |
| `ReplaceParagraph` | Replace entire paragraph content |
| `DeleteParagraph` | Remove paragraph by content or index |
| `ApplyStyle` | Apply existing or new style |
| `CloneStyle` | Copy style from existing document |
| `AddTable` | Insert table |
| `AddImage` | Insert image |

## Public API Surface

```csharp
// Entry points
IDocumentBuilder DocumentBuilder.Create(string path);
IDocumentBuilder DocumentBuilder.Open(string path);

// Core interfaces
interface IDocumentBuilder {
    IDocumentBuilder AddParagraph(string text, string? style = null);
    IDocumentBuilder InsertAfter(string targetText, IContentBuilder content);
    IDocumentBuilder ReplaceText(string find, string replace);
    IDocumentBuilder ApplyStyle(string styleId);
    void Save(string? path = null);
}

// Instruction engine
class InstructionEngine {
    void Execute(Document document, IEnumerable<Instruction> instructions);
}
```

## Style Preservation Strategy
1. Load and cache all existing styles from `styles.xml`
2. Never mutate original style definitions
3. Reference by style ID when adding content
4. Allow style override at element level (local formatting)
5. Preserve document defaults (fonts, spacing, etc.)

## Next Steps
1. Initialize solution structure
2. Create AGENTS.md for development guidelines
3. Implement core models
4. Build fluent API
5. Add serialization support
