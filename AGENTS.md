# DocxEditor - Agent Guidelines

## Project Overview
.NET 9 document editor for creating and editing DOCX files via instruction sets (JSON/YAML) or fluent C# API.

## Development Rules

### Technology Stack
- .NET 9 (9.0.313)
- DocumentFormat.OpenXml for DOCX manipulation
- System.Text.Json for JSON parsing
- YamlDotNet for YAML parsing
- Spectre.Console for CLI output

### Code Conventions
- Use C# 12 features (primary constructors, collection expressions, etc.)
- Follow standard .NET naming conventions (PascalCase for public APIs)
- Use nullable reference types enabled project-wide
- Prefer immutable data structures where possible
- Use `var` only when type is obvious from right-hand side

### Architecture Patterns
- **Instruction Pattern**: All document operations are modeled as instruction objects
- **Builder Pattern**: Fluent API for composing operations
- **Strategy Pattern**: Different instruction executors for different operation types
- **Repository Pattern**: Abstract document storage (file system, stream, etc.)

### Commit Strategy
- Commit after each completed phase or significant milestone
- Use conventional commits: `feat:`, `fix:`, `docs:`, `test:`, `refactor:`, `chore:`
- Never commit broken code
- Write meaningful commit messages explaining WHY, not WHAT

### Testing Requirements
- Write tests for all public APIs
- Use xUnit for unit tests
- Use snapshot testing for XML output comparison
- Minimum 80% code coverage for Core library

### Style Preservation Rules
- Never mutate existing style definitions in template documents
- Always reference styles by ID when adding content
- Cache loaded styles to avoid repeated parsing
- Preserve document defaults (fonts, spacing, etc.)

### Error Handling
- Use custom exception types for domain errors
- Validate all inputs before processing
- Provide meaningful error messages with context
- Never swallow exceptions silently

### Performance Considerations
- Use streaming for large documents
- Minimize OpenXML part lookups
- Cache frequently accessed document parts
- Dispose of WordprocessingDocument instances properly

## Project Structure
```
DocxEditor/
├── DocxEditor.Core/           # Core library
│   ├── Models/                # Data models
│   ├── Instructions/          # Instruction definitions
│   ├── Builders/              # Fluent API builders
│   ├── Serialization/         # JSON/YAML parsers
│   ├── DocumentEngine/        # OpenXML operations
│   └── Exceptions/            # Custom exceptions
├── DocxEditor.Cli/            # CLI application
│   └── Commands/              # CLI commands
└── DocxEditor.Tests/          # Test project
    ├── Unit/                  # Unit tests
    └── Integration/           # Integration tests
```

## Key Decisions
- Use OpenXML SDK instead of manual XML manipulation
- Support both imperative (fluent API) and declarative (JSON/YAML) interfaces
- Preserve all existing styles when editing documents
- Support template-based document creation

## Expanded Project Scope
The project has grown beyond DOCX to include:
- **PPTX → Typst → PDF conversion pipeline** (primary active work)
- **XLSX support** (basic read/write)
- **Typst integration** via TypstBridge-first `TypstCompilerService`, with `typstsharp`/Typst CLI fallback paths retained
- **Reference files in `examples/REF/`** for visual regression testing

## PPTX→Typst Conversion Critical Rules
- **Use regex on `OuterXml` for unreliable OOXML attributes** (`marL`, `indent`, `algn`, `val`, `char`, `type`). `OpenXmlElement.GetAttribute()` crashes on missing attributes.
- **`spcPct` values are 1/1000ths of a percent**: divide by `100000.0` (e.g., `120000` = `120%` = `1.2`).
- **Typst `par(leading:)` is additive** to Typst's default line advance, not a direct replacement for PPTX line spacing.
- **List items must be separate arguments**: `#enum[item1][item2]`, never `#enum[all text]`.
- **Never apply global auto-fit** — it breaks body paragraph wrapping.

## Font Handling Specifics
- **Carlito** is the Linux fallback for missing `Aptos`/`Calibri`
- **Never override embedded PPTX fonts** with `--font-path`; combine paths via `Path.PathSeparator`
- Extract embedded fonts from `ppt/fonts/` and pass via `--font-path`

## User Workflow Constraints
- Incremental, targeted fixes only — never broad visual-fidelity sweeps
- Always verify against reference PDFs in `examples/REF/` before considering a fix complete
- Never commit unless explicitly asked
- Run `dotnet test` after changes
- Run conversions sequentially (parallel `dotnet run` causes build file locks)

## Common Pitfalls
- Legacy TypstSharp can fail on hardened Linux; TypstBridge is the primary backend and Typst CLI remains the safety-net fallback
- `PlaceholderValues` and `SchemeColorValues` parsing via SDK is unreliable → regex fallback
- Table styles: start with header bold/white, avoid complex partial per-cell strokes initially

## Reference Files & Testing
- Reference PPTX/PDF pairs: ``, `` in `examples/REF/`
- Generated outputs: `examples/output/ref/`
- PNG mode: `--format png` generates per-slide images
- Smoke tests must pass on `` and `.pptx`

@RTK.md
