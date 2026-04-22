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
