# OfficeEditor - Agent Guidelines

> **Load first, always.** Domain-specific rules live in `agent-instructions/AGENTS.pptx.md` and `agent-instructions/AGENTS.typst.md` — load those only when the task touches PPTX or Typst code. 

## Project Overview

.NET 9 (9.0.313) suite for creating and editing Office documents — **DOCX, PPTX, XLSX** — via instruction sets (JSON/YAML), fluent C# APIs, or a **declarative JSON vocabulary for generating beautiful PPTX from scratch**. Also: edit existing PPTX with smart text/image replacement, extract brand profiles, convert PPTX → Typst → PDF/PNG/SVG, and verify visual fidelity with a per-primitive RMSE parity suite.

The repo folder is named `DocxEditor/` for historical reasons; the product is **OfficeEditor** (see `README.md`).

## Setup Verification (mandatory first step)

Run before any work:

```bash
dotnet --version          # expect 9.0.x
git status                # expect clean or only intended changes
```

## Technology Stack

- .NET 9, C# 12 (primary constructors, collection expressions)
- DocumentFormat.OpenXml — DOCX, PPTX, XLSX
- TypstBridge (primary) → `typstsharp` (legacy PDF fallback) → `typst` CLI (safety net)
- System.Text.Json, YamlDotNet, Spectre.Console
- xUnit for tests

## Code Conventions

- Nullable reference types enabled project-wide
- PascalCase public APIs; `_camelCase` private fields
- Prefer immutable data structures
- Validate all inputs; custom exceptions for domain errors; never swallow exceptions
- Style preservation: never mutate existing style definitions in template documents; reference styles by ID; cache loaded styles; preserve document defaults

## Architecture Patterns

- **Instruction Pattern** — operations modeled as instruction objects (see `DocxEditor.Core/Instructions/`, `PptxEditor.Core/Instructions/`)
- **Generation Pipeline** — JSON → parse (schema + loud validator) → expand (archetypes → components → primitives) → layout (pure C#, once) → emit (OOXML + Typst dual path). See `PptxEditor.Core/Generation/`.
- **Builder Pattern** — fluent API for composing operations (see `*.Builders/`)
- **Strategy Pattern** — different executors per instruction type
- **Repository Pattern** — abstract document storage (file system, stream)

## Workflow Constraints

- **Never commit unless explicitly asked.** Inspect `git status` and `git diff` first.
- **Incremental, targeted fixes only** — never broad visual-fidelity sweeps.
- Run `dotnet test` after every change. Smoke tests must pass on `examples/REF/PPTX/.pptx` and `AetherLink-Glass-Shareholder-Overview.pptx`.
- Run conversions **sequentially** — parallel `dotnet run` causes build file locks.
- Visual fixes: always verify against reference PDFs in `examples/REF/` before claiming done.
- Use conventional commits (`feat:`, `fix:`, `docs:`, `test:`, `refactor:`, `chore:`) — explain **why**, not what.

## Project Structure

```
DocxEditor/                         # repo folder (historical name)
├── DocxEditor.Core/                # DOCX: Models, Instructions, Builders, DocumentEngine
├── DocxEditor.Cli/                 # DOCX CLI
├── DocxEditor.Tests/               # xUnit tests (unit + integration)
├── PptxEditor.Core/                # PPTX: Builders, Converters, Models, Services, Variables
│   └── Generation/                 #   Model, Schema, Layout, Emit/Ooxml, Emit/Typst, Components, Archetypes, Fixtures, Design
├── XlsxEditor.Core/                # XLSX: Builders, Variables
├── OfficeEditor.Core/              # Shared services (TypstCompilerService, Variables, Exceptions)
├── OfficeEditor.Cli/               # Multi-format CLI (create, edit, detect, merge, generate)
├── OfficeEditor.Api/               # ASP.NET Core API: deck sessions, slide previews/thumbnails, instruction/anatomy endpoints
├── OfficeEditor.Api.Tests/         # xUnit tests for the API
├── OfficeEditor.Mcp/               # MCP stdio host (JSON-RPC): deck_anatomize, deck_replace_element, deck_render_slide, deck_generate
├── OfficeEditor.Mcp.Tests/         # xUnit tests for the MCP host
├── OfficeEditor.Web.Client/        # Vite + Tailwind web frontend
├── TypstBridge/                    # Native + managed wrapper around Typst (primary backend)
├── examples/                       # Sample programs + REF/ for visual regression
├── tools/                          # visual-diff suite, pptx-benchmark, convert tools
```

## Domain-Specific Guides

- **PPTX work** → read `AGENTS.pptx.md` (regex-on-OuterXml, spcPct math, list args, font handling, reference files)
- **Typst work** → read `AGENTS.typst.md` (backend selection, output formats, native vs CLI/legacy output differences)

## Key Decisions

- OpenXML SDK over manual XML manipulation
- Imperative (fluent API) + declarative (JSON/YAML) interfaces
- Preserve all existing styles when editing documents
- Template-based document creation
- **Layout once, emit twice** — single C# layout pass shared by OOXML (delivery) and Typst (#place-only preview) emitters
- **Every primitive ships with both emitters + parity fixture** — no half-tested features
- **Typst preview is the spec of record** for ambiguous OOXML rendering — match OOXML to the preview, not vice versa

## Reference Files

- `examples/REF/PPTX/.{pptx,pdf}` and `AetherLink-Glass-Shareholder-Overview.{pptx,pdf}` — PPTX smoke tests
- `examples/REF/PPTX/AetherLink-Glass-Shareholder-Overview.{pptx,pdf}` and `-Architecture-Overview.{pptx,pdf}` — brand/token mining REF decks
- `examples/REF/DOCX/-entreprise-bcp-pme.{docx,pdf}`, `Monitoring Report Template.{docx,pdf}`, `Annual reporting template ENGLISH_0.{docx,pdf}` — DOCX regression
- Generated outputs: `examples/output/ref/`
- PNG mode: `--format png` generates per-slide images at the configured PPI
- Token sets: `PptxEditor.Core/Generation/Design/` — mined from REF decks (, AetherLink)
- Generation fixtures: `PptxEditor.Core/Generation/Fixtures/` — per-primitive parity test decks
- CLI demo decks: `office-editor-full-deck.json` (20 slides, all features), `repo-intro-deck.json` (8 slides)
