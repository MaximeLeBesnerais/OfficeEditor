# OfficeEditor

A .NET 9 suite for **creating, editing, generating, and rendering Office documents** — DOCX, PPTX, XLSX — without Office or LibreOffice. Fluent C# builders, JSON/YAML instruction sets, and a **declarative JSON vocabulary that generates complete decks and documents**. Rendering is handled by **TypstBridge**, a native Rust bridge around the Typst engine: PPTX/DOCX → PDF, per-slide PNG, and SVG, in milliseconds per slide.

## Features

- **Three formats, one model** — Word (DOCX), PowerPoint (PPTX), Excel (XLSX); create from scratch or edit existing files with style preservation
- **JSON → document generation** — declarative, loudly-validated vocabulary: cover, content, tables, KPIs, themes (PPTX today; DOCX and XLSX vocabularies on the roadmap — see `docs/roadmap-*.md`)
- **Rendering** — native TypstBridge (Typst 0.15.1): PDF, PNG at configurable PPI, SVG; whole-deck renders in well under a second; per-slide timings exposed everywhere
- **Fluent C# APIs** — `DocumentBuilder`, `PresentationBuilder`, `WorkbookBuilder` (file, stream, or in-memory `byte[]`)
- **Instruction sets** — JSON/YAML edit operations against existing documents
- **Variables & mail merge** — `{{variable}}` detection, merge, batch; `{{#if}}/{{#each}}` template logic
- **Markdown → DOCX** — styled conversion via Markdig
- **Surfaces** — unified CLI, ASP.NET Core API, MCP stdio host (4 `deck_*` tools), and a web demo app
- **Brand profiles** — extract theme colors/fonts from existing decks into reusable token sets

## Quick Start

### CLI

```bash
# Run from source (or install the tool — see Installation)
dotnet run --project OfficeEditor.Cli -- <command>

# Create documents (format auto-detected from extension)
officeeditor create output.docx --text "Hello World"
officeeditor create output.pptx --title "My Presentation"
officeeditor create output.xlsx --sheet "Sales"

# Generate a full deck from a JSON vocabulary
officeeditor generate demo/demo-deck.json --output deck.pptx

# Detect variables in templates
officeeditor detect template.docx

# Merge template with data
officeeditor merge template.pptx data.json output.pptx

# Edit with instructions
officeeditor edit document.docx --instructions instructions.json
```

### C# API — generate a deck from JSON (the flagship path)

```csharp
using PptxEditor.Core.Generation.Archetypes;
using PptxEditor.Core.Generation.Components;
using PptxEditor.Core.Generation.Emit.Ooxml;
using PptxEditor.Core.Generation.Layout;
using PptxEditor.Core.Generation.Schema;

var json = await File.ReadAllTextAsync("deck.json");   // "version": "2.0" vocabulary
var result = new GenerationDocumentParser().Validate(json);   // loud validator
if (!result.IsValid) { /* field-path errors with suggestions */ }

var doc        = ArchetypeExpander.Expand(result.Document!);
var components = ComponentExpander.Expand(doc);
var layout     = new LayoutResolver().Resolve(components);    // layout once…
var pptx       = new OoxmlEmitter().Emit(layout);             // …emit OOXML…
await File.WriteAllBytesAsync("deck.pptx", pptx.Bytes);
// …and the same layout feeds the Typst emitter for PDF/PNG/SVG previews.
```

### C# API — builders (all three formats)

```csharp
using DocxEditor.Core.Builders;
using PptxEditor.Core.Builders;
using XlsxEditor.Core.Builders;

// DOCX
using var doc = DocumentBuilder.Create("report.docx");
doc.AddParagraph("Annual Report", "Heading1");
doc.AddMarkdown("This is **bold** and *italic*, with a list:\n- one\n- two");
doc.Save();

// PPTX — build, then render
using var deck = PresentationBuilder.Create("slides.pptx");
deck.AddSlide();
deck.CurrentSlide.AddTitle("Q4 Review").AddSubtitle("Sales");
deck.AddSlide();
deck.CurrentSlide.AddTitle("Numbers").AddTable(new List<List<string>>
{
    new() { "Product", "Q1", "Q2" },
    new() { "Widget", "100", "200" }
});
deck.Save();

byte[] pdf = deck.ExportToPdf();                                   // whole-deck PDF
byte[][] pngs = deck.ExportThumbnails(new ThumbnailOptions { Ppi = 150 });
string typ = deck.ExportToTypst();                                 // Typst source

// XLSX
using var book = WorkbookBuilder.Create("data.xlsx");
var sheet = book.AddWorksheet("Sales");
sheet.AddHeaderRow(new List<string> { "Product", "Q1", "Q2" })
     .AddDataRow(new List<string> { "Widget", "100", "200" }, 2)
     .AddFormulaRow(new List<string> { "Total", "=SUM(B2:B2)", "=SUM(C2:C2)" }, 3);
book.Save();

// In-memory (services, Azure Functions, APIs)
using var mem = DocumentBuilder.Create();
mem.AddParagraph("Hello");
BinaryOfficeDocument bin = mem.ToBinaryDocument();   // bin.Bytes, bin.ContentType
```

### Variables, mail merge, template logic

```csharp
var variables = builder.DetectVariables();          // find {{vars}} — all formats
builder.MergeVariables(new Dictionary<string, string>
{
    ["clientName"] = "Acme Corp",
    ["date"] = "2026-01-31"
});

// Template logic (DOCX/PPTX/XLSX template engines)
// {{#if amount > 100}}…{{/if}}   {{#each items}}…{{/each}}
```

## Installation

NuGet packages are published at **0.1.0** (currently *unlisted* — install by exact version; they don't appear in search yet):

```bash
dotnet add package MaximeLB.PptxEditor.Core --version 0.1.0   # PPTX
dotnet add package MaximeLB.DocxEditor.Core --version 0.1.0   # DOCX
dotnet add package MaximeLB.XlsxEditor.Core --version 0.1.0   # XLSX
dotnet tool install -g MaximeLB.OfficeEditor.Cli --version 0.1.0   # `officeeditor` command
```

| Package | Contents |
|---|---|
| `MaximeLB.OfficeEditor.Core` | Shared services (TypstCompilerService), variables, models |
| `MaximeLB.DocxEditor.Core` | DOCX builder, instructions, markdown, DOCX→Typst converter |
| `MaximeLB.PptxEditor.Core` | PPTX builder, converters, generation pipeline, brand profiles |
| `MaximeLB.XlsxEditor.Core` | XLSX builder, variables |
| `MaximeLB.TypstBridge.Managed` | Managed wrapper + native TypstBridge (osx-arm64, linux-x64, win-x64) |
| `MaximeLB.OfficeEditor.Cli` | Unified `officeeditor` CLI (dotnet tool) |

Or build from source:

```bash
git clone https://github.com/MaximeLeBesnerais/OfficeEditor.git
cd OfficeEditor && dotnet build        # 0 warnings, 0 errors (enforced)
```

## Repository map — where things live

| What | Where |
|---|---|
| **Reference corpus** (license-clean fixtures + PowerPoint/Word-rendered ground-truth PDFs) | `examples/REF/` — PPTX: `sales_acceleration_deck` (primary, 16 slides, SmartArt), `AetherLink-Glass-Shareholder-Overview` (styled, brand mining), `northwind-investor-40` (stress), `northwind-launch-review` (pitch), `northwind-demo` (self-generated smoke deck). DOCX: `annual-report`, `monitoring-report` (`{{variable}}` templates) |
| **Generation JSON examples** | `demo/demo-deck.json` (15 slides + `demo/themes.json` presets), `demo/deck.json` (6 slides), `decks/repo-overview.json` |
| **Roadmaps** | `docs/roadmap-pptx.md`, `docs/roadmap-docx.md`, `docs/roadmap-xlsx.md` — everything targets v0.5 |
| **Design token sets** (mined brand profiles) | `PptxEditor.Core/Generation/Design/` |
| **Parity fixtures** (per-primitive generation tests) | `PptxEditor.Core/Generation/Fixtures/` |
| **Tools** | `tools/convert-pptx`, `tools/convert-docx`, `tools/visual-diff`, `tools/pptx-benchmark` |
| **Typst upgrade notes** | `docs/typst-0.15.md` |
| **Agent/ contributor rules** | `AGENTS.md`, `agent-instructions/` |
| Local-only dev fixtures (gitignored) | `local-ref/` — never commit |

## Entry points

| Component | Run it |
|---|---|
| Unified CLI | `dotnet run --project OfficeEditor.Cli -- <command>` |
| API | `dotnet run --project OfficeEditor.Api --urls http://localhost:5001` |
| Web demo | `make dev` → http://localhost:5173/ (`/demo` = the app) |
| MCP host (JSON-RPC stdio) | `dotnet run --project OfficeEditor.Mcp` |
| Examples | `dotnet run --project examples` |
| Convert / diff / bench tools | `dotnet run --project tools/convert-pptx -- <in> <out> [--format pdf\|png]` · `tools/convert-docx` · `tools/visual-diff -- --suite pptx\|docx\|gen` · `tools/pptx-benchmark` |
| CLI demo | `make -f Makefile.demo demo` (preflight → convert REF deck → generate deck, prints timings, opens PDFs) |

### The demo

`make dev`, open http://localhost:5173/ — four tabs, all timings server-measured:

1. **Render** — pick a whitelisted REF deck → timed PNG/SVG gallery ("16 slides in 394ms, 24.6ms/slide")
2. **Generate** — edit the demo deck's title live + swap theme presets → PPTX + previews + **downloadable .pptx**
3. **Any render** — upload any `.pptx`
4. **Compare** — OfficeEditor Engine vs headless LibreOffice, side-by-side slides and timings (typically >10× faster)

## Architecture

```
DocxEditor/                          # repo folder (historical name; product is OfficeEditor)
├── DocxEditor.Core/                 # DOCX: Builders, Content, Markdown, Instructions, Converters
├── PptxEditor.Core/                 # PPTX: Builders, Converters, Variables
│   └── Generation/                  #   JSON vocab: Schema (loud validator) → Archetypes → Components
│       #                            #   → Layout (once, pure C#) → Emit/OOXML + Emit/Typst (twice)
├── XlsxEditor.Core/                 # XLSX: Builders, Variables
├── OfficeEditor.Core/               # Shared: TypstCompilerService, Variables, Exceptions
├── OfficeEditor.Cli/                # Unified multi-format CLI
├── OfficeEditor.Api/                # ASP.NET Core: deck sessions, previews, generation, demo/compare
├── OfficeEditor.Mcp/                # MCP stdio host: deck_anatomize, deck_replace_element, …
├── OfficeEditor.Web.Client/         # React/Vite/Tailwind demo app (/demo)
├── TypstBridge/                     # Rust native bridge + managed wrapper (Typst 0.15.1)
├── examples/                        # Sample programs + REF corpus
├── demo/, decks/                    # Generation JSON decks
└── tools/                           # convert-pptx, convert-docx, visual-diff, pptx-benchmark
```

Key design decisions:

- **Layout once, emit twice** — one C# layout pass feeds the OOXML emitter (delivery) and the Typst emitter (preview); no second layout engine
- **Every primitive ships with both emitters + a parity fixture** — no half-tested features
- **Typst preview is the spec of record** for ambiguous OOXML rendering
- **Style preservation** — editing never mutates existing style definitions; styles referenced by ID

## API & MCP surfaces

**API** (`OfficeEditor.Api`): deck upload/sessions, per-slide previews (`png|svg`, ETag-cached), deck anatomy, edit instructions, JSON generation with timings (`generationMilliseconds`, `totalMilliseconds`), demo endpoints (timed REF renders, upload render, OfficeEditor-vs-LibreOffice compare), `/api/convert` for one-off conversions.

**MCP** (`OfficeEditor.Mcp`, stdio JSON-RPC): `deck_anatomize`, `deck_replace_element`, `deck_render_slide`, `deck_generate`.

## Testing

```bash
dotnet test                 # full suite: ~1,250 tests across 4 projects, 0 failures
```

- xUnit; per-primitive parity fixtures with RMSE thresholds for the generation pipeline
- Merged line coverage across suites: **82.3%** (measured with coverlet; union of all test runs)
- Typst-dependent tests are env-gated: `OE_RUN_TYPST_COMPILE_TESTS=1 dotnet test`
- CI: `build-test` on PRs and `main` pushes — Release build (warnings = errors) + full suite + coverage gate (41% floor); runs only when C#-relevant paths change
- Visual regression: `dotnet run --project tools/visual-diff -- --suite pptx|docx|gen` (baselines are per-machine, not committed)

## Performance

PPTX render pipeline (PptxEditor → Typst → PNG/PDF) vs headless LibreOffice
(`tools/pptx-benchmark`, median of 5 warm runs, Apple Silicon — treat as orders of magnitude):

| Deck | Slides | OfficeEditor warm (total) | Per slide | LibreOffice warm (total) | Per slide | Speedup |
|---|---|---|---|---|---|---|
| sales_acceleration_deck | 16 | 329.4 ms | 20.6 ms | 5,924.0 ms | 370.2 ms | ~18.0× |
| AetherLink-Glass-Shareholder-Overview | 15 | 417.4 ms | 27.8 ms | 15,940.3 ms | 1,062.7 ms | ~38.2× |
| northwind-launch-review | 12 | 102.5 ms | 8.5 ms | 2,034.2 ms | 169.5 ms | ~19.8× |
| northwind-investor-40 | 40 | 322.2 ms | 8.1 ms | 5,629.7 ms | 140.7 ms | ~17.5× |

Same artifacts both sides: OfficeEditor renders slide PNGs natively in one compile (open + whole-deck PNG @150ppi); LibreOffice cannot rasterize PPTX, so its total is `soffice --convert-to pdf` **+ `pdftoppm` rasterization at 150dpi** — and rasterization is the dominant cost. Cold starts (fresh process, JIT + backend probes): ~310–830 ms per deck, deck-dependent.
Reproduce: `dotnet run --project tools/pptx-benchmark` (methodology in `tools/pptx-benchmark/README.md`).

## Typst compilation backend

`TypstCompilerService` compiles through **TypstBridge** (in-process native bridge, Typst 0.15.1) with the external `typst` CLI as a safety net. Supports PDF/SVG/PNG, multi-page output, working-directory assets, explicit font paths (`--font-path`), PNG PPI, persistent compile sessions, and diagnostics.

## Status & limitations

**Version:** 0.1.0 (pre-1.0). Public API may change; the TypstBridge ABI is intentionally fluid until v1.0.

**Platforms:** pure managed .NET 9 + per-RID native TypstBridge (osx-arm64, linux-x64, win-x64 built in CI).
- **macOS (arm64)** — development platform; everything verified here
- **linux-x64** — tested
- **Windows** — expected to work; runtime verification in CI is on the PPTX roadmap

**Rendering (PPTX → Typst → PDF/PNG/SVG):**
- Text, images, shapes, and tables render with good fidelity
- Charts, SmartArt, and animations have **partial support** — complex instances may render simplified or be omitted
- Font fidelity varies by platform; installing Microsoft Office fonts (Aptos, Calibri) improves accuracy, but exact PowerPoint parity is not guaranteed (font metrics, line breaking, layout engines differ)

**NuGet:** packages published at 0.1.0 but currently *unlisted* — install by exact version.

## License

MIT — Copyright 2026 Maxime Le Besnerais

## Contributing

The repo uses a protected `main` / working `dev` branch model:

1. Branch from `dev`; CI runs on your PR (build + full tests + coverage gate; C#-path-gated)
2. PRs into `main` must come from `dev` (`guard-main` enforces it)
3. Releases: tag `v*` on `main` → native matrix → NuGet trusted publishing

See `AGENTS.md` for engineering conventions and `docs/roadmap-*.md` for what's planned.

## Author

**Maxime Le Besnerais**
