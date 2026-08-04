# OfficeEditor

A .NET 9 library suite for **creating, editing, generating, and rendering Office documents** — DOCX, PPTX, XLSX — without requiring Office or LibreOffice. It provides fluent C# builders, JSON/YAML instruction sets, and a loudly validated declarative JSON vocabulary for complete PPTX decks. Rendering uses **TypstBridge**, a native Rust bridge around Typst: PPTX → PDF/PNG/SVG, DOCX → PDF, and XLSX → PDF/PNG/SVG.

> **Now public** — OfficeEditor 0.7.0 is open source under the MIT license.

## Features

- **Three formats, one model** — Word (DOCX), PowerPoint (PPTX), Excel (XLSX); create from scratch or edit existing files with style preservation
- **JSON workflows** — PPTX has the full declarative generation vocabulary; **DOCX has declarative JSON generation** (flow + positioned tiers, design themes, semantic report archetypes, see `docs/docx-generation.md`); **XLSX has a rich instruction/generation engine** (typed cells, named styles with fills/borders, layout, tables) wired into `officeeditor generate --output *.xlsx` and renderable to PDF/PNG/SVG through the Typst pipeline (images and row-replication remain)
- **Rendering** — native TypstBridge (Typst 0.15.1): PPTX PDF/PNG/SVG, DOCX PDF, and XLSX PDF/PNG/SVG (formulas render cached `<v>` values only, no evaluation); whole-deck timings are exposed by the PPTX surfaces
- **Fluent C# APIs** — `DocumentBuilder`, `PresentationBuilder`, `WorkbookBuilder` (file, stream, or in-memory `byte[]`)
- **Instruction sets** — JSON/YAML DOCX operations, JSON PPTX edit operations, and a v1 JSON XLSX builder vocabulary
- **Variables & mail merge** — `{{variable}}` detection and replacement across all three formats, plus DOCX batch merge
- **Markdown → DOCX** — rich styled conversion via Markdig (headings 1–6, nested emphasis, real hyperlinks through a safe `http`/`https`/`mailto` scheme allowlist with internal-anchor fallback, images, footnotes, tables, task lists, emoji, YAML front matter, custom style maps)
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

# Generate a DOCX report from a JSON vocabulary (design theme optional)
officeeditor generate report.json --output report.docx --theme corporate

# Generate a workbook from a JSON instruction set
officeeditor generate workbook.json --output workbook.xlsx

# Render an XLSX JSON instruction set to PDF (single file) or PNG (a directory of page-NNN.png pages)
officeeditor generate workbook.json --output workbook.pdf
officeeditor generate workbook.json --output workbook.png

# Detect variables in templates
officeeditor detect template.docx

# Merge template with data
officeeditor merge template.pptx data.json output.pptx

# DOCX-only Markdown → DOCX with a template, custom style map, and strict mode
dotnet run --project DocxEditor.Cli -- markdown guide.md guide.docx --template base.docx --style-map styles.json --strict

# DOCX-only instruction editing is available through the source CLI
dotnet run --project DocxEditor.Cli -- edit document.docx --instructions instructions.json
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
sheet.SetColumnWidth("A", 24).SetRowHeight(1, 28);
sheet.MergeCells("A4:C4");
book.Save();

// In-memory (services, Azure Functions, APIs)
using var mem = DocumentBuilder.Create();
mem.AddParagraph("Hello");
BinaryOfficeDocument bin = mem.ToBinaryDocument();   // bin.Bytes, bin.ContentType
```

### Variables and mail merge

```csharp
var variables = builder.DetectVariables();          // find {{vars}} — all formats
builder.MergeVariables(new Dictionary<string, string>
{
    ["clientName"] = "Acme Corp",
    ["date"] = "2026-01-31"
});

// Lower-level template-engine classes also exist for conditional/loop expansion,
// but they are not wired into the unified CLI or public demo surfaces.
```

## Installation

NuGet packages are listed at **0.7.0**:

```bash
dotnet add package MaximeLB.PptxEditor.Core --version 0.7.0   # PPTX
dotnet add package MaximeLB.DocxEditor.Core --version 0.7.0   # DOCX
dotnet add package MaximeLB.XlsxEditor.Core --version 0.7.0   # XLSX
dotnet tool install -g MaximeLB.OfficeEditor.Cli --version 0.7.0   # `officeeditor` command
```

| Package | Contents |
|---|---|
| `MaximeLB.OfficeEditor.Core` | Shared services (TypstCompilerService), variables, models |
| `MaximeLB.DocxEditor.Core` | DOCX builder, instructions, markdown, DOCX→Typst converter |
| `MaximeLB.PptxEditor.Core` | PPTX builder, converters, generation pipeline, brand profiles |
| `MaximeLB.XlsxEditor.Core` | XLSX builder/read API, variables, JSON instructions |
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
| **Design token sets** (mined brand profiles) | `PptxEditor.Core/Generation/Design/` |
| **Parity fixtures** (per-primitive generation tests) | `PptxEditor.Core/Generation/Fixtures/` |
| **Tools** | `tools/convert-pptx`, `tools/convert-docx`, `tools/convert-xlsx`, `tools/visual-diff`, `tools/pptx-benchmark` |
| **Agent/ contributor rules** | `AGENTS.md` |

## Entry points

| Component | Run it |
|---|---|
| Unified CLI | `dotnet run --project OfficeEditor.Cli -- <command>` |
| API | `dotnet run --project OfficeEditor.Api --urls http://localhost:5001` |
| Web demo | `make dev` → http://localhost:5173/ (`/demo` = the app) |
| MCP host (JSON-RPC stdio) | `dotnet run --project OfficeEditor.Mcp` |
| Examples | `dotnet run --project examples` |
| Convert / diff / bench tools | `dotnet run --project tools/convert-pptx -- <in> <out> [--format pdf\|png\|svg\|typ]` · `dotnet run --project tools/convert-docx -- <in> <out> [--format pdf\|png\|svg\|typ]` · `dotnet run --project tools/convert-xlsx -- <in> <out> [--format pdf\|png\|svg\|typ\|json]` · `dotnet run --project tools/visual-diff -- --suite pptx\|gen\|xlsx` · `dotnet run --project tools/pptx-benchmark` |
| CLI demo | `make -f Makefile.demo demo` (preflight → convert REF deck → generate deck, prints timings, opens PDFs) |

### The demo

`make dev`, open http://localhost:5173/ — four tabs, all timings server-measured:

1. **Render** — pick a whitelisted REF deck → timed PNG/SVG gallery (the current benchmark is about 329 ms / 20.6 ms per slide for the 16-slide sales deck on Apple Silicon)
2. **Generate** — edit the demo deck's title live + swap theme presets → PPTX + previews + **downloadable .pptx**
3. **Any render** — upload any `.pptx`
4. **Compare** — OfficeEditor Engine vs headless LibreOffice, side-by-side slides and timings (typically >10× faster)

The API and web client are local demos, not production multi-tenant services. They have no complete authentication, quota, sandbox, or tenant-isolation layer. See [SECURITY.md](SECURITY.md).

## Architecture

```
DocxEditor/                          # repo folder (historical name; product is OfficeEditor)
├── DocxEditor.Core/                 # DOCX: Builders, Content, Markdown, Instructions, Converters
├── PptxEditor.Core/                 # PPTX: Builders, Converters, Variables
│   └── Generation/                  #   JSON vocab: Schema (loud validator) → Archetypes → Components
│       #                            #   → Layout (once, pure C#) → Emit/OOXML + Emit/Typst (twice)
├── XlsxEditor.Core/                 # XLSX: Builders/read API, Variables, JSON Instructions
├── OfficeEditor.Core/               # Shared: TypstCompilerService, Variables, Exceptions
├── OfficeEditor.Cli/                # Unified multi-format CLI
├── OfficeEditor.Api/                # ASP.NET Core: deck sessions, previews, generation, demo/compare
├── OfficeEditor.Mcp/                # MCP stdio host: deck_anatomize, deck_replace_element, …
├── OfficeEditor.Web.Client/         # React/Vite/Tailwind demo app (/demo)
├── TypstBridge/                     # Rust native bridge + managed wrapper (Typst 0.15.1)
├── examples/                        # Sample programs + REF corpus
├── demo/, decks/                    # Generation JSON decks
└── tools/                           # convert-pptx, convert-docx, convert-xlsx, visual-diff, pptx-benchmark
```

Key design decisions:

- **Layout once, emit twice** — one C# layout pass feeds the OOXML emitter (delivery) and the Typst emitter (preview); no second layout engine
- **Every primitive ships with both emitters + a parity fixture** — no half-tested features
- **Typst preview is the spec of record** for ambiguous OOXML rendering
- **Style preservation** — editing never mutates existing style definitions; styles referenced by ID

## API & MCP surfaces

**API** (`OfficeEditor.Api`): deck upload/sessions, per-slide previews (`png|svg`, ETag-cached), deck anatomy, edit instructions, JSON generation with timings (`generationMilliseconds`, `totalMilliseconds`), demo endpoints (timed REF renders, upload render, OfficeEditor-vs-LibreOffice compare), and `/api/convert` for one-off conversions — JSON → DOCX/XLSX through the declarative generators (empty JSON makes a blank document), Markdown → DOCX, DOCX → PDF, PPTX → PDF/PNG/SVG, and XLSX → PDF/PNG/SVG (PNG/SVG return the first page).

**MCP** (`OfficeEditor.Mcp`, stdio JSON-RPC): `deck_anatomize`, `deck_replace_element`, `deck_render_slide`, `deck_generate`.

## Testing

```bash
dotnet test                 # full suite: 2,300+ cases across 5 test projects
```

- xUnit; per-primitive parity fixtures with RMSE thresholds for the generation pipeline
- Coverage: merged-union line coverage across all test projects (`scripts/check-coverage.py`); CI gate floor **83%** — source of truth is `COVERAGE_THRESHOLD` in `.github/workflows/ci.yml`
- Typst-dependent tests are env-gated: `OE_RUN_TYPST_COMPILE_TESTS=1 dotnet test`
- CI: `build-test` on PRs and `main` pushes — Release build (warnings = errors) + full suite + coverage gate (83% floor); runs only when C#-relevant paths change
- Visual regression: PPTX and generation suites are runnable today. The DOCX suite is wired but still needs generated reference outputs under `examples/output/ref/docx/`.

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

**Version:** 0.7.0 (pre-1.0). Public API may change; the TypstBridge ABI is intentionally fluid until v1.0.

**Platforms:** pure managed .NET 9 + per-RID native TypstBridge (osx-arm64, linux-x64, win-x64 built in CI).
- **macOS (arm64)** — development platform; everything verified here
- **linux-x64** — tested
- **Windows** — expected to work; runtime verification in CI is still pending

**Rendering (PPTX → Typst → PDF/PNG/SVG):**
- Text, images, shapes, and tables render with good fidelity
- Clustered bar/column charts render; other chart families can fall back to placeholders. SmartArt renders from pre-rendered drawing shapes with documented theme-color limits. Animations are outside the static preview model.
- Font fidelity varies by platform; installing Microsoft Office fonts (Aptos, Calibri) improves accuracy, but exact PowerPoint parity is not guaranteed (font metrics, line breaking, layout engines differ)

**NuGet:** packages are listed at 0.7.0.

## Security

OfficeEditor libraries run in the caller's process and are not a sandbox for hostile documents. Deployments accepting untrusted uploads must provide their own isolation, resource limits, authentication, and filesystem policy. The API/web demo and LibreOffice comparison path are local conveniences, not the product security boundary. See [SECURITY.md](SECURITY.md) for reporting instructions and the full trust model.

## License

MIT — Copyright 2026 Maxime Le Besnerais

This product bundles third-party components under their own licenses; see [THIRD-PARTY-NOTICES.md](THIRD-PARTY-NOTICES.md) (notably Typst, Apache-2.0).

## Contributing

The repo uses a protected `main` / working `dev` branch model:

1. Branch from `dev`; CI runs on your PR (build + full tests + coverage gate; C#-path-gated)
2. PRs into `main` must come from `dev` (`guard-main` enforces it)
3. Release policy: create `v*` tags from tested `main` commits → native matrix → NuGet trusted publishing

See `AGENTS.md` for engineering conventions and the release process.

## Author

**Maxime Le Besnerais**
