# OfficeEditor - Agent Guidelines

> **Load first, always.** Domain-specific rules live in `agent-instructions/AGENTS.pptx.md` and `agent-instructions/AGENTS.typst.md` — load those only when the task touches PPTX or Typst code. 

## Project Overview

.NET 9 (9.0.313) suite for creating and editing Office documents — **DOCX, PPTX** via instruction sets (JSON/YAML), a **declarative JSON vocabulary for generating beautiful PPTX from scratch**, and **XLSX** via fluent C# APIs and a variable templating pipeline. JSON instructions for XLSX are on the Phase 2 roadmap. Also: edit existing PPTX with smart text/image replacement, extract brand profiles, convert PPTX → Typst → PDF/PNG/SVG, and verify visual fidelity with a per-primitive RMSE parity suite.

The repo folder is named `DocxEditor/` for historical reasons; the product is **OfficeEditor** (see `README.md`).

## Setup Verification (mandatory first step)

Run before any work:

```bash
dotnet --version          # expect 9.0.x or newer (projects target net9.0; 10.x SDK works)
git status                # expect clean or only intended changes
```

## Technology Stack

- .NET 9, C# 12 (primary constructors, collection expressions)
- DocumentFormat.OpenXml — DOCX, PPTX, XLSX
- TypstBridge (primary) → `typst` CLI (safety net)
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
- Run `dotnet test` after every change. Smoke tests must pass on `examples/REF/PPTX/northwind-demo.pptx` and the license-clean REF corpus below.
- Run conversions **sequentially** — parallel `dotnet run` causes build file locks.
- Visual fixes: always verify against reference PDFs in `examples/REF/` before claiming done.
- Use conventional commits (`feat:`, `fix:`, `docs:`, `test:`, `refactor:`, `chore:`) — explain **why**, not what.
- **Never delete a stalled agent's worktree/branch without proving it is stalled.** Check the agent/session status AND inspect the worktree for uncommitted work (`git status`, recent file mtimes) immediately before removal — an earlier check is not proof. If work exists, recover it (commit it or restore from the agent's session transcript) before deleting anything.

## Workflow (branches, CI, release)

- **Branches:** `dev` is the default/working branch; `main` is protected. PRs into `main` must come from `dev` (enforced by `guard-main.yml`); both branches require the `build-test` check; owner may self-merge/bypass (warnings expected on direct pushes).
- **CI (`ci.yml`):** runs on PRs and pushes to `main` only — pushes to `dev` trigger nothing. `build-test` = Release build + full suite + coverage gate (merged-union counting via `scripts/check-coverage.py`; the threshold is defined as `COVERAGE_THRESHOLD` in `.github/workflows/ci.yml` — that file is the source of truth, do not hardcode it here) and runs only when the diff touches C#-relevant paths (`.cs/.csproj/.sln/.props/.targets`, `TypstBridge/**`, the gate script, `ci.yml` itself); docs-only changes skip it (a skipped check counts as success).
- **Warnings are errors** (`TreatWarningsAsErrors` in `Directory.Build.props`) — the build must stay at 0/0.
- **Release (`release.yml`):** push a `v*` tag on `main` → native TypstBridge matrix (osx-arm64, linux-x64, win-x64) → `dotnet pack` of the six packages → NuGet push via trusted publishing (OIDC, keyless; policy + `NUGET_USER` secret already configured).
- Typst-dependent tests are env-gated: `OE_RUN_TYPST_COMPILE_TESTS=1 dotnet test …`. Brand-profile snapshots regenerate with `OE_UPDATE_SNAPSHOTS=1`.

## Entry Points

| Component | Run it |
|---|---|
| Unified CLI (create/edit/detect/merge/generate) | `dotnet run --project OfficeEditor.Cli -- <command>` |
| DOCX-only CLI | `dotnet run --project DocxEditor.Cli -- <command>` |
| API (deck sessions, previews, generate, demo endpoints) | `dotnet run --project OfficeEditor.Api --urls http://localhost:5001` |
| Web client (demo app) | `cd OfficeEditor.Web.Client && npm run dev` → http://localhost:5173 (`/demo` = the app) |
| API + web together | `make dev` (root Makefile) |
| MCP stdio host (JSON-RPC) | `dotnet run --project OfficeEditor.Mcp` |
| Example programs (all formats) | `dotnet run --project examples` |
| Convert tools | `dotnet run --project tools/convert-pptx -- <in.pptx> <out> [--format pdf\|png]` · `tools/convert-docx` (same shape) |
| Deck fidelity measurement | `python3 scripts/rmse.py <deck.pptx> [--format md\|json] [--no-render] [--force]` — renders ours, reuses/builds refs, per-slide RMSE % with mean/median/p85/p90/<10%/<15%/worst-5 |
| Visual regression | `dotnet run --project tools/visual-diff -- --suite pptx\|docx\|gen` |
| Benchmark (vs LibreOffice) | `dotnet run --project tools/pptx-benchmark` |

## Demo

Two demo paths, both fully local:

- **Web demo** — `make dev`, open http://localhost:5173/. Four tabs: *Render* (whitelisted REF decks → timed PNG/SVG gallery), *Generate* (edit `demo/demo-deck.json` title + theme presets → PPTX + previews + download), *Any render* (upload any .pptx), *Compare* (OfficeEditor Engine vs headless LibreOffice, side-by-side timings). All timings are server-measured.
- **CLI demo** — `make -f Makefile.demo demo` at repo root: preflight checks + Release build, renders `examples/REF/PPTX/sales_acceleration_deck.pptx` to PNGs+PDF, generates `demo/deck.json` → PPTX → render, prints per-slide timings, opens the PDFs. Output under `demo/output/` (gitignored); `make -f Makefile.demo clean` wipes it.

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
├── demo/                           # Demo decks (demo-deck.json, themes.json) + CLI demo output/
├── decks/                          # Authored generation decks (repo-overview.json)
├── local-ref/                      # Local-only dev fixtures (gitignored — never commit)
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

- `examples/REF/PPTX/sales_acceleration_deck.{pptx,pdf}` — PRIMARY: license-clean 16-slide sales deck with 5 SmartArt diagram parts (slide 15)
- `examples/REF/PPTX/AetherLink-Glass-Shareholder-Overview.{pptx,pdf}` — 15-slide styled glass deck; brand-profile / token-mining source
- `examples/REF/PPTX/northwind-investor-40.{pptx,pdf}` — 40-slide stress deck
- `examples/REF/PPTX/northwind-launch-review.{pptx,pdf}` — 12-slide pitch deck (slide-level solid backgrounds)
- `examples/REF/PPTX/northwind-demo.{pptx,pdf}` — self-made 15-slide PPTX smoke-test deck (generated by OfficeEditor)
- `examples/REF/DOCX/annual-report.{docx,pdf}` and `monitoring-report.{docx,pdf}` — license-clean templated DOCX (`{{variable}}` placeholders), python-docx authored with Word-rendered PDFs
- Generated outputs: `examples/output/ref/`
- PNG mode: `--format png` generates per-slide images at the configured PPI
- Token sets: `PptxEditor.Core/Generation/Design/` — mined from REF decks (see `Design/README.md`)
- SmartArt dev corpus: `local-ref/smartarts/`  (local-only, gitignored)
- Generation fixtures: `PptxEditor.Core/Generation/Fixtures/` — per-primitive parity test decks
- Demo/generation decks: `demo/demo-deck.json` (15-slide Northwind Labs — web demo + themes in `demo/themes.json`), `demo/deck.json` (6-slide Makefile demo deck), `decks/repo-overview.json` (self-description deck)
