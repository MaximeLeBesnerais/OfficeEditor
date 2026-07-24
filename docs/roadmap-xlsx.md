# XLSX roadmap

XLSX is the least complete format in OfficeEditor: the variable pipeline (`XlsxVariableDetector`, `XlsxVariableReplacer`, `XlsxTemplateEngine`) is mature, but `WorkbookBuilder`/`WorksheetBuilder` are alpha-grade with known latent bugs, there is **no read API** (so "edit existing workbook" is not a real story), and **no instruction engine** — `examples/Xlsx/instructions/sample.json` is orphaned vaporware that nothing parses. The arc: **correctness → read/write round-trip → declarative JSON → formatting → charts/advanced**. Phases are dependency-ordered; each ships with tests per repo convention.

**Product priorities (owner-set):** **JSON → XLSX generation is a co-flagship deliverable** (with JSON → DOCX) — the value is uniformity across the suite (same JSON-in → document-out story everywhere), AI-simplicity (a model can author one declarative shape per format), and organisation (reports as data, not code). The vocabulary must be rich from v1: typed cells, formulas, **pictures, merged cells, explicit column widths & row heights**, number formats, and the row-replication loop — not a headers-and-rows toy.

## Phase 0 — Correctness & honesty

Fix what would embarrass us in front of Excel before adding surface area.

- **Ordered row/cell insertion** — `WorksheetBuilder.GetOrCreateRow`/`GetOrCreateCell` blindly `Append`, so out-of-order writes (e.g. `AddCell("B10", ...)` then `AddCell("A2", ...)`) produce non-monotonic `<row>`/`<c>` sequences that Excel may flag for "repair". Insert at the correct sorted position instead. **Why:** silent file corruption on a valid use pattern. **S.** **Acceptance:** test writes cells in shuffled order; saved file opens clean; row/cell order verified via OpenXML inspection.
- **Formula `DataType` fix** — `AddCell(ref, formula, isFormula: true)` forces `CellValues.Number`, which is wrong for string/bool-returning formulas (e.g. `=IF(...)`, `=CONCAT(...)`). Omit `DataType` for formula cells and let Excel infer, or accept an explicit result-type hint. **S.** **Acceptance:** string-returning formula round-trips without a repair prompt.
- **Sheet-name validation** — `AddWorksheet` accepts anything: duplicates, >31 chars, `: \ / ? * [ ]`. Duplicates are worse than a validation gap: `_worksheets[name] = …` silently overwrites the dictionary entry, orphaning the first sheet's part while both `Sheet` entries remain in the workbook. Loud validation with a domain exception (per AGENTS.md), checked against `_worksheets`. **S.** **Acceptance:** each invalid case throws a typed exception with a clear message.
- **Column-name math past Z** — `GetColumnName` (`WorksheetBuilder.cs`) breaks at wrap boundaries (0-based index 676 yields `ZA` where Excel expects `ZZ`). Fix the base-26 loop and pin the boundaries (ZZ itself, Z→AA, ZZ→AAA) with tests. **S.** **Acceptance:** generated references match Excel's naming across the full column range.
- **Per-workbook table-ID uniqueness** — `AddTable` assigns `Id = TableDefinitionParts.Count() + 1` *per worksheet part*, so two tables on different sheets collide (Excel requires workbook-wide unique table IDs). Track next table ID on `WorkbookBuilder`. **S.** **Acceptance:** tables on two sheets → distinct IDs in saved package.
- **De-vaporware the sample** — either delete `examples/Xlsx/instructions/sample.json` + the "JSON Instructions" section of `examples/Xlsx/README.md`, or mark it explicitly as *planned schema, not yet executable*. **S.** **Acceptance:** no repo doc claims a working JSON→XLSX path.
- **Doc honesty** — the root `README.md` was rewritten and is now accurate, but `AGENTS.md` ("DOCX, PPTX, XLSX — via instruction sets (JSON/YAML)") and `decks/repo-overview.json` ("The same instruction model drives DOCX, PPTX and XLSX") still overclaim XLSX. Qualify both: instructions are DOCX/PPTX today; XLSX is fluent-API + variables only. **S.**
- **`AddChart` decision** — `IWorksheetBuilder.AddChart` throws `NotSupportedException`. Replace with a descriptive `UnsupportedXlsxException` (docs pointing at Phase 4) or remove from the interface until Phase 4. Do not leave a bare `NotSupportedException` on a public interface. **S.** **Acceptance:** calling it throws a typed exception whose message names the roadmap phase.

**Phase 0 size: M total. Blocks everything — do not build on a corrupting base.**

## Phase 1 — Read API & edit foundations

Unlock for any real "open → inspect → modify → save" workflow.

- **`GetCellValue(sheet, cellRef)`** with type-aware returns: shared-string resolution, numbers as `double`, booleans, dates via number-format detection; formula cells return cached value **and** expose `GetCellFormula` for the raw formula string. **M.**
- **Cell/range iteration** — `GetRange("A1:D10")`, row enumeration (`GetRows()` / `GetRow(int)`), used-range discovery (`GetDimensions()` → `SheetDimension` or computed min/max). **M.**
- **Null/missing-cell semantics** — define and document: empty cell vs absent cell vs empty string; consistent return (`null` + `CellExists`) rather than exceptions for reads. **S.**
- **Edit foundations** — overwrite-in-place via existing `GetOrCreateCell` (now sorted from Phase 0); delete cell/row/clear-range operations. **M.**

**Why:** without reads, "edit existing workbook" (a stated product pillar) is impossible; the Phase 2 `edit` command depends on this. **Size: L.** **Acceptance:** round-trip test — build a workbook with `WorkbookBuilder`, reopen it, assert every value/formula/dimension read back matches; edit one cell of a pre-existing file and verify all untouched content survives byte-level part comparison.

## Phase 2 — JSON → XLSX instruction engine

The stated #1 user need. Mirror `DocxEditor.Core/Instructions` (instruction objects, strategy executors) and the PPTX generation pipeline's loud-validator posture.

- **Schema v1** — seed from `examples/Xlsx/instructions/sample.json` (which is already plausible): `version`, `worksheets[].name`, `headers`, `rows`, `cells[]` (`address`, `value` | `formula`), `variables`. Extend with: typed cells (`number`/`string`/`boolean`/`date`), `numberFormat`, basic `style` reference, and — **flagship scope, not Phase 3 leftovers** — `images[]` (anchor cell, size, alt text), `merges[]` (range, overlap-validated), `columnWidths` / `rowHeights` (explicit per index, plus an `auto` width heuristic). **M.**
- **Parser + loud validator** — `XlsxEditor.Core/Instructions/`: schema validation with precise errors (sheet name, address syntax, unknown keys rejected, `value` XOR `formula`). Custom domain exceptions per AGENTS.md. **M.**
- **Executor** — instruction objects → `WorkbookBuilder` calls; `{{var}}` resolution at generation time via the existing variables data flow (generation-time substitution, not `XlsxVariableReplacer` post-pass). **M.**
- **Row-replication loop** — `repeat`/`foreach` block over a data array that clones a row template N times with per-iteration `{{item.field}}` binding and row-index-aware formula rewriting (`=SUM(B2:D2)` → row N). This is the gap `XlsxTemplateEngine` (cell-text-scoped `{{#each}}` only) cannot fill — and **the hardest item in this roadmap**: formula shifting requires real formula parsing (A1 refs, ranges, absolute `$` anchors, sheet-qualified refs), not regex. Treat it as its own work item with its own fixture set. **L.**
- **CLI wiring** — `officeeditor create out.xlsx --instructions file.json`; extend `officeeditor generate` (currently PPTX-only) to dispatch on `--type xlsx` / file extension; `edit` for xlsx consumes set-cell/insert-row/delete instructions on top of Phase 1. **M.**

**Size: XL (largest phase).** **Acceptance:** `sample.json` executes end-to-end producing a workbook identical to the equivalent fluent-API output; validator rejects 10+ malformed inputs with actionable messages; row-replication fixture generates a 100-row table with correct shifted formulas; a rich fixture (images, merged title, explicit widths/heights, number formats) opens in Excel/LibreOffice with zero repairs; tests live beside `WorkbookBuilderTests`/`XlsxTemplateEngineTests`.

## Phase 3 — Formatting & layout

Replace the raw `styleId` string hack (`AddCell(ref, value, styleId)` doing `uint.Parse`) with a real style model.

- **Style builder** — fluent/declarative named styles: number formats (currency, percent, date), fonts, fills, borders, alignment; dedupe into `CellFormats` (generalize `EnsureHeaderStyleIndex`); styles referenceable from JSON schema v1.1. **L.**
- **Column widths & row heights** — explicit set + auto-width heuristic for generated files (max content length per column, capped). **M.**
- **Merged cells** — `MergeCells` part with overlap validation. **S.**
- **Freeze panes** — top row / first column / arbitrary `SheetView` pane. **S.**
- **Conditional formatting (basic)** — cell-value and expression rules with `DifferentialFormat`; data bars/color scales optional stretch. **M.**

**Size: L.** **Acceptance:** a generated report fixture (styled headers, currency columns, frozen top row, merged title, conditional red/green variance column) opens in Excel/LibreOffice with zero repairs and renders as specified; styles also expressible in JSON and from fluent API with identical output.

## Phase 4 — Charts & advanced

- **`AddChart` for real** — bar/line/pie/column via `DrawingsPart` + `ChartPart` (DrawingML charts), anchored to a range; or permanently remove `AddChart`/`ChartType` from `IWorksheetBuilder` and document the decision. No middle state. **L.**
- **Defined names** — workbook-scoped named ranges; usable in formulas. **S.**
- **Data validation basics** — list/number-range/date rules with prompts and error alerts. **M.**
- **Autofilter API beyond tables** — plain-range autofilter without requiring a `TableDefinitionPart`. **S.**
- **Multi-sheet stress tests** — 50+ sheets, cross-sheet formula webs, 100k-row sheet write/read round-trip within a time budget (address the O(n)-per-write shared-string lookup in `WorkbookBuilder` first). **M.**

**Size: L.** **Acceptance:** chart fixture files validate against OOXML schema and render in Excel; if charts are cut instead, interface is clean and docs say so.

## Out of scope (and why)

- **Formula evaluation engine** — Excel/LibreOffice compute on open; duplicating the function library is a project, not a feature.
- **XLSX→PDF/Typst preview** — Typst can't evaluate Excel formulas, so previews would show wrong values; low value, high confusion. Not in v1.
- **Pivot tables** — pivot cache + records + definition parts are a large spec surface with thin demand versus charts.
- **Macros/VBA** — `vbaProject.bin` passthrough only; authoring macros is out of product scope (and a security surface).
- **Shared-formula optimization** — write one `<f>` per cell; file size cost is negligible at our target scales.

## Release gates

Everything in this roadmap targets **v0.5** — one version train. Within it: Phase 0 lands first (no known file-corruption paths, no vaporware in docs — it blocks any public XLSX claim), Phases 1–2 are the minimum credible "JSON → XLSX" story, Phase 3 completes formatting, and charts (Phase 4) ship **or** are explicitly cut by release day — limbo is not acceptable.
