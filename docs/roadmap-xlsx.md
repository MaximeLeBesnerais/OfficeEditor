# XLSX roadmap

XLSX is the least complete format in OfficeEditor, but the XLSX hardening work on this branch closed the correctness and honesty gap: the variable pipeline (`XlsxVariableDetector`, `XlsxVariableReplacer`, `XlsxTemplateEngine`) is mature and now rewrites inline-string cells; `WorkbookBuilder`/`WorksheetBuilder` enforce **Excel's real coordinate bounds** (columns A–XFD, rows 1–1,048,576) and emit schema-ordered rows/cells; a **read API** exists (`GetCellValue`, `GetCellFormula`, `GetRange`, `GetRows`/`GetRow`, `GetDimensions`, `GetColumnWidth`/`GetRowHeight`, `GetMergeRanges`); and there is now a **working JSON instruction engine** (`XlsxEditor.Core/Instructions/`): parser + loud validator + executor over a v1 vocabulary — `version`, `worksheets[].name`, `headers`, `rows`, `cells[]` (`address`, `value` XOR `formula`), `variables`, and `style` (a raw numeric styleId string). The Phase-2 *rich* vocabulary is still out: no typed cells, `numberFormat`, images, merges, or column widths / row heights in JSON, no row-replication loop, and no CLI wiring. The arc: **correctness → read/write round-trip → declarative JSON → formatting → charts/advanced**. Phases are dependency-ordered; each ships with tests per repo convention.

**Product priorities (owner-set):** **JSON → XLSX generation is a co-flagship deliverable** (with JSON → DOCX) — the value is uniformity across the suite (same JSON-in → document-out story everywhere), AI-simplicity (a model can author one declarative shape per format), and organisation (reports as data, not code). The vocabulary must be rich from v1: typed cells, formulas, **pictures, merged cells, explicit column widths & row heights**, number formats, and the row-replication loop — not a headers-and-rows toy.

## Phase 0 — Correctness & honesty

Fix what would embarrass us in front of Excel before adding surface area.

**Phase 0 is complete** — every item below landed on the `night/xlsx-hardening` branch with tests beside `WorkbookBuilderTests`.

- **Ordered row/cell insertion** — **Done.** `GetOrCreateRow`/`GetOrCreateCell` insert at the correct sorted position (`InsertRowAtSortedPosition`/`InsertCellAtSortedPosition`), so shuffled writes (e.g. `AddCell("B10", …)` then `AddCell("A2", …)`) produce monotonic `<row>`/`<c>` sequences instead of non-monotonic ones that Excel may flag for repair. Tests write cells in shuffled order and verify the saved package validates.
- **Formula `DataType` fix** — **Done.** `AddCell(ref, formula, isFormula: true)` sets `<f>`, clears the stale `<v>`/`DataType`, and omits `DataType` so Excel infers the result type — no more `CellValues.Number` forced on string/bool-returning formulas (`=IF(...)`, `=CONCAT(...)`).
- **Sheet-name validation** — **Done.** `AddWorksheet` rejects empty, >31 chars, illegal `: \ / ? * [ ]`, leading/trailing apostrophe, and case-insensitive duplicates (the old silent dictionary overwrite that orphaning the first sheet's part is gone) — all with a typed `XlsxException`.
- **Column-name math past Z** — **Done.** Verified a non-issue; `GetColumnName` uses the standard algorithm and boundary tests (ZZ, Z→AA, ZZ→AAA) are pinned.
- **Per-workbook table-ID uniqueness** — **Done.** `WorkbookBuilder` owns `NextTableId()` and registers table display names workbook-wide (case-insensitive), scanning the max existing ID and names on open, so two tables on different sheets can no longer collide.
- **De-vaporware the sample** — **Done (decision: keep, now executable).** `examples/Xlsx/instructions/sample.json` is a v1 sample that `XlsxInstructionParser.ParseFromFile` + `XlsxInstructionExecutor.Execute` run end-to-end: the `_schema_status` "planned — not yet executable" header is gone and the "JSON Instructions" section of `examples/Xlsx/README.md` now documents the working engine.
- **Doc honesty** — **Open.** This roadmap and `examples/Xlsx/README.md` now tell the real story. `AGENTS.md` ("JSON instructions for XLSX are on the Phase 2 roadmap") and `decks/repo-overview.json` ("XLSX uses fluent C# APIs") now *understate* the engine; qualifying them is a cross-repo doc change outside the XLSX-doc scope of this branch.
- **`AddChart` decision** — **Done.** `IWorksheetBuilder.AddChart` now throws a descriptive `XlsxException` naming Phase 4, not a bare `NotSupportedException` on a public interface.

**Phase 0 size: M total. Complete — no known file-corruption paths remain, and Phase 1+ can build on this base.**

## Phase 1 — Read API & edit foundations

Unlock for any real "open → inspect → modify → save" workflow.

**Substantially complete on this branch.** The read API is real; the remaining gaps are typed returns and number-format/date coercion.

- **`GetCellValue(sheet, cellRef)`** — **Done, with caveats.** Shared-string and inline-string cells resolve to their text; number cells return their text form (`string?`); formula cells return the cached `<v>` value and `GetCellFormula` returns the stored formula with a leading `=`. Not yet: numeric returns as `double`, number-format-based date detection.
- **Cell/range iteration** — **Done.** `GetRange("A1:D10")`, `GetRows()` / `GetRow(int)`, and used-range discovery via `GetDimensions()`.
- **Null/missing-cell semantics** — **Done.** Reads return `null` for absent cells; `CellExists` distinguishes absent from present-but-empty; reads never throw for missing cells.
- **Edit foundations** — **Done.** Overwrite-in-place via sorted `GetOrCreateCell` (Phase 0); `DeleteCell`, `DeleteRow`, `ClearRange` for deletion.

**Why:** without reads, "edit existing workbook" (a stated product pillar) is impossible; the Phase 2 `edit` command depends on this. **Size: L.** **Acceptance:** the build → reopen → read-back round-trip is covered (`RoundTrip_BuildReopenReadback_ShouldMatchWrittenValues`, plus preservation tests asserting untouched content survives single-cell edits).

## Phase 2 — JSON → XLSX instruction engine

The stated #1 user need. The v1 engine has landed; the rich vocabulary and wiring are the remaining work. It mirrors `DocxEditor.Core/Instructions` (instruction objects, shared validator) and the PPTX pipeline's loud-validator posture.

- **Schema v1** — **Partially done.** The base vocabulary is implemented: `version`, `worksheets[].name`, `headers`, `rows`, `cells[]` (`address`, `value` XOR `formula`), `variables`, and `style` as a numeric styleId string. Extensions still pending: typed cells (`number`/`string`/`boolean`/`date`), `numberFormat`, and the **flagship** `images[]`, `merges[]`, `columnWidths` / `rowHeights` — merges and widths/heights exist in the fluent API but are **not** JSON-expressed yet.
- **Parser + loud validator** — **Done.** `XlsxEditor.Core/Instructions/`: unknown keys are rejected (typos fail loudly), sheet-name rules match `AddWorksheet`, addresses are checked against Excel's real bounds, cells must have `value` XOR `formula` (formulas must start with `=`), `version` is gated to `1.0`, and the unsupported `type`/`numberFormat` fields are rejected loudly rather than silently dropped. Domain `XlsxException`s per AGENTS.md.
- **Executor** — **Done.** Instruction objects → `WorkbookBuilder` calls with generation-time `{{var}}` resolution (unresolved placeholders fail loudly, never written verbatim); `Execute` re-runs the shared validator so programmatically-built sets are held to the same rules as parsed JSON.
- **Row-replication loop** — **Not started.** `repeat`/`foreach` over a data array with row-index-aware formula rewriting remains **the hardest item in this roadmap**: it needs real formula parsing (A1 refs, ranges, absolute `$` anchors, sheet-qualified refs), not regex. Treat it as its own work item with its own fixture set.
- **CLI wiring** — **Not started.** `officeeditor create out.xlsx --instructions file.json` is not wired; `generate` is PPTX-only; `edit` for xlsx is unimplemented.

**Size: XL (largest phase).** **Acceptance (status):** `sample.json` executes end-to-end producing a valid workbook ✓; the validator rejects 10+ malformed inputs with actionable messages ✓ (see `XlsxInstructionTests`); row-replication fixture ✗; rich fixture (images, merged title, explicit widths/heights, number formats) ✗.

## Phase 3 — Formatting & layout

Replace the raw `styleId` string hack (`AddCell(ref, value, styleId)` doing `uint.Parse`) with a real style model. Two items have already landed in the fluent API.

- **Style builder** — **Not started.** Fluent/declarative named styles (number formats, fonts, fills, borders, alignment) remain; `EnsureHeaderStyleIndex` is still the only style machinery.
- **Column widths & row heights** — **Partially done.** Explicit `SetColumnWidth(string column, double width)` and `SetRowHeight(int rowIndex, double height)` — Excel-bounds-validated, in-place updates, ranged-`<col>` splitting that preserves unrelated attributes — plus `GetColumnWidth`/`GetRowHeight` read-back. The `auto` width heuristic remains.
- **Merged cells** — **Done.** Fluent `MergeCells(string range)` / `UnmergeCells(string range)` / `GetMergeRanges()` with single-cell, reversed, duplicate and overlapping-range validation; `<mergeCells>` is emitted in schema position. (JSON `merges[]` is still Phase 2 scope.)
- **Freeze panes** — **Not started.** Top row / first column / arbitrary `SheetView` pane.
- **Conditional formatting (basic)** — **Not started.** Cell-value and expression rules with `DifferentialFormat`.

**Size: L.** **Acceptance:** a generated report fixture (styled headers, currency columns, frozen top row, merged title, conditional red/green variance column) opens in Excel/LibreOffice with zero repairs and renders as specified; styles also expressible in JSON and from fluent API with identical output. The merged-title and explicit-widths/heights parts of that fixture are expressible fluently today; styled headers, frozen panes and conditional variance are not.

## Phase 4 — Charts & advanced

- **`AddChart` for real** — the interface decision is **resolved** (`AddChart` now throws a descriptive `XlsxException` naming Phase 4); the actual bar/line/pie/column `DrawingsPart` + `ChartPart` implementation is **not started**. **L.**
- **Defined names** — **Not started.** Workbook-scoped named ranges; usable in formulas. **S.**
- **Data validation basics** — **Not started.** List/number-range/date rules with prompts and error alerts. **M.**
- **Autofilter API beyond tables** — **Not started.** Plain-range autofilter without requiring a `TableDefinitionPart`. **S.**
- **Multi-sheet stress tests** — **Not started.** 50+ sheets, cross-sheet formula webs, 100k-row sheet write/read round-trip within a time budget. **M.**

**Size: L.** **Acceptance:** chart fixture files validate against OOXML schema and render in Excel; if charts are cut instead, interface is clean and docs say so.

## Out of scope (and why)

- **Formula evaluation engine** — Excel/LibreOffice compute on open; duplicating the function library is a project, not a feature.
- **XLSX→PDF/Typst preview** — Typst can't evaluate Excel formulas, so previews would show wrong values; low value, high confusion. Not in v1.
- **Pivot tables** — pivot cache + records + definition parts are a large spec surface with thin demand versus charts.
- **Macros/VBA** — `vbaProject.bin` passthrough only; authoring macros is out of product scope (and a security surface).
- **Shared-formula optimization** — write one `<f>` per cell; file size cost is negligible at our target scales.

## Release gates

Everything in this roadmap targets **v0.5** — one version train. Phase 0 has landed (no known file-corruption paths, no vaporware in XLSX docs) and Phase 1's read API is in. The Phase 2 v1 engine is in, but the rich vocabulary (typed cells, images, merges/widths/heights in JSON), row-replication and CLI wiring are what make the "JSON → XLSX" story complete. Phase 3 completes formatting (merged cells and explicit widths/heights are already fluent; the style builder, freeze panes and conditional formatting are not), and charts (Phase 4) ship **or** are explicitly cut by release day — limbo is not acceptable.
