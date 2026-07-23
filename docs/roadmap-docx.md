# DOCX roadmap

DOCX support in OfficeEditor is at mixed maturity: the `DocxToTypstConverter` (3,058 lines — styles cascade, lists, tables, images, shapes, headers/footers, sections, fields) is **beta+**, the `DocumentBuilder` fluent API is **beta**, and the instruction engine (`InstructionEngine`, 4 working ops) is **alpha**. The strategic arc mirrors what PPTX already achieved: **correctness → completeness → declarative generation → ecosystem parity**. Every phase ships with tests per AGENTS.md ("no half-tested features"); sizes are S (< 1 day), M (days), L (week+).

**Product priorities (owner-set):** **JSON → DOCX generation is the flagship DOCX deliverable** — its value is precisely that DOCX rendering is hard: free text boxes, shapes, floating pictures, wrap and z-order are painful to hand-write in raw OpenXML, so a declarative vocabulary that absorbs those quirks is the product. **DOCX rendering (docx → Typst → PDF/PNG) is the co-flagship**: it is the proof loop for generation (generate → render → compare, the same discipline as PPTX) and a user-facing feature in its own right. Phase order below is dependency order, but Phase 3 + the preview pipeline are the destination, not an afterthought.

## Phase 0 — Hygiene & correctness (do first)

Unblock public release: nothing here is new capability, everything is trust.

- **Fix or replace `examples/Docx/instructions/sample.json`** — uses 4 nonexistent op types (`AddHeading`, `AddBulletList`, `Replace`, `AddTable`); only `addParagraph` exists. Either rewrite it to the 4 working ops or (better) implement the missing ops in `InstructionEngine`. **Why:** first file a new user runs; currently broken. **S.** *Acceptance:* sample runs end-to-end via CLI; an integration test executes it.
- **Complete the instruction engine** — implement `AddRichContentInstruction` / `ReplaceWithRichContentInstruction` (currently fall into `NotSupportedException`) by delegating to the existing `DocumentBuilder.AddRichContent` / `ReplaceWithRichContent`. Add a loud schema validator for instruction JSON (unknown `type`, missing required fields → descriptive error, not silent skip). **S–M.** *Acceptance:* all 6 ops in `Models/Instructions.cs` dispatch; malformed JSON fails loudly with field-level messages.
- **Fix `DocxTemplateEngine` formatting loss** — variable merge rebuilds each paragraph as a single `Run` (`paragraph.Append(new Run(new Text(text)))`), destroying run-level formatting. Replace runs in-place, preserving `RunProperties` of the first affected run. **Why:** silent data corruption in the flagship mail-merge scenario. **M.** *Acceptance:* bold/italic/mixed-format paragraphs survive `MergeVariables`; regression test with a multi-run template; `MergeBatch` gets its first test.
- **Create the numbering part for lists** — `ContentBlockRenderer` hardcodes `NumberingId.Val = 1/2` without creating a `NumberingDefinitionsPart`, producing lists that only render if the template already defines those IDs. Create/lookup abstract numbering definitions on demand. **M.** *Acceptance:* bullet + ordered lists render in a document created from scratch; round-trip through Word/LibreOffice without repair prompts.
- **Restore DOCX visual-regression outputs** — `examples/output/ref/` has `pptx/` but no `docx/`; the visual-diff suite is a stub for DOCX. Generate reference outputs for the license-clean DOCX corpus (`examples/REF/DOCX/annual-report.docx`, `examples/REF/DOCX/monitoring-report.docx`) and wire them into `tools/visual-diff`. **M.** *Acceptance:* docx RMSE check runs in CI alongside the pptx suite.
- **README/example honesty pass** — `examples/Docx/README.md` and NuGet READMEs must state what works today (no images, no headers/footers via builder, instruction set = 4–6 edit ops). **S.** *Acceptance:* no documented feature throws `NotSupportedException`.

## Phase 1 — Builder completeness

Ordered by dependency; each item extends `DocumentBuilder` / `IDocumentBuilder` and reuses `ContentBlock` models where possible.

1. **Run-level formatting API** — bold/italic/underline/font/size/color/highlight on runs, not just paragraph styles. Foundation for everything below. **Why:** prerequisite for hyperlinks, table styling, markdown fidelity. **M.** *Acceptance:* fluent `AddFormattedRun(...)`; markdown bold/italic round-trips to run properties.
2. **Images / pictures** — `AddImage(stream, options)` with inline + anchored placement, sizing, alt text. Emitter logic already proven in the converter's image handling. **M.** *Acceptance:* image inserted from file and stream survives open-in-Word; converter round-trips it back to Typst.
3. **Hyperlinks** — external URL + internal bookmark anchors; depends on (1) for run styling. **S.** *Acceptance:* clickable link in Word; relationship part created correctly.
4. **Headers & footers API** — `AddHeader/Footer` per section with content blocks; converter already reads them, so semantics are known. **M.** *Acceptance:* header/footer visible in Word and in Typst preview output.
5. **Sections & page setup** — page size/margins/orientation, column count, section breaks; exposes what the converter already parses. **M.** *Acceptance:* multi-section document with mixed orientation; headers/footers (4) attach per-section.
6. **Table styling API** — column widths, cell shading, borders, header-row repeat, alignment; beyond the current plain-grid tables. **M.** *Acceptance:* styled table matches a reference render; converter round-trip preserves it.
7. **List numbering schemes** — user-definable abstract numbering (decimal, alpha, roman, bullet glyphs, multi-level) building on Phase 0's numbering-part creation. **M.** *Acceptance:* `AddList(items, scheme: NumberingScheme.LowerRoman, levels: 3)`; restart/continue semantics tested.
8. **Theme / color scheme support** — named theme colors + fonts referenced by styles, mirroring the PPTX token sets in `Generation/Design/`. **M–L.** *Acceptance:* a token set (colors/fonts/spacing) applies document-wide; styles reference theme colors by ID, never mutate template style definitions (AGENTS.md).

## Phase 2 — DOCX ⇄ Markdown

`MarkdownParser` + `AddMarkdown` already do md → docx. This phase closes the loop. **It does not depend on Phase 1:** the exporter *reads*, and the reading muscle already exists — the `DocxToTypstConverter` parses hyperlinks, images, lists, tables, headers/footers today. The exporter consumes the converter's block walk, not the builder's write API, so it can run in parallel with Phase 1.

- **`DocxToMarkdownExporter`** — walks the same block model the converter uses and emits CommonMark-ish: headings (style → `#` level), paragraphs, bold/italic/code runs, lists (numbering-aware, from Phase 0 numbering work), tables (pipe syntax), block quotes, fenced code, images as `![alt](path)` with optional asset extraction, hyperlinks as `[text](url)` (Phase 1.3). **Why:** explicit user need; makes OfficeEditor a two-way doc tool. **L.** *Acceptance:* exports `examples/REF/DOCX/annual-report.docx` to readable md; `md → AddMarkdown → docx → export → md` round-trip test is idempotent on the block level (exact byte equality not required).
- **Degradation strategy (documented)** — shapes/textboxes → blockquote callout with text content; anchored images → inline link + comment; headers/footers → HTML comment block; fields (PAGE) → dropped with comment. Lossy elements are enumerated in the exporter's docs, not silently swallowed.
- **Plumbing** — expose as `builder.ExportMarkdown()`, `OfficeEditor.Cli convert --format md`, and `tools/convert-docx --format md`. **S** once exporter exists.

## Phase 3 — JSON → DOCX generation (flagship)

Declarative vocabulary mirroring `PptxEditor.Core/Generation/`: **schema + loud validator → document model → OOXML emitter**, reusing `ContentBlock`/`ContentBlockRenderer` as the primitive layer. Distinct from the Phase 0 edit engine: **generate = new document from JSON; edit = ops against an existing document** — separate entry points, shared primitives.

The model has **two tiers** — this is what makes it a product and not a paragraph printer:

- **Flow tier** — sections → blocks (paragraph/heading/list/table/image/callout/pagebreak), styles, page setup. This is the ContentBlocks layer as it exists today.
- **Positioned tier** — the hard DOCX quirks, first-class: free text boxes (absolute-anchored, DrawingML `wsp` + VML fallback semantics), shapes (rect/line/callout with fill/stroke), floating pictures (anchor type, wrap mode — square/tight/through/top-bottom, distances), z-order, rotation. The converter's read side (3,058 lines covering exactly these cases) is the semantics reference: anything the positioned tier emits must round-trip through the converter's parser.

- **Document model + schema** — `Generation/Model`: document → sections → blocks (paragraph/heading/list/table/image/callout/pagebreak), styles, page setup. JSON Schema + loud validator per AGENTS.md conventions (unknown keys, bad enum values → field-path errors). **L.**
- **OOXML emitter** — builds a `WordprocessingDocument` from the model via `DocumentBuilder` internals; template-optional (blank or user template with style preservation). **L.**
- **Design tokens** — reuse Phase 1.8 theme support so generated documents can take a token set (fonts, colors, spacing) like PPTX decks do. **M.**
- **Fixtures + parity** — per-block fixture JSON files with rendered-reference outputs, wired into visual-diff (Phase 0), matching the "every primitive ships with emitters + fixture" rule. The render half of every fixture runs through the Phase 4 preview pipeline (converter → Typst) — generation and rendering prove each other. **M.**
- *Acceptance (phase):* a JSON document equivalent to the REF docx renders through the Typst pipeline within agreed RMSE of the builder-produced equivalent; a fixture exercising every positioned-tier primitive (anchored textbox, wrapped floating image, two overlapping shapes with defined z-order) round-trips through the converter without loss; validator rejects 100% of malformed fixture cases with useful messages.

## Phase 4 — Ecosystem parity

Bring DOCX to the surface area PPTX already has.

- **Preview pipeline (co-flagship — land early, with Phase 3)** — converter → Typst → PNG/SVG: extend `tools/convert-docx` and `ConversionService` beyond `pdf|typ` to `png|svg` (Typst 0.15.1 via TypstBridge supports both; PPTX path in `DeckGenerationService` is the template). This is the render half of the generation proof loop, so it ships alongside the vocabulary, not after it. **M.**
- **API document sessions + anatomy** — DOCX equivalent of deck sessions: upload → page previews/thumbnails → instruction application, plus a `document_anatomy` endpoint exposing the block tree. **L.**
- **`docx_*` MCP tools** — `docx_anatomize`, `docx_replace_element`, `docx_render_page`, `docx_generate`, mirroring the `deck_*` set in `OfficeEditor.Mcp`. **M.**
- **CLI `generate` for docx** — `office-editor generate --format docx deck.json` consuming the Phase 3 vocabulary. **S–M.**
- **NuGet README polish** — quickstart for builder, instructions, markdown, and JSON generation once they exist; badge the 0.1.0-unlisted → listed transition. **S.**

## Out of scope (and why)

- **Track changes / comments** — deferred: revision model is a large OOXML surface with low demand vs. generation features.
- **Footnotes / endnotes** — deferred: niche for the target doc-gen scenarios; converter already skips them predictably.
- **Fields beyond PAGE/NUMPAGES** — deferred: TOC/cross-ref fields require layout-dependent computation Typst-side can't replicate cheaply.
- **Macros / VBA** — never: security surface, out of product scope.
- **Full Word pixel fidelity** — non-goal; Typst preview is the spec of record (AGENTS.md).

## Release gates

**Phase 0 blocks any public release** — a broken sample and formatting-destroying merge are credibility killers. Everything in this roadmap targets **v0.5**: the phases are dependency-ordered within the doc, but the version is one train — hygiene (Phase 0) lands first inside it, generation and ecosystem parity (Phases 3–4) complete it.
