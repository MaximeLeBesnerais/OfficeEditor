# DOCX Roadmap — Progress Log

## Phase 0 — Hygiene & correctness

| Item | Status | Commit | Note |
|---|---|---|---|
| 0.1 — Fix sample.json + integration test | DONE | `467e6dd` | Rewrote sample.json to use 4 working ops (create, addParagraph, replaceText, insertAfter); added integration test |
| 0.2 — Complete instruction engine | DONE | `2b5beaa` | Added addRichContent/replaceWithRichContent to both JSON & YAML parsers, engine dispatch, ContentBlock polymorphic deserialization, loud `DocxInstructionValidator` with field-level errors |
| 0.3 — Fix variable-merge formatting loss | DONE | `8efcdd9` | Replaced in-place run rebuild in DocxVariableReplacer.ReplaceVariablesAcrossRuns; preserves RunProperties of affected runs; DocxTemplateEngine has same single-Run rebuild pattern but is orphaned (no production callers) — noted, not fixed |
| 0.4 — Create numbering part for lists | DONE | `d95137e` | Added EnsureNumberingDefinitions() to DocumentBuilder with abstract numbering for bullet (ID 2) and decimal (ID 1); ContentBlockRenderer invokes it before rendering lists; 2 new tests verify definitions exist on from-scratch documents |
| 0.5 — Restore DOCX visual-regression outputs | SKIPPED | — | Requires `typst` CLI and TypstBridge native library unavailable in this environment; suite wiring exists in tools/visual-diff/SuiteCatalog.cs |
| 0.6 — README/example honesty pass | DONE | `19a3ea6` | Updated examples/Docx/README.md with 6-op instruction table, supported limitations list; no documented feature throws NotSupportedException |

## Phase 1 — Builder completeness

| Item | Status | Commit | Note |
|---|---|---|---|
| 1.3 — Hyperlinks | DONE | `5802cd5` | Added AddHyperlink(url, displayText, style?) to IDocumentBuilder; creates HyperlinkRelationship on MainDocumentPart; 2 tests |
| 1.1 — Run-level formatting API | NOT STARTED | | |
| 1.2 — Images / pictures | NOT STARTED | | |
| 1.4 — Headers & footers API | NOT STARTED | | |
| 1.5 — Sections & page setup | NOT STARTED | | |
| 1.6 — Table styling API | NOT STARTED | | |
| 1.7 — List numbering schemes | NOT STARTED | | |
| 1.8 — Theme / color scheme support | NOT STARTED | |

## Phase 2 — DOCX ⇄ Markdown

Not started.

## Phase 3 — JSON → DOCX generation

Not started.

## Phase 4 — Ecosystem parity

Not started.
