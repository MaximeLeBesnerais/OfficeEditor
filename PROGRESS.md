# DOCX Roadmap — Progress Log

## Phase 0 — Hygiene & correctness

| Item | Status | Commit | Note |
|---|---|---|---|
| 0.1 — Fix sample.json + integration test | DONE | `467e6dd` | Rewrote sample.json to use 4 working ops (create, addParagraph, replaceText, insertAfter); added integration test |
| 0.2 — Complete instruction engine | DONE | _(pending)_ | Added addRichContent/replaceWithRichContent to both JSON & YAML parsers, engine dispatch, ContentBlock polymorphic deserialization, loud `DocxInstructionValidator` with field-level errors |
| 0.3 — Fix variable-merge formatting loss | PENDING | | |
| 0.4 — Create numbering part for lists | PENDING | | |
| 0.5 — Restore DOCX visual-regression outputs | PENDING | | |
| 0.6 — README/example honesty pass | PENDING | | |

## Phase 1 — Builder completeness

Not started.

## Phase 2 — DOCX ⇄ Markdown

Not started.

## Phase 3 — JSON → DOCX generation

Not started.

## Phase 4 — Ecosystem parity

Not started.
