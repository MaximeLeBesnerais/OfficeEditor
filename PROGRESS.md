# XLSX Roadmap Progress

## Phase 0 — Correctness & honesty ✅ (COMPLETE)

| Item | Status | Commit | Note |
|---|---|---|---|
| Ordered row/cell insertion | DONE | `7733be3` | Sorted insert in GetOrCreateRow/GetOrCreateCell; tests for out-of-order rows and columns |
| Formula DataType fix | DONE | `7733be3` | Removed forced CellValues.Number; formulas now omit DataType; Excel infers it |
| Sheet-name validation | DONE | `7733be3` | Duplicate, >31 chars, illegal chars (`:\/?*[]`), empty/whitespace checks with XlsxException |
| Column-name math past Z | DONE | `7733be3` | Rewrote GetColumnName with standard algorithm; full-range 0-16383 test + 13 boundary tests |
| Per-workbook table-ID | DONE | `7733be3` | `_nextTableId` counter on WorkbookBuilder; `NextTableId()` method; loads max ID from existing tables |
| De-vaporware sample | DONE | `d5bca9c` | Added `_schema_status` field to sample.json; README explicitly states "planned — Phase 2, not yet executable" |
| Doc honesty | DONE | `dd68b20` | AGENTS.md: qualified XLSX as fluent-API + variables; 3 overclaiming slides in repo-overview.json fixed |
| AddChart decision | DONE | `7733be3` | XlsxException with Phase 4 reference; test verifies message contains "Phase 4" |

## Phase 1 — Read API & edit foundations

Not started.

## Phase 2 — JSON → XLSX instruction engine

Not started.

## Phase 3 — Formatting & layout

Not started.

## Phase 4 — Charts & advanced

Not started.
