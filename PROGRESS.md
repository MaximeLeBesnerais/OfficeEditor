# XLSX Roadmap Progress

## Phase 0 — Correctness & honesty ✅ COMPLETE

| Item | Status | Commit | Note |
|---|---|---|---|
| Ordered row/cell insertion | DONE | `7733be3` | Sorted insert in GetOrCreateRow/GetOrCreateCell |
| Formula DataType fix | DONE | `7733be3` | Removed forced CellValues.Number; formulas omit DataType |
| Sheet-name validation | DONE | `7733be3` | Duplicate, >31 chars, illegal chars, empty/whitespace with XlsxException |
| Column-name math past Z | DONE | `7733be3` | Standard algorithm; full-range 0-16383 test + 13 boundary tests |
| Per-workbook table-ID | DONE | `7733be3` | `_nextTableId` counter on WorkbookBuilder; cross-sheet uniqueness |
| De-vaporware sample | DONE | `d5bca9c` | `_schema_status` field; README explicitly states "planned Phase 2" |
| Doc honesty | DONE | `dd68b20` | AGENTS.md + 3 slides in repo-overview.json qualified |
| AddChart decision | DONE | `7733be3` | XlsxException with Phase 4 reference |

## Phase 1 — Read API & edit foundations ✅ COMPLETE

| Item | Status | Commit | Note |
|---|---|---|---|
| GetCellValue (type-aware) | DONE | `52a59ad` | Shared-string resolution, numbers, formulas |
| GetCellFormula | DONE | `52a59ad` | Exposes raw formula text |
| Cell/range iteration | DONE | `52a59ad` | GetRange, GetRows, GetRow, GetDimensions |
| Null/missing-cell semantics | DONE | `52a59ad` | `null` for missing/non-existent cells; `CellExists` for disambiguation |
| Edit foundations | DONE | `52a59ad` | DeleteCell, DeleteRow, ClearRange, overwrite-in-place |

17 new tests; round-trip test verifies build→reopen→readback matches.

## Phase 2 — JSON → XLSX instruction engine ⚠️ PARTIAL

| Item | Status | Commit | Note |
|---|---|---|---|
| Schema v1 | DONE | `7ee223b` | XlsxInstructionSet, WorksheetInstruction, CellInstruction record types |
| Parser + loud validator | DONE | `7ee223b` | 10+ malformed-input rejection cases with actionable messages |
| Executor | DONE | `7ee223b` | Instructions → WorkbookBuilder; variable resolution; mixed formula/value rows |
| Row-replication loop | ⚠️ BLOCKED | — | Requires real formula parser (A1 refs, ranges, `$` anchors, sheet-qualified refs). Too complex for current context budget. Documented as remaining work. |
| CLI wiring | ⚠️ SKIPPED | — | OfficeEditor.Cli is outside the scope fence. Must be done from the main worktree. |

29 tests: parser (6), validator (10), executor (8), end-to-end (1).

## Phase 3 — Formatting & layout

Not started.

## Phase 4 — Charts & advanced

Not started.
