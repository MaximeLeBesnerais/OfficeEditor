# XLSX Roadmap Progress

## Phase 0 — Correctness & honesty

| Item | Status | Commit | Note |
|---|---|---|---|
| Ordered row/cell insertion | DONE | TBD | Sorted insert in GetOrCreateRow/GetOrCreateCell; tests for out-of-order rows and columns |
| Formula DataType fix | DONE | TBD | Removed forced CellValues.Number; formulas now omit DataType |
| Sheet-name validation | DONE | TBD | Duplicate, >31 chars, illegal chars, empty/whitespace checks with XlsxException |
| Column-name math past Z | DONE | TBD | Rewrote GetColumnName with standard algorithm; full-range 0-16383 test + boundary tests |
| Per-workbook table-ID | DONE | TBD | _nextTableId counter on WorkbookBuilder; cross-sheet uniqueness test |
| De-vaporware sample | pending | — | — |
| Doc honesty | pending | — | — |
| AddChart decision | DONE | TBD | XlsxException with Phase 4 reference; test verifies message content |
