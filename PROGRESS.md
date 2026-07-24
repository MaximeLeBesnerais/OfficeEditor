# SmartArt Investigation — Progress Log

## 2026-07-24: Initial Investigation

### Completed
- [x] Read AGENTS.md and AGENTS.pptx.md — understood repo conventions, PPTX rules
- [x] Explored SmartArt code in `PptxToTypstConverter.cs` (lines 1986–2340):
  - `ConvertDiagramGraphicFrame`: extracts text from drawing-part shapes, positions them, fires warnings
  - No shape geometry, fill, strokes, or connectors extracted
- [x] Analyzed `sales_acceleration_deck.pptx` slide 15:
  - 1 SmartArt diagram: `process5` layout, 4 nodes (SUSTAIN, DIAGNOSE, DESIGN, DELIVER)
  - `roundRect` nodes + `rightArrow` connectors
  - Drawing part (`drawing1.xml`) contains pre-rendered shape tree with positions, fills, strokes
- [x] Analyzed showeet dev corpus (`local-ref/smartarts/smartarts.pptx`):
  - 159 diagrams across 9 categories (process 42, list 33, picture 25, relationship 18, cycle 16, hierarchy 13, pyramid 5, matrix 5, officeonline 2)
  - All 159 have drawing parts (2008 namespace extension)
  - 5 algorithm types observed: `snake`, `sp`, `tx`, `hierChild`, `conn`
- [x] Verified build: 0 warnings, 0 errors
- [x] Verified tests: 991 passed, 16 failed (pre-existing Typst-not-available failures, unrelated)
- [x] Wrote SMARTART-REPORT.md: §1 (Inventory), §2 (Current Behavior), §3 (Classification), §4 (Support Matrix), §5 (Implementation Approach), §6 (Open Questions)

### Key Findings
1. **Current approximation is text-only** — text extracts correctly, shapes/connectors/styles are lost
2. **Drawing parts contain pre-rendered geometry** — shape-tree extraction (Tier 1) would cover ~95% of diagram types with moderate effort
3. **DataModel-driven layout (Tier 2)** is an L-sized effort but would primarily benefit process-flow + org-chart diagrams
4. **Zero SmartArt tests exist** — any implementation must start with test fixtures

### Next Steps
- [ ] Optional: prototype Tier 1 shape-tree extraction for `process5` (sales_acceleration_deck slide 15)
- [ ] Write FINAL-REPORT.md
- [ ] Commit all work
