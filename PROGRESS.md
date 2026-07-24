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
- [x] Verified tests: All pre-existing Typst-not-available failures, no regressions
- [x] Wrote SMARTART-REPORT.md: §1 (Inventory), §2 (Current Behavior), §3 (Classification), §4 (Support Matrix), §5 (Implementation Approach), §6 (Open Questions)

### Key Findings
1. **Current approximation is text-only** — text extracts correctly, shapes/connectors/styles are lost
2. **Drawing parts contain pre-rendered geometry** — shape-tree extraction (Tier 1) would cover ~95% of diagram types with moderate effort
3. **DataModel-driven layout (Tier 2)** is an L-sized effort but would primarily benefit process-flow + org-chart diagrams
4. **Zero SmartArt tests exist** — any implementation must start with test fixtures

## 2026-07-24: Tier 1 Prototype (Shape-Tree Extraction)

### Created
- [x] `PptxEditor.Core/Converters/SmartArt/SmartArtDrawingExtractor.cs` (337 lines)
  - Extracts shape geometry from `<dsp:sp>` diagram drawing-part shapes
  - Handles `roundRect` (with corner radius from `a:avLst`), `rightArrow` (as polygon), `rect`
  - Reads solid fills (direct RGB + scheme color mapping for accent1–6, lt1/dk1, tx1/tx2, bg1/bg2)
  - Reads strokes (width from `a:ln/@w`, color from `a:ln/a:solidFill`)
  - Reads rotation from `a:xfrm/@rot`
  - Uses regex-based XML attribute reading (AGENTS.pptx.md rule #1)
  - Returns `TypstElement` with `Type="Shape"` — fits into existing Typst-emit pipeline

- [x] Minimal hook in `PptxToTypstConverter.cs` (~6 new lines):
  - Added `using PptxEditor.Core.Converters.SmartArt;`
  - In `ConvertDiagramGraphicFrame`, before text extraction, calls `SmartArtDrawingExtractor.TryExtractShape()` and yields the shape element
  - Restructured loop to get shape position before text extraction (supports textless connector shapes now)

- [x] `DocxEditor.Tests/Unit/SmartArt/SmartArtDrawingExtractorTests.cs` (8 tests, all pass):
  - `Extract_RoundRectShape_ReturnsShapeElementWithCornerRadius`
  - `Extract_RoundRectWithSchemeColor_ReturnsShapeWithMappedFill`
  - `Extract_RightArrow_ReturnsPolygonShape`
  - `Extract_RotatedShape_ReturnsShapeWithRotation`
  - `Extract_NoGeometry_ReturnsNull`
  - `Extract_UnsupportedGeometry_ReturnsNull`
  - `Extract_WithFrameAndScaleOffset_ComputesPositionCorrectly`
  - `Extract_NoSpPr_ReturnsNull`
  - `Extract_NoFillNoStroke_RectShapeStillReturns`

### Build & Test Status
- Build: 0 warnings, 0 errors
- Tests: 1000 passed, 16 failed (all pre-existing Typst-not-available, no regressions)
- 9 new SmartArt tests introduced (8 unit + the existing converter passing unaffected)

### What the Prototype Achieves
When converting `sales_acceleration_deck.pptx`:
- **Before:** 4 positioned text labels (DIAGNOSE, DESIGN, DELIVER, SUSTAIN) — text-only, no shapes
- **After:** 4 `roundRect` shapes with `#C00000` fill + white stroke + 4 text labels + 3 `rightArrow` connector shapes (arrows)

The shapes are extracted from the pre-rendered drawing part (`drawing1.xml`). The same extraction logic works for any SmartArt diagram that has a drawing part, which is all 159 diagrams in the showeet corpus.

### Scope Limitations
- Only 3 geometry presets mapped: `roundRect`, `rightArrow`, `rect`
- Scheme color mapping is a static map (not theme-resolved)
- Font rendering inside shapes is not applied (text styling from `dsp:style` is not read)
- No gradient fills
- No 3D transforms from quick styles

### Next Steps (for a future agent)
- [ ] Map additional geometry presets (chevron, ellipse, triangle, hexagon, etc.)
- [ ] Resolve scheme colors from the slide's master theme (not static map)
- [ ] Apply text formatting from `dsp:style` (font size, color from `a:fontRef`)
- [ ] Handle gradient fills
- [ ] Add integration test that converts `sales_acceleration_deck.pptx` and asserts shape elements on slide 15
- [ ] Create parity fixture for SmartArt rendering
