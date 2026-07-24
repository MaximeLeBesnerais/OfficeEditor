# SmartArt Investigation Report

> **Date:** 2026-07-24 | **Branch:** `agent/smartart-investigation` | **Status:** Phase 1 readiness assessment

---

## 1. Inventory: SmartArt Diagram Types Across the Corpus

### 1.1 Sales Acceleration Deck (`examples/REF/PPTX/sales_acceleration_deck.pptx`)

Slide 15 contains **1 SmartArt diagram** (Diagram 16, id=17):

| Property | Value |
|---|---|
| Layout type | `urn:microsoft.com/office/officeart/2005/8/layout/process5` |
| Layout category | `process` |
| Quick style | `urn:microsoft.com/office/officeart/2005/8/quickstyle/simple1` (category `simple`) |
| Color style | `urn:microsoft.com/office/officeart/2005/8/colors/accent1_2` (category `accent1`) |
| Node count | 4 nodes (SUSTAIN, DIAGNOSE, DESIGN, DELIVER) |
| Node shape | `roundRect` (with `adj=0.1` corner radius) |
| Connectors | 3 `rightArrow` shapes between sequential nodes |
| Algorithm | `snake` with `flowDir=row`, direction-dependent `grDir` |
| Drawing part | `ppt/diagrams/drawing1.xml` — pre-rendered `dsp:spTree` with positioned shapes |

The diagram renders as a **2×2 process grid** (top row: DIAGNOSE → DESIGN, bottom row: SUSTAIN → DELIVER) with arrows between horizontal neighbors and a vertical dogleg arrow from DESIGN down to DELIVER.

### 1.2 Showeet Dev Corpus (`local-ref/smartarts/smartarts.pptx` — gitignored, never commit)

The dev corpus contains **159 SmartArt diagrams** across **9 categories**. Below are category counts and representative layout types (layout names generic — no specific showeet content committed):

| Category | Count | % | Representative layout types (algorithm) |
|---|---|---|---|
| **process** | 42 | 26.4% | process1–5, chevron1–2, hProcess3–11, vProcess5, StepUp/DownProcess, CircleProcess, CircleArrowProcess, CircleAccentTimeline, equation1–2, arrow2–6, pieProcess, PhasedProcess, IncreasingArrows/CircleProcess, ConvergingText, DescendingProcess, InterconnectedBlockProcess, RandomtoResultProcess, SubStepProcess |
| **list** | 33 | 20.8% | list1, vList2–6, hList1–9, pList1–2, bList2, LinedList, VerticalAccentList, VerticalCurvedList, VerticalCircleList, SquareAccentList, PictureAccentList, BracketList, BlockDescendingList, TabList, target3, VaryingWidthList, AlternatingHexagons, AlternatingPictureBlocks, architecture, hierarchy3, lProcess2 |
| **picture** | 25 | 15.7% | HexagonCluster, AccentedPicture, BendingPictureCaption/Blocks/List/SemiTransparentText, BubblePictureList, CaptionedPictures, CircularPictureCallout, FramedTextPicture, PictureAccentBlocks/List, PictureGrid, PictureLineup, PictureStrips, RadialPictureList, SnapshotPictureList, SpiralPicture, ThemePictureAccent/Grid/AlternatingAccent, TitledPictureBlocks, TitlePictureLineup |
| **relationship** | 18 | 11.3% | venn1–3, balance1, funnel1, gear1, target1, arrow1–5, radial2/4, rings+Icon, CircleRelationship, OpposingIdeas, PlusandMinus, ReverseList |
| **cycle** | 16 | 10.1% | cycle1–8, radial1/3/5/6, RadialCluster, HexagonRadial, gear1, chart3 |
| **hierarchy** | 13 | 8.2% | hierarchy1–2/4–6, orgChart1, HorizontalOrganizationChart, HorizontalMultiLevelHierarchy, HalfCircleOrganizationChart, NameandTitleOrganizationalChart, CirclePictureHierarchy, pictureOrgChart+Icon |
| **pyramid** | 5 | 3.1% | pyramid1–4 |
| **matrix** | 5 | 3.1% | matrix1–3, cycle4 |
| **officeonline** | 2 | 1.3% | Picture Frame, TabbedArc+Icon |

### 1.3 Diagram XML Data Model Structure

Every SmartArt graphic in a PPTX consists of **5 interrelated XML parts**, referenced from the slide via `<dgm:relIds>` inside a `<p:graphicFrame>`:

| Part | Relationship attr | Namespace | Content |
|---|---|---|---|
| **data** | `r:dm` | `dgm:dataModel` | Node tree: `<dgm:ptLst>` (points with modelId, type, text, shape type) + `<dgm:cxnLst>` (connections with srcId→destId, transition points) |
| **layout** | `r:lo` | `dgm:layoutDef` | Layout algorithm tree: `<dgm:layoutNode>` hierarchy with `<dgm:alg type="…"/>` (snake, hierChild, sp, tx, conn), `<dgm:shape>`, `<dgm:forEach>`, constraints, rules |
| **quickStyle** | `r:qs` | `dgm:styleDef` | Style labels: `<dgm:styleLbl>` per named style (`node0`, `node1`, `sibTrans2D1`, etc.) with 3D scene, shape3D, text properties, font/line/fill/effect references |
| **colors** | `r:cs` | `dgm:colorsDef` | Color labels: `<dgm:styleLbl>` with fill/line/text color lists (scheme → tint/shade/lum/sat transforms) |
| **drawing** | (via `dgm:extLst` in data, 2008 ns) | `dsp:drawing` | Pre-rendered shape tree: `<dsp:spTree>` containing `<dsp:sp>` elements with positions (`a:xfrm`), geometry presets (`roundRect`, `rightArrow`), fills, strokes, text bodies, and `modelId` back-references |

**Key layout algorithms** observed:

| Algorithm type | Description | Used by |
|---|---|---|
| `snake` | Multi-row flow with direction control (`grDir`, `flowDir`, `contDir`) | process5, chevron1, vProcess5 |
| `sp` | Space-filling (pack shapes into a container) | pList1, vList2, hList3, pyramid2, funnel1, most list/picture layouts |
| `tx` | Text-only layout (position text at node position) | venn1–2, matrix1–3, pyramid1, arrow1/3–5, radial2/4, cycle2 |
| `hierChild` | Hierarchical layout (parent → children with offsets) | hierarchy1–2/4–6, orgChart1, all org-chart variants |
| `conn` | Connector routing (arrows between nodes) | sibTrans nodes in process5, arrow6, radial2 |

### 1.4 Node Point Types

The data model defines point types via `type` attribute on `<dgm:pt>`:

| Type | Role |
|---|---|
| `doc` | Root/document node (container, no visual) |
| `node` | Visual shape node (has text, positioned by algo) |
| `parTrans` | Parent transition (invisible routing between parent→child) |
| `sibTrans` | Sibling transition (invisible routing between siblings) |
| `pres` | Presentation override (references an associated data point, carries style idx + layout vars) |
| (untyped) | Defaults to `node` behavior |

---

## 2. Current Behavior: How `PptxToTypstConverter` Handles SmartArt

### 2.1 Code Location

All SmartArt-handling code lives in `PptxEditor.Core/Converters/PptxToTypstConverter.cs`:

| Lines | Method / Region | Purpose |
|---|---|---|
| 1986–2008 | `graphicFrame.GraphicData.Uri` check | Dispatch: if URI contains `/drawingml/2006/diagram`, call `ConvertDiagramGraphicFrame` |
| 1997–2006 | Warning emission | Fires `SlideWarning` on every SmartArt encounter ("approximated as positioned text") |
| 2003–2005 | Fallback placeholder | If approximation yields nothing, emits a gray box with "SmartArt diagram (not supported)" |
| 2096–2198 | `ConvertDiagramGraphicFrame` | Main extraction method (see §2.2) |
| 2200–2204 | `_diagramNamespaces` | Namespace constants (2006 and 2008 diagram namespaces) |
| 2206–2214 | `ExtractTextFromDiagramShape` | Extracts `txBody` text from a diagram shape |
| 2291–2340 | `GetDiagramShapePosition` | Reads shape position from `txXfrm` (2008 ns) or `spPr/xfrm` (2006 ns) |

### 2.2 Approximation Algorithm (Current)

`ConvertDiagramGraphicFrame` follows this path:

1. **Part resolution** (lines 2103–2149):
   - Looks for `dgm:relIds` element in the graphic data → extracts all `r:dm`, `r:lo`, `r:qs`, `r:cs` relationship IDs
   - Resolves each `rId` to an `OpenXmlPart` via `slidePart.GetPartById()`
   - Adds child parts transitively
   - Falls back to `slidePart.GetPartsOfType<DiagramPersistLayoutPart>()`
   - Final fallback: iterates ALL slide parts (brute force)

2. **Shape extraction** (lines 2152–2197):
   - For each candidate part, looks for `dsp:spTree` → `<dsp:sp>` elements where `LocalName == "sp"` and namespace is a diagram namespace
   - For each matching shape:
     - Extracts text via `ExtractTextFromDiagramShape()`: finds `<dsp:txBody>`, delegates to `ExtractTextFromTextBody()`
     - Gets position via `GetDiagramShapePosition()`: tries `<dsp:txXfrm>` first (2008 ns), then `<dsp:spPr>/<a:xfrm>`
     - Emits a `TypstElement` of type `"Text"` at the shape's position

3. **What gets through**:
   - **Text content only** — positioned at the correct (x, y) coordinates
   - **No shapes** — no roundRect, no arrows, no fills, no strokes
   - **No connectors** — connector shapes (sibling/parent transitions) are ignored if they have no text
   - **No styles** — colors from scheme/accent are not resolved

4. **Warnings** (lines 1997–2005):
   - Every SmartArt diagram triggers: `"SmartArt diagram '<name>' was approximated as positioned text; diagram layout and styling may differ from the original."`
   - If no text was extracted: `"SmartArt diagram '<name>' could not be approximated and was replaced by a placeholder."`

### 2.3 What the Sales Acceleration Deck Produces

When converting `sales_acceleration_deck.pptx`, slide 15 produces **4 positioned text labels** (DIAGNOSE, DESIGN, DELIVER, SUSTAIN) at their correct positions. The `roundRect` shapes, `rightArrow` connectors, red fills (`accent1` scheme), white strokes, and white text colors are **all lost**. The warning fires once for the diagram.

### 2.4 What Tests Exist

Searching the test suite for SmartArt/diagram tests:
- **None.** There are zero SmartArt-specific unit tests or integration tests in the test suite today.
- The converter tests (`PptxToTypstConverterSingleSlideTests`, `PptxToTypstConverterFullDocumentTests`) exercise SmartArt slides only indirectly through the REF corpus conversions; no assertions on SmartArt output quality exist.

---

## 3. Classification per Diagram Type

### 3.1 Classification Criteria

| Tier | Label | Criterion |
|---|---|---|
| **A** | Renders structurally | Node text + shapes + connectors are all rendered; orientation and relative positions match the original. Style colors are approximated. |
| **B** | Degrades gracefully | Node text is rendered at correct positions; shapes/connectors are missing or replaced with simple boxes. Informational content (text hierarchy) is preserved. |
| **C** | Broken | Text is missing, misplaced, or garbled. Placeholder box shown. |

### 3.2 Classification of Top Categories (based on current converter)

| Category | Current Tier | Evidence |
|---|---|---|
| **process** | **B** (degrades gracefully) | Text extracted; positions correct. Shapes (chevrons, arrows, circles) and connectors missing. Example: `sales_acceleration_deck` slide 15 — text is readable but boxes/arrows absent. |
| **list** | **B** (degrades gracefully) | Text extracted from list item shapes. Positioning depends on layout algorithm output in drawing part. Vertical lists likely degrade better than complex grid layouts. |
| **picture** | **B–C** (varies) | Picture-based SmartArt has images as primary content; text extraction alone misses the diagram's purpose. Some picture layouts may have text labels; those extract but the picture context is lost. |
| **relationship** | **B** (degrades gracefully) | Text labels for relationship nodes extract. Connector lines/arrows between nodes are lost. Venn diagrams lose their overlapping-circle semantics. |
| **cycle** | **B** (degrades gracefully) | Cyclic arrangement text extracts but circular layout context and cycle arrows are lost. Radial/gear variants degrade to scattered text. |
| **hierarchy** | **B** (degrades gracefully) | Org-chart text extracts with names/titles but the hierarchical layout (horizontal bands, reporting lines) collapses to plain text. |
| **pyramid** | **B** (degrades gracefully) | Pyramid segment text extracts; the triangular arrangement and segment proportions are lost. |
| **matrix** | **B** (degrades gracefully) | Quadrant/axis labels extract but the matrix grid structure and axis labels are lost. |
| **officeonline** | **C** (broken) | TabbedArc+Icon and PictureFrame are newer layout types; may not render in 2008 drawing namespace. |

### 3.3 Why No Type Reaches Tier A

The current converter intentionally skips:
- **Shape geometry** — `<dsp:spPr>` with `<a:prstGeom prst="roundRect">`, `<a:xfrm rot="5400000">` (rotation), `<a:adjLst>` (adjustment handles) are not parsed
- **Shape fills** — `<a:solidFill><a:schemeClr val="accent1"/></a:solidFill>` is not read
- **Shape strokes** — `<a:ln w="25400">` border outlines are not read
- **Connectors** — Shapes whose `txBody` has no content (e.g., arrow-only shapes) are skipped
- **Group-level transforms** — The `dsp:spTree` root-level `grpSpPr/xfrm` offset is not applied

The **drawing part** (`drawing1.xml`) already contains all the pre-rendered geometry — positions, fills, strokes, and geometry presets. The converter just needs to read these XML elements and emit corresponding Typst primitives.

### 3.4 Algorithm-Driven Layouts (Not in Drawing Part)

Some SmartArt diagrams store layout *algorithmically* via `dgm:layoutDef` (the layout part) rather than pre-rendered shapes in the drawing part. These are:
- **`tx` algorithm types** — text-only nodes, no shapes. Already handled adequately (text extraction is the right approach for these).
- **Edge cases where drawing part is missing** — if a diagram has only layout definition but no drawing part, the converter today fails silently (placeholder).
- **`conn` algorithm types** — connector routing (arrows between nodes computed from source/target positions). The drawing part generally contains the pre-computed connector shapes for common diagram types.

---

## 4. Proposed Support Matrix

### 4.1 Philosophy

Rather than implementing the OOXML SmartArt layout engine (which would require parsing `dgm:layoutDef` algorithms like `snake`, `hierChild`, `sp`, `conn` and applying constraints — an **L-sized** effort per the roadmap), we propose a **two-tier strategy** that maximizes ROI:

- **Tier 1: Drawing-part shape extraction** (moderate enhancement of existing code) — reads pre-rendered shapes from `drawing1.xml` and emits them as Typst primitives. This covers **ALL diagram types that have drawing parts**, which is the vast majority.
- **Tier 2: DataModel-driven layout** (full implementation) — read the data model + layout definition + styles/colors and re-layout from scratch. Reserved for diagram types where the drawing part is insufficient or the fidelity gain justifies the cost.

### 4.2 Tier 1: Shape-Tree Extraction (Drawing Part)

**What it does:** Enhance `ConvertDiagramGraphicFrame` to read `<dsp:sp>` elements from the drawing part and emit more than just text:

1. **Read shape geometry** from `<dsp:spPr>/<a:prstGeom>`:
   - Map `prst` values to Typst shapes: `roundRect` → `#rect`, `rightArrow` → custom polygon, `rect` → `#rect`, `ellipse` → `#circle` or `#ellipse`, etc.
   - Read `<a:xfrm>` for position, size, and rotation
   - Read `<a:adjLst>` for shape-specific adjustments (e.g., corner radius on `roundRect`)

2. **Read fills** from `<dsp:spPr>`:
   - `<a:solidFill>` → extract `<a:schemeClr val="…">` or `<a:srgbClr val="…">`
   - Resolve scheme colors through the slide's theme (add a scheme-color resolver)

3. **Read strokes** from `<dsp:spPr>/<a:ln>`:
   - Stroke width, color, dash style

4. **Read text** (already done) — enhance with text formatting:
   - `<dsp:style>/<a:fontRef idx="minor">` → font choice
   - `<a:rPr sz="2300">` → font size

5. **Connector shapes** — include shapes that have no text (currently skipped)

**Estimated effort:** **M** (1 week)
**Coverage:** ~95% of the 159 diagrams in the showeet corpus would benefit
**Regression risk:** Low — additive code, no change to existing shape/text paths
**Dual-emit discipline:** Create parity fixtures with the enhanced extraction → OOXML shapes + Typst shapes should match

**Files touched:**
- **New:** `PptxEditor.Core/Converters/SmartArt/DrawingPartShapeExtractor.cs` — extraction module not in the main converter
- **New:** `PptxEditor.Core/Converters/SmartArt/DiagramSchemeColorResolver.cs` — scheme color resolution helper
- **Edit:** `PptxToTypstConverter.cs` — minimal hook (2-3 lines) to call the extractor
- **New:** `DocxEditor.Tests/Unit/SmartArt/DrawingPartShapeExtractorTests.cs` — unit tests
- **New:** `PptxEditor.Core/Generation/Fixtures/smartart-process-parity.json` — parity fixture

### 4.3 Tier 2: DataModel-Driven Layout Expansion

**What it does:** Read `dgm:dataModel`, `dgm:layoutDef`, `dgm:styleDef`, `dgm:colorsDef` and run layout algorithms to generate shapes:

1. Parse the data model tree (nodes, connections, transitions)
2. Parse the layout definition (`<dgm:layoutNode hierarchy>`)
3. Implement layout algorithms: `snake`, `sp`, `hierChild`, `conn`, `tx`
4. Apply constraints (`<dgm:constrLst>`) and rules (`<dgm:ruleLst>`)
5. Resolve presentation overrides (`<dgm:presOf>`) to apply styles/colors
6. Emit Typst shapes

**Estimated effort:** **L** (multi-week, per roadmap)
**Coverage:** Required for edge cases where drawing part is missing or when layout fidelity matters
**Dual-emit discipline:** OOXML path emits native `dsp:spTree` → PowerPoint re-renders; Typst path uses our layout output

**Recommended top-N types for Tier 2 (ranked by frequency × ROI):**

| Priority | Category | Layout types | Algo | Frequency | Rationale |
|---|---|---|---|---|---|
| 1 | **process** | process1–5, hProcess, vProcess | `snake` | 42 (26%) | Most common in business decks; snake algorithm shared across process variants |
| 2 | **hierarchy** | hierarchy1–6, orgChart | `hierChild` | 13 (8%) | Org charts are visually distinctive; text-only approximation loses their essence |
| 3 | **list** | vList, hList, pList | `sp` | 33 (21%) | Space-filling layouts are complex to re-layout; but drawing parts usually exist |

### 4.4 Support Matrix Summary

| Tier | Status | Diagram types | What user sees | Warning message |
|---|---|---|---|---|
| **Tier 1 (drawing-part extraction)** | Proposed | All types with drawing part (≥90%) | Positioned shapes with fills, strokes, text. Colors approximate but may not match theme exactly. | "SmartArt diagram rendered from pre-rendered shapes; theme colors may differ from the original." |
| **Tier 2a (snake algo)** | Proposed — Phase 1 | Process-flows (process1–5, hProcess, vProcess, chevron) | Full fidelity: correct shapes, connectors, colors from data model | "SmartArt rendered with full layout fidelity." (no warning needed) |
| **Tier 2b (hierChild algo)** | Proposed — Phase 2 | Org charts, hierarchies | Full fidelity | (no warning) |
| **Tier 2c (sp algo)** | Proposed — Phase 3 | Lists, pictures, relationships | Full fidelity | (no warning) |
| **Unsupported** | Current baseline → Phase 0 doc | Remaining edge cases | Text-only approximation with fallback placeholder | "SmartArt diagram '<name>' not fully supported; rendering as positioned text. See docs/smartart-support-matrix.md." |

### 4.5 How Tiers Fit the Dual-Emit Discipline

The repo's **layout once, emit twice** discipline applies differently to conversion vs. generation:

- **For conversion (existing PPTX → Typst):** The "layout" already happened in PowerPoint. Our job is extract-and-emit. Both paths (OOXML re-emit from parts + Typst emit from extracted shapes) can share the same extraction model.
- **For generation (JSON → PPTX):** Future SmartArt generation would require a layout engine that produces both OOXML `dgm:*` XML (for PowerPoint) and Typst shapes (for preview). This is out of scope for Phase 1.

**Parity fixture strategy:**
1. Create a minimal Typst source (from extracted shapes) and render it
2. Create the equivalent OOXML (by cloning the drawing part shapes)
3. Pixel-diff both against reference PDF

### 4.6 Risk: Drawing Part May Not Exist

The drawing part (`drawing1.xml`) is an optional Microsoft-extended part (2008 namespace schema extension referenced via `<a:ext>` in `dataModel`). If absent:
- The Tier 1 approach degrades back to current behavior (text approximation + warning)
- This would trigger a need for Tier 2 (DataModel-driven layout) for those specific diagram types

### 4.7 Docs Matrix

A `docs/smartart-support-matrix.md` should list every layout type by category with:
- Layout type identifier
- Current rendering tier (A/B/C)
- Whether drawing part extraction improves it
- Whether DataModel-driven layout is planned
- Notes about any known degradation

---

## 5. Implementation Approach: Tier 1 (Shape-Tree Prototype)

### 5.1 Architecture

```
PptxEditor.Core/Converters/SmartArt/
  DrawingPartShapeExtractor.cs    — Main extraction: reads dsp:spTree → SmartArtShape model
  SmartArtShape.cs                — Model: position, geometry, fill, stroke, text, rotation
  DiagramSchemeColorResolver.cs   — Resolves scheme/HSL colors from slide theme

PptxEditor.Core/Converters/
  PptxToTypstConverter.cs         — Hook: if SmartArt, call DrawingPartShapeExtractor instead of current path

DocxEditor.Tests/Unit/SmartArt/
  DrawingPartShapeExtractorTests.cs — Unit tests with embedded XML fixtures
```

### 5.2 Key Design Decisions

1. **New files only** — no restructuring of `PptxToTypstConverter.cs`. A single-line hook replaces the `ConvertDiagramGraphicFrame` call.
2. **Extract model → emit** — extract `SmartArtShape` objects (position, geometry type, fill, stroke, text) then emit Typst from the model. This is the dual-emit bridge: the same model can feed an OOXML emitter if needed.
3. **Geometry preset map** — a `Dictionary<string, TypstShapeProducer>` mapping PPTX preset geometry names to Typst shape functions. Start with `roundRect`, `rightArrow`, `rect`, and add presets incrementally.
4. **Scheme colors** — resolve scheme→RGB via the slide's master theme. The existing `StyleResolver` may already have this capability; if not, add a lightweight resolver.

### 5.3 Prototype Scope (Optional, Post-Report)

Given the complexity of a full Tier 1 implementation (>1 week), the optional prototype would focus on:

1. `DrawingPartShapeExtractor` that reads `dsp:sp` from the drawing part
2. Shape extraction for `roundRect` and `rightArrow` presets (covers the sales_acceleration_deck case)
3. Solid fill color extraction (scheme + direct RGB)
4. Text extraction with formatting (already partially done)
5. A minimal hook in the converter that replaces the current diagram path
6. A test that verifies shapes are extracted from `sales_acceleration_deck.pptx` slide 15

---

## 6. Open Questions

1. **Does every SmartArt diagram in the wild have a drawing part?** The showeet corpus shows 159/159 have drawing parts. Microsoft Office always generates them (the 2008 namespace extension is standard since Office 2010). Edge case: third-party OOXML generators might omit them.

2. **Scheme color resolution:** The converter's `StyleResolver` is built for slide layouts. Do we need a separate path for diagram scheme colors, or can we reuse the existing theme resolver? Investigation needed.

3. **Gradient fills in SmartArt:** `simple1` quick style uses solid fills but other quick styles may use gradients. Should Tier 1 handle gradient fills, or emit a flat approximation?

4. **3D transforms:** The quick style includes `scene3d` with camera/lighting. Should Tier 1 approximate or ignore 3D transforms?

5. **Image placeholders in SmartArt:** Picture-category diagrams reference images. Should Tier 1 extract image references from drawing-part shapes?

---

## Appendix A: Key XML Namespaces

| Prefix | URI |
|---|---|
| `dgm:` | `http://schemas.openxmlformats.org/drawingml/2006/diagram` |
| `dsp:` | `http://schemas.microsoft.com/office/drawing/2008/diagram` |
| `a:` | `http://schemas.openxmlformats.org/drawingml/2006/main` |
| `r:` | `http://schemas.openxmlformats.org/officeDocument/2006/relationships` |
| `p:` | `http://schemas.openxmlformats.org/presentationml/2006/main` |

## Appendix B: File Map

| File | Lines | SmartArt role |
|---|---|---|
| `PptxToTypstConverter.cs:1986–2008` | 23 | Diagram detection + dispatch |
| `PptxToTypstConverter.cs:2096–2198` | 103 | Shape extraction from diagram parts |
| `PptxToTypstConverter.cs:2200–2204` | 5 | Diagram namespace set |
| `PptxToTypstConverter.cs:2206–2214` | 9 | Text extraction from diagram shape |
| `PptxToTypstConverter.cs:2291–2340` | 50 | Shape position extraction |
| `PptxToTypstConverter.cs:1997–2006` | 10 | Warning emission |
| `packages/` — OpenXML SDK | — | Provides `DiagramPersistLayoutPart`, `GraphicFrame` types |

---

## 7. SmartArt ⇄ JSON Vocabulary Specification

> **Status:** Spec (round 3 implementation guide) — 2026-07-24

### 7.1 JSON Model Shape

The single JSON object that represents one SmartArt diagram:

```jsonc
{
  // $schema: "https://officeeditor.ai/schemas/smartart/v1.json",
  "version": "1.0",
  "layout": "process5",                    // short-name suffix from loTypeId URN
  "layoutCategory": "process",             // process | list | cycle | hierarchy | matrix | pyramid | picture | relationship
  "quickStyle": "simple1",                 // from qsTypeId URN
  "colorStyle": "accent1_2",               // from csTypeId URN
  "nodes": [
    {
      "id": "n1",                          // stable logical id — NOT the OOXML GUID
      "text": "SUSTAIN",                    // display text
      "order": 0,                           // sequence in the layout (srcOrd from cxnLst)
      "level": 0                            // hierarchy depth (0 = direct child of root)
    }
  ],
  "edges": [
    {
      "id": "e1",
      "from": "n1",                         // source node id
      "to": "n2",                           // dest node id
      "type": "directed",                   // directed | bidirectional | parentOf
      "label": ""                           // connector text (rare)
    }
  ]
}
```

**Design decisions:**

- **`id` uses a logical naming scheme** (`n1`, `n2`, …) rather than the OOXML GUID. The GUID is opaque and non-deterministic; a stable logical id enables merge, diff, and re-generation.
- **`order` comes from `srcOrd`** on `dgm:cxn` where `srcId=doc`. It captures the layout sequence independent of the algorithm.
- **`level` = 0** for direct children of the `doc` root. For hierarchical diagrams (`hierChild` algorithm), deeper nesting is represented with `level = 1, 2, …`.
- **Edges model explicit connections** between sibling nodes (where the layout includes connector shapes). In most layout types edges are inferred from the layout algorithm; explicit edges are present only when the data model defines them (e.g., cycle diagrams, relationship diagrams).
- **Design tokens** (colours, fonts) are NOT embedded in the model — they are resolved from the document's theme at generation time, matching how the rest of the generation pipeline works (see `DesignTokens` in `FixtureCatalog.cs`).

### 7.2 Layout-Type Mapping

Each SmartArt layout type URN maps to a **layout algorithm** and a **visual template**. The mapping table below is the driver for round 3's layout engine:

| Layout suffix | Algorithm | Visual template | Nodes per row | Connectors? | Round-3 priority |
|---|---|---|---|---|---|
| `process1`–`process5` | `snake` | Horizontal process flow with arrows | 2–3 | Yes (rightArrow) | P0 |
| `chevron1`–`chevron2` | `snake` | Chevron-shaped process blocks | 3–4 | No (imbricated) | P1 |
| `hProcess3`–`hProcess11` | `snake` | Horizontal multi-row process | varies | Yes | P1 |
| `vProcess5` | `snake` | Vertical process flow | 1 | Yes | P1 |
| `hierarchy1`–`hierarchy6` | `hierChild` | Org chart with boxes + connectors | 1+ | Yes (rightArrow+dogleg) | P0 |
| `orgChart1` | `hierChild` | Horizontal org chart | varies | Yes | P1 |
| `list1`, `vList2`–`vList6` | `sp` | Vertical bullets in containers | 1 | No | P2 |
| `hList1`–`hList9` | `sp` | Horizontal arrangement | 3–5 | No | P2 |
| `cycle1`–`cycle8` | `tx` | Circular text-only | N/A | No (text-only) | P3 |
| `venn1`–`venn3` | `tx` | Overlapping-circles text | N/A | No (text-only) | P3 |
| `pyramid1`–`pyramid4` | `sp` | Stacked segments | 1 | No | P2 |
| `matrix1`–`matrix3` | `tx` | Grid text | N/A | No (text-only) | P3 |

**Algorithm key:**
- **`snake`** — multi-row flow with `grDir`, `flowDir`, `contDir` controls. Nodes laid out row-major; connectors between sequential nodes.
- **`hierChild`** — parent–child tree with offsets. Each node positioned relative to parent, with transition shapes (parTrans/sibTrans) computing connector routes.
- **`sp`** — space-filling. Nodes tile the available container space. Tile direction depends on the layout variant.
- **`tx`** — text-only. No shapes generated; text is placed at computed positions.

### 7.3 WRITE Half Design (Round 3)

The WRITE path takes the JSON model above and produces:
1. **OOXML diagram parts** (`dgm:dataModel`, `dgm:layoutDef`, `dgm:styleDef`, `dgm:colorsDef`)
2. **A synthesized drawing part** (`dsp:spTree`) with positioned shapes
3. **The `dgm:relIds` element** referencing all five parts in the slide's graphic frame

#### 7.3.1 Architecture

```
JSON model ──┬──► DataModelGenerator ──► dgm:data1.xml
             ├──► LayoutDefGenerator ──► dgm:layout1.xml
             ├──► StyleDefGenerator ───► dgm:quickStyle1.xml
             ├──► ColorsDefGenerator ──► dgm:colors1.xml
             └──► DrawingPartGenerator ──┬──► drawing1.xml (OOXML shapes)
                                         └──► #place shapes (Typst source)
```

#### 7.3.2 Layout Engine (small, pure C#)

The layout engine for the WRITE half has a reduced scope compared to the full OOXML layout engine:

1. **Input:** Node list with text, order, level, and a layout algorithm choice.
2. **Output:** Positioned shape descriptors (x, y, w, h, rotation, geometry preset, fill, stroke, text).

**Implemented algorithms:**
- **`snake` engine** — Given `flowDir` (row/col), `contDir` (sameDir/reverseDir), `breakCount` (nodes per row), positions nodes in a grid and routes right-angle connector arrows between sequential nodes.
- **`hierChild` engine** — Given parent–child edges, stacks children horizontally below each parent, routes dogleg connectors.
- **`simpleSp` engine** — Tiles nodes into a fixed aspect ratio container with configurable padding and direction.

These are **pure C# layout** (no OOXML constraint resolution) — they compute positions directly from the JSON parameters. This matches the repo's "layout once, emit twice" discipline:
- The layout engine outputs a `List<LayoutShape>` (position, size, geometry, text).
- The OOXML emitter writes those shapes into `dsp:spTree` (drawing part) AND the diagram parts.
- The Typst emitter writes `#place(rect(…))` and `#place(line(…))` elements.

#### 7.3.3 Data Model Synthesis

The WRITE half must also **synthesize** a valid `dgm:dataModel` that references the layout/styling/colors parts:

```xml
<dgm:dataModel>
  <dgm:ptLst>
    <dgm:pt modelId="{doc-guid}" type="doc">
      <dgm:prSet loTypeId="urn:…layout/process5" loCatId="process"
                 qsTypeId="urn:…quickstyle/simple1" qsCatId="simple"
                 csTypeId="urn:…colors/accent1_2" csCatId="accent1"/>
    </dgm:pt>
    <!-- One pt per node, type="node", with prSet + text + modelId -->
  </dgm:ptLst>
  <dgm:cxnLst>
    <!-- Connections from doc → each node, srcOrd matching "order" in JSON -->
  </dgm:cxnLst>
  <dgm:extLst>
    <a:ext uri="{2008-diagram-ns}">
      <dsp:dataModelExt relId="rIdDrawing" minVer="…"/>
    </a:ext>
  </dgm:extLst>
</dgm:dataModel>
```

The model IDs are deterministic GUIDs derived from the logical node `id` via UUIDv5 to ensure stable round-trip.

#### 7.3.4 Style & Colour Synthesis

- **QuickStyle** — a minimal `dgm:styleDef` with `styleLbl` entries for `node0` and `node1` referencing fill/line/effect references from the theme.
- **Colours** — a `dgm:colorsDef` with `styleLbl` entries for fill/line/text colour lists using scheme colour references (`accent1`, `lt1`, etc.) with optional tint/shade transforms.
- **Theme integration** — all colours are scheme references resolved through the document's master theme.

#### 7.3.5 Dual-Emit Discipline

Every generated SmartArt produces:
1. **OOXML** — a real `.pptx` with valid `dgm:*` parts + `dsp:spTree` → opens in PowerPoint with native rendering.
2. **Typst** — a `.typ` source with `#place` blocks for each shape → renders to PDF/PNG via `TypstBridge`.

The Typst render is the **spec of record** for visual fidelity. Parity fixtures will be created per layout type (e.g., `smartart-process5-parity`) with a JSON model as input, both emitters run, and the output PDFs diffed with the same RMSE thresholds used for other generation primitives.

#### 7.3.6 Validation Rules

The import validator (READ path) checks:
- `layout` is a known suffix — reject unknown layouts
- Every `node.id` is referenced by at least one edge (or is a root child)
- `order` values are sequential and dense starting from 0
- `level` is ≤ `maxDepth` for the layout type (e.g., process layouts are always level 0)
- `text` length ≤ 200 characters (constraint from OOXML data model)

The export validator (WRITE path) checks:
- Node count matches the layout type's min/max (e.g., process5 requires ≥3 nodes)
- Hierarchical layouts must have at least one level-1 node
- Cycle layouts must have edge count = node count (closed loop)
