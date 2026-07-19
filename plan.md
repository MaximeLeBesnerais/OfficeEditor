# Phase 5 — Slide Layout Engine & Generation Vocabulary

> Status: approved direction, pre-implementation.
> Context: docs/PROJECT-CONTEXT.md, agent-instructions/AGENTS.pptx.md, phase-4 report.
> This plan covers the "from-scratch beautiful PPTX" path (Path B): a semantic
> slide model + layout engine + dual emitters (OOXML for delivery, Typst for
> exact preview), driven by a strict JSON vocabulary for AI/human authors.

## 1. Goals & non-goals

### Goals
- Generate beautiful PPTX **from scratch** via a declarative JSON vocabulary —
  no template file required.
- **Exact preview by construction**: the generator only emits what the Typst
  emitter renders faithfully; layout is resolved once, shared by both emitters.
- AI-consumable: strict JSON Schema, loud validation errors, vocabulary small
  enough to live entirely in a prompt (~12 core concepts).
- Design tokens (palette, fonts, radii, metrics) as first-class input, minable
  from existing decks via `BrandProfileExtractor`.

### Non-goals (v1)
- No percentages — `pt` only.
- No `wrap` — fixed canvas; overflow is an error/warning, never a reflow.
- No `z-index` — paint order = document order.
- No CSS/HTML as the input format (syntax is a plugin; maybe later as a thin
  front-end compiling to the same JSON).
- No Tier-3 rendering (radial gradients, pattern fills, text-on-path, glow, 3D).
- No runtime font adaptation (that's the bakery; separate track).

## 2. Architecture

```
author (AI / human / app)
   │  JSON document (design tokens + container tree) — validated by JSON Schema
   ▼
Semantic model (C# POCOs)  ── PptxEditor.Core/Generation/Model/
   ▼
LayoutResolver (pure: tree+tokens+font metrics → absolute rects)
   │  resolves row/column/grid, justify/align, grow/aspect, text measurement
   ▼
Absolute draw tree (every element: x, y, w, h in pt, resolved fills/fonts)
   ▼                    ▼
OOXML emitter        Typst emitter (#place only — Typst does NO layout)
(delivery PPTX)      (preview PNG/SVG via TypstBridge)
   ▼                    ▼
   └──── parity harness: visual-diff per primitive fixture ────┘
```

### Load-bearing rules
1. **Layout is resolved exactly once, in C#.** Typst is a dumb renderer
   (`#place` + absolute boxes). Never let Typst's own grid/align lay out
   content — parity would double in cost.
2. **Every primitive ships with both emitters + one parity fixture in the same
   PR.** Untested parity = primitive does not exist.
3. **Coordinates are the escape hatch, not the API.** Children inside a layout
   container take no x/y. `.at(x, y)` exists only on layout-less containers.
4. **Preview tricks never enter the PPTX** (font substitution, shrink
   compensation). The PPTX is native OOXML throughout.
5. AGENTS.pptx.md rules apply to all OOXML emission (regex attribute reads,
   spcPct units, no global autofit, per-shape normAutofit only).

## 3. Vocabulary specification

### 3.1 Design tokens (document root)

```json
"design": {
  "palette":  { "primary": "#0B3D91", "accent": "#FF6B00", "ink": "#1A1A1A",
                "paper": "#FFFFFF", "muted": "#8A94A6" },
  "fonts":    { "display": "Aptos Display", "body": "Aptos" },
  "shape":    { "cornerRadius": 0, "cardStyle": "flat" },
  "metrics":  { "marginPt": 43, "gutterPt": 18, "titleSizePt": 30, "bodySizePt": 14 }
}
```

- All colors referenced by token name in content; raw hex allowed but warned
  (off-token drift is how decks get ugly).
- Token sets are seeded by running `BrandProfileExtractor` on reference decks
  (pres-pro, REMOVED, AetherLink, FusionFest) and hand-tuned.

### 3.2 Containers (layout)

```
container                       // root = the slide (fixed w/h from slide size)
  .row(gap) | .column(gap) | .grid(cols, gap, rowGap?, columnGap?)
  .justify(start|center|end|space-between|space-evenly)
  .align(start|center|end|stretch)
  .padding(all | [v, h] | [top, right, bottom, left])

child size constraints:
  .fixed(w, h) | .grow(n) | .aspect("16:9") | .alignSelf(...)

escape hatch:
  .at(x, y)                     // absolute placement, only on layout-less parents

sugar (identical code path, exists for AI ergonomics):
  spreadH(items, container)     // row + space-evenly
  spreadV(items, container)
  centerIn(item, container)
```

Semantics:
- `.grow(n)` = share of *remaining* space on the layout axis after fixed
  children are subtracted.
- `.aspect()` resolves against whichever dimension the parent constrains first.
- Grid: explicit `cols`; rows derived from child count; no auto-placement
  beyond document order.
- Overflow policy per container: `"overflow": "error" | "shrink" | "clip"`.
  Defaults: `shrink` for text-bearing leaves, `error` for layout containers.
  `shrink` uses TextFitService (fontScale ≥ MinScale, else warning).

### 3.3 Rendering primitives (Tier 1 + gradients)

| Primitive | Notes for emitters |
|---|---|
| `text` | box + anchor + align + insets; `textAlign` lives HERE (never confuse with container `.align`); runs with bold/italic/color/size |
| `rect` | per-corner radius (OOXML `round1Rect`/`round2SameRect` adj values ↔ Typst `rect(radius: (top-left: …))`) |
| `line` / `connector` | straight only in v1 |
| `ellipse` | |
| `image` | fit modes from F7 (`fill`/`crop`/`contain`) |
| `group` | transform propagation; paint order = document order |
| `gradient` | **linear only**: multi-stop, angle, alpha (Typst `gradient.linear` ↔ `a:lin`); shadow via faked offset rect in Typst, native `a:effectLst/outerShdw` in OOXML |

Tier 2 (fast follow, separate PRs): freeform polygons (`custGeom` ↔ Typst
`curve`), real blur shadows.

### 3.4 JSON shape (what the AI emits)

```json
{
  "version": "2.0",
  "design": { "...tokens..." },
  "slides": [
    {
      "type": "container",
      "layout": { "mode": "row", "gap": 18, "justify": "space-evenly", "align": "center" },
      "children": [
        { "type": "card", "size": { "grow": 1, "aspect": "4:3" },
          "content": { "title": "+34%", "subtitle": "Revenue" } },
        { "type": "card", "size": { "grow": 1, "aspect": "4:3" },
          "content": { "title": "12k", "subtitle": "Users" } }
      ]
    }
  ]
}
```

- Two surfaces, one model: fluent C# API for humans-in-code; plain JSON for AI.
- JSON Schema is the contract: unknown property → loud error with path and
  suggestion ("unknown property 'lable'. Did you mean 'label'?").
- Structured-outputs/tool-calling consumers get the schema directly.

## 4. Components (prebuilt, v1 set)

Components are C# functions over primitives — NOT a second layout system:
`card`, `kpi`, `title_block`, `bullet_list`, `divider`, `badge`, `image_card`,
`table_block` (delegates to existing table emission). Max 8 in v1. Each has a
token-driven style and an overflow contract.

Archetype slide functions (cover, section, kpi_row, two_col, table_slide) are
thin compositions of components — added AFTER primitives + components are
parity-green.

## 5. Testing & parity harness

- **LayoutResolver**: pure module, zero OOXML/Typst imports. Golden-file tests:
  tree in → expected absolute rects out. Cover: grow distribution, justify
  modes, aspect resolution, nested containers, padding, overflow=shrink with
  real font metrics, overflow=error paths.
- **Parity fixtures**: one generated deck per primitive (rect radii matrix,
  gradient angles/stops, text anchors, image fit modes…). Render via
  PowerPoint-PDF (ground truth) and Typst-SVG; visual-diff RMSE per fixture,
  committed baselines, thresholds per primitive.
- **The Typst render is the spec of record** for ambiguous cases: OOXML is
  matched TO the preview, not vice versa.
- Smoke gate unchanged: both REF decks green; generated-deck fixtures join the
  visual-diff suite.

## 6. Workstreams (worktree-per-agent, strict file ownership)

| WS | Scope | Owns | Depends on |
|---|---|---|---|
| P1 | JSON Schema + model POCOs + validator (loud errors) | `Generation/Model/`, `Generation/Schema/` | — |
| P2 | LayoutResolver (core: row/column/grid/justify/align/grow/aspect/padding) | `Generation/Layout/` | P1 |
| P3 | Text measurement integration (overflow: shrink/error via TextFitService) | `Generation/Layout/TextMeasure.cs` | P2, F6 |
| P4 | OOXML emitter (primitives Tier 1 + linear gradient) | `Generation/Emit/Ooxml/` | P2 |
| P5 | Typst emitter (#place-only; gradient, radius, shadow fake) | `Generation/Emit/Typst/` | P2 |
| P6 | Components v1 (card, kpi, title_block, …) | `Generation/Components/` | P2, P4/P5 for fixtures |
| P7 | Parity harness: fixture generator + visual-diff integration + baselines | `tools/visual-diff/`, `Generation/Fixtures/` | P4+P5 |
| P8 | Token mining: BrandProfileExtractor runs on 4 REF decks → 2–3 token sets | `Generation/Design/` (data only) | F9 |
| P9 | API + MCP surface: `deck_generate` tool, POST /api/decks/generate, SVG live preview | API, `OfficeEditor.Mcp/` | P1–P5 |
| P10 | Archetype slide functions (cover, section, kpi_row, two_col) | `Generation/Archetypes/` | P6 |

Merge order: P1 → P2 → P3 → (P4 ∥ P5) → P7 → P6 → (P8 ∥ P9 ∥ P10).
Cross-file needs via /integration/*.patch as before.

## 7. Definition of done (Phase 5)

1. `deck_generate` (MCP + API): JSON in → PPTX + per-slide SVG/PNG previews out,
   warm < 500ms/deck for ≤16 slides.
2. Parity suite green: every Tier-1 primitive + linear gradient within its
   RMSE threshold vs PowerPoint ground truth.
3. A 6-slide demo deck generated from one prompt-quality JSON, visually
   reviewed next to pres-pro (family resemblance check).
4. Validator rejects malformed docs with actionable errors (tested with 20
   adversarial docs, incl. CSS-isms like `flex-wrap`, `z-index`).
5. Overflow: a kpi_row with deliberately oversized text shrinks or errors
   loudly — never silently clips.

## 8. Open questions

1. Units inside JSON numbers: bare number = pt (assumed) or require "43pt"?
2. `cardStyle: "flat" | "outline" | "shadow"` — is 3 enough for v1?
3. Per-corner radius in JSON: object `{tl, tr, br, bl}` or CSS-like shorthand
   `[8, 8, 0, 0]`? (Leaning object — self-documenting for the AI.)
4. Slide size: 16:9 only in v1, or accept 4:3 at generation time?
5. Do archetype functions get exposed in the JSON as `{"type": "kpi_row"}` or
   stay C#-only sugar? (Leaning: expose — the AI thinks in slide types.)
```
