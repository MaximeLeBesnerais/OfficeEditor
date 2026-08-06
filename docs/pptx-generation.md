# Declarative PPTX generation

OfficeEditor can generate a complete `.pptx` deck from a loudly validated JSON document. The document is a declarative tree of **design tokens + containers + primitives + components + archetypes**; layout is resolved exactly once in C#, and the same resolved tree feeds both the OOXML emitter (the delivered `.pptx`) and the Typst emitter (PDF/PNG/SVG previews). See the sibling guide [Declarative DOCX generation](docx-generation.md) for the analogous DOCX vocabulary.

The current vocabulary version is `2.0`. The canonical artifact is the JSON schema at [`PptxEditor.Core/Generation/Schema/deck.schema.json`](../PptxEditor.Core/Generation/Schema/deck.schema.json) (`$id: https://officeeditor.dev/schemas/deck-2.0.json`), and every rule in this guide is grounded in it. Working deck examples live at [`demo/deck.json`](../demo/deck.json), [`demo/demo-deck.json`](../demo/demo-deck.json), and [`decks/repo-overview.json`](../decks/repo-overview.json).

## Generate a deck

```bash
officeeditor generate demo/deck.json --output deck.pptx
```

The `.pptx` output extension selects PPTX generation. The C# pipeline, which the CLI wraps, is:

```csharp
using PptxEditor.Core.Generation.Archetypes;
using PptxEditor.Core.Generation.Components;
using PptxEditor.Core.Generation.Emit.Ooxml;
using PptxEditor.Core.Generation.Layout;
using PptxEditor.Core.Generation.Schema;

var json = await File.ReadAllTextAsync("deck.json");            // "version": "2.0" vocabulary
var result = new GenerationDocumentParser().Validate(json);     // loud validator
if (!result.IsValid) { /* field-path errors with suggestions */ }

var doc        = ArchetypeExpander.Expand(result.Document!);    // archetypes → components
var components = ComponentExpander.Expand(doc);                 // components → primitives
var layout     = new LayoutResolver().Resolve(components);      // layout once…
var pptx       = new OoxmlEmitter().Emit(layout);               // …emit OOXML…
await File.WriteAllBytesAsync("deck.pptx", pptx.Bytes);
// …and the same resolved layout feeds the Typst emitter for PDF/PNG/SVG previews.
```

## Document envelope and validation

A minimal deck contains a version, a design block, and at least one slide:

```json
{
  "version": "2.0",
  "design": { "palette": { "paper": "#FFFFFF", "ink": "#1F2937", "accent": "#2563EB", "muted": "#6B7280" } },
  "slides": [
    {
      "type": "container",
      "layout": { "mode": "column", "gap": 18, "justify": "center" },
      "fill": "paper",
      "padding": 46,
      "children": [
        { "type": "text", "text": "Hello from OfficeEditor.", "fontSize": 30, "color": "ink" }
      ]
    }
  ]
}
```

The root properties are (`required: ["version", "design", "slides"]`):

| Property | Required | Description |
|---|---:|---|
| `version` | Yes | Must be the string `"2.0"` (`const`). |
| `slideSize` | No | `"16:9"` (default) or `"4:3"`. `"16:9"` → 960 × 540 pt canvas; `"4:3"` → 720 × 540 pt. Any other value (including square/1:1) is rejected by the enum. |
| `design` | Yes | Palette, fonts, shape, and metrics design tokens (see below). `palette` is required inside it. |
| `slides` | Yes | Non-empty array of slide roots. Each root is a `container` element or an archetype slide type. |

Validation is strict and loud:

- Unknown properties are errors at every vocabulary level (`additionalProperties: false` throughout the schema).
- Required fields, JSON types, numeric ranges, enum values, token references, aspect ratios, crop bounds, and geometry rules are checked.
- The validator collects errors and reports each with a JSON path; `Validate` returns a result whose `IsValid` is false when anything is wrong. The parse/generate entry points throw `GenerationValidationException` (whose `Issues` carries all errors) for invalid input — an invalid document has no parsed model.
- The component and archetype layers validate their payloads after parsing with the same loud style: `ComponentException` with JSON paths and suggestions.
- Warnings do not invalidate a document. The schema flags off-palette `#RRGGBB` hex literals ("off-token drift"), and the layout resolver reports non-fatal diagnostics such as text shrunk below its minimum scale. Warnings are returned separately.

## Design tokens

`design` carries the deck's theme. Content references colors by token name; tokens map to `#RRGGBB` literals.

```json
"design": {
  "palette": { "primary": "#147C7A", "accent": "#C79A3B", "ink": "#22302A", "paper": "#FFFFFF", "muted": "#8A8578" },
  "fonts": { "display": "Avenir Next", "body": "Helvetica Neue" },
  "shape": { "cornerRadius": 0, "cardStyle": "outline" },
  "metrics": { "marginPt": 46, "gutterPt": 18, "titleSizePt": 44, "bodySizePt": 16 }
}
```

| Object | Properties |
|---|---|
| `palette` | Required. Token name → `#RRGGBB` (`^#[0-9a-fA-F]{6}$`). Content references colors by token name; a raw `#RRGGBB` literal outside the palette is accepted with a warning. |
| `fonts` | Optional `display` and `body` family names. A `font` field on text may reference either slot token or a raw family name. |
| `shape` | Optional `cornerRadius` (pt, default `0`) and `cardStyle` (`flat` \| `outline` \| `shadow`, default `flat`). |
| `metrics` | Optional `marginPt` (default `43`), `gutterPt` (default `18`), `titleSizePt` (default `30`), `bodySizePt` (default `14`) — all non-negative except the text sizes, which must be positive. |

All dimensions in the vocabulary are **points (pt)**; there are no percentages.

## Slides and elements

Every slide root is a `container` element (or an archetype slide — see below). The root container fills the whole slide canvas; `size` and `at` are rejected on it. Children are laid out in **document order, which is also the paint order** (there is no z-index).

### Common child geometry

Every element accepts `size` and `at`:

- `size` — child size constraints: `w` and/or `h` (positive pt), `grow` (share of remaining space on the layout axis), `aspect` (`"W:H"`, e.g. `"16:9"`), and `alignSelf` (`start` \| `center` \| `end` \| `stretch`).
- `at` — absolute placement escape hatch `{ "x": …, "y": … }` (non-negative pt), only on children of **layout-less** parents.

Geometry rules are enforced with loud errors:

- `at` inside a layout container is rejected ("position via layout mode, grow, justify and align").
- Children of a layout-less parent (free canvas) require `at` **and** a fixed `size` (`w` + `h`, or one dimension plus `aspect`); `grow` there is rejected.
- `size`/`at` on the slide root are rejected (the root is the canvas).

### Container

```json
{
  "type": "container",
  "layout": { "mode": "column", "gap": 16, "justify": "start", "align": "stretch" },
  "padding": [46, 46, 46, 46],
  "fill": "paper",
  "stroke": { "color": "ink", "width": 1 },
  "radius": { "tl": 8, "tr": 8, "br": 0, "bl": 0 },
  "shadow": { "color": "ink", "dy": 2, "blur": 6, "alpha": 0.25 },
  "overflow": "clip",
  "children": [ … ]
}
```

| Property | Meaning |
|---|---|
| `layout` | `{ "mode": "row" \| "column" \| "grid", … }`. `gap` (default `0`); grid mode also requires `cols` and accepts `rowGap`/`columnGap` (both grid-only — rejected on row/column). `justify` (main axis): `start` \| `center` \| `end` \| `space-between` \| `space-evenly`, default `start`. `align` (cross axis): `start` \| `center` \| `end` \| `stretch`, default `stretch`. Without `layout` the container is a **free canvas** whose children must use `at` + fixed `size`. |
| `padding` | `edgeInsets`: a single number (all edges), `[v, h]`, or `[top, right, bottom, left]`; non-negative pt. |
| `overflow` | `error` (default) \| `shrink` \| `clip`. Per-container policy; text elements default to `shrink` instead. |
| `fill` | A color (token or `#RRGGBB`) or a linear `gradient` (see below). |
| `stroke` | A color or `{ "color": …, "width": 1 }` (width default `1`). |
| `radius` | A number (all corners) or `{ "tl", "tr", "br", "bl" }` per-corner pt. |
| `shadow` | `{ "color": … }` (required) plus `dx`, `dy`, `blur`, `alpha`. Native `a:effectLst/outerShdw` in the OOXML output; a faked offset rect in the Typst preview (never in the PPTX). |
| `children` | Required (may be `[]` for a blank container). Element array in document order — which is also the paint order. |
| `notes` | Speaker notes — **slide root only**; see [Speaker notes](#speaker-notes). |

### Text

```json
{
  "type": "text",
  "text": "Documents as code",
  "font": "display",
  "fontSize": 58,
  "color": "paper",
  "bold": true,
  "textAlign": "center",
  "anchor": "middle",
  "insets": [8, 16],
  "overflow": "shrink",
  "size": { "h": 74 }
}
```

Text content is exactly one of `text` (a single uniformly formatted string) or `runs` (at least one styled run):

| Property | Meaning |
|---|---|
| `text` / `runs` | One is required. A run requires its own `text` and may carry `font`, `fontSize`, `color`, `bold`, `italic`. |
| `font` | Font slot token (`display` \| `body`) or a raw family name. |
| `fontSize` | Positive pt. |
| `color` | Palette token or `#RRGGBB`. |
| `bold` / `italic` | Boolean, default `false`. |
| `textAlign` | `left` \| `center` \| `right`, default `left`. Lives on text elements, never on containers. |
| `anchor` | `top` \| `middle` \| `bottom`, default `top` (vertical anchor in the box). |
| `insets` | `edgeInsets` padding inside the text box. |
| `overflow` | Defaults to `shrink` for text. |
| `shadow` | Optional drop shadow. |
| `size` / `at` | Common geometry. |

### Rect, ellipse, line/connector, image, group

```json
{ "type": "rect", "fill": "accent", "radius": 6, "size": { "w": 200, "h": 3 } }
```

- **`rect`** — `fill`, `stroke`, per-corner `radius`, `shadow`, `size`, `at`.
- **`ellipse`** — `fill`, `stroke`, `shadow`, `size`, `at`.
- **`line` / `connector`** — straight only in v1. `orientation` (`horizontal` default \| `vertical`), `stroke`, `size`, `at`. Authored as `"type": "line"` it emits a plain shape; `"connector"` emits an OOXML `cxnSp`.
- **`image`** — `src` (required) plus `fit`, `crop`, `alt`, `size`, `at`. See [Images and assets](#images-and-assets).
- **`group`** — layout-less grouping with a required `children` array placed via `at` in group coordinates; paint order is document order. Accepts `size`/`at` only.

### Fill gradients

A `fill` may be a linear gradient instead of a color:

```json
"fill": {
  "angle": 135,
  "stops": [
    { "color": "primary", "offset": 0 },
    { "color": "accent", "offset": 1, "alpha": 0.8 }
  ]
}
```

`angle` is degrees (mapped between OOXML `a:lin` and Typst `gradient.linear`); `stops` needs at least two entries, each with `color` and `offset` (0–1) and optional `alpha` (0–1). Only linear gradients are supported.

## Components

Components are prebuilt C# functions over primitives — not a second layout system. They expand in the regular component pass and share the same layout and emitters. The v1 component set is `card`, `kpi`, `title_block`, `bullet_list`, `divider`, `badge`, `image_card`, `table_block`. A component element is:

```json
{
  "type": "card",
  "content": { "title": "Create", "body": "Decks from declarative JSON." },
  "size": { "grow": 1 }
}
```

`content` is the strongly typed payload; unknown payload keys are loud errors with JSON paths. All colors are palette tokens or `#RRGGBB`; all dimensions are points. Text overflow inside components shrinks by default; structural overflow (a surface that cannot fit) errors.

| Component | Payload |
|---|---|
| `card` | `title` (required), optional `subtitle`, optional `body` (grows into remaining space). |
| `kpi` | `value` (required), `label` (required), optional `delta` (trend line, rendered in the accent color). |
| `title_block` | `title` (required, display font at `metrics.titleSizePt`), optional `subtitle`, optional `kicker` (accent overline). |
| `bullet_list` | `items` (required, at least one, one line each), optional `title` (heading), optional `markerColor` (default `"accent"`). |
| `divider` | Straight rule: `color` (default `"muted"`), `widthPt` (default `1`), `orientation` (`horizontal` default \| `vertical`). |
| `badge` | Pill label: `text` (required), `color` (fill, default `"accent"`), `textColor` (default `"paper"`). Height comes from font metrics unless fixed; give it a width or grow on the parent axis. |
| `image_card` | `source` (required; image resolution rules as for the `image` primitive), optional `title`, `subtitle`, `alt`, `fit` (default `crop`), `imageGrow` (default `1`). |
| `table_block` | `columns` (required, ≥ 1), `rows` (required; each row exactly `columns` cells), `header` (default `true`), `columnWeights` (relative widths, one per column), `rowHeight` (pt; defaults to the body-size line estimate + cell padding). |

## Archetype slides

Archetypes are the prompt-friendly authoring surface: whole-slide functions that compose components under the hood. A slide root may be `"type": "cover"`, `"section"`, `"kpi_row"`, `"two_col"`, or `"table_slide"` instead of `"container"`. Expansion happens exactly once, after parsing and before the component pass — an archetype composes `ComponentElement` nodes that the regular component pass then expands, so there is one layout and one expansion path. An archetype slide fills the whole slide (`size`/`at` are not allowed on it), carries no surface properties of its own, and accepts only `type`, `content`, and `notes`. Archetypes never nest: `two_col` slots must be components, and archetype names are rejected at child level.

| Archetype | Payload | Composes |
|---|---|---|
| `cover` | `title` (required), optional `subtitle`, `kicker` | Vertically centered title block (kicker, display title, muted subtitle) with an accent rule beneath, on generous editorial margins. |
| `section` | `title` (required), optional `index` (e.g. `"02"`), `subtitle`, `kicker` | Bottom-anchored section divider behind an accent rule. With `index`, a big-figure KPI cell beside the title block; without, a standalone title block. |
| `kpi_row` | `title`, `subtitle` optional; `kpis` (required, 1–6 of `{ "value", "label", "delta" }`) | Optional title block over a row of equally wide KPI cards (grow 1 each) filling the remaining height. |
| `two_col` | `title`, `subtitle` optional; `left` and `right` required component slots (`{ "type": …, "content": … }`); optional `weights` (exactly two positive grow weights, default `[1, 1]`) | Optional title block over two component cells sharing the row by grow weight. The archetype assigns slot geometry — `size`/`at` on a slot are rejected. |
| `table_slide` | `title`, `subtitle` optional; `columns` (required, ≥ 1), `rows` (required), `header` (default `true`), `columnWeights`, `rowHeight` | Optional title block over a `table_block` that fills the remaining height. |

Example `two_col` slide:

```json
{
  "type": "two_col",
  "content": {
    "title": "What this demo shows",
    "weights": [3, 2],
    "left": {
      "type": "bullet_list",
      "content": { "title": "Two pipelines, one Makefile", "items": ["Render an existing deck", "Generate a new deck from JSON"] }
    },
    "right": {
      "type": "card",
      "content": { "title": "Layout once, emit twice", "body": "One C# layout pass feeds both emitters." }
    }
  }
}
```

## Speaker notes

Speaker notes are per-slide metadata, supported since the v2.0 additive schema change (commit `72d5a1d`):

```json
{
  "type": "cover",
  "content": { "title": "State of the Product" },
  "notes": "Remind the room: numbers are Q3, all regions included."
}
```

Rules, as enforced by the parser and emitters:

- `notes` is an optional string on **any slide root** — a `container` root or an archetype slide root. It is read through the same path for both.
- A `notes` property on a **nested** container is rejected with a loud validation error ("'notes' is only valid on the slide root (speaker notes are per-slide metadata).") rather than silently dropped.
- The OOXML emitter writes the notes into a real `NotesSlidePart` wired to the slide, so PowerPoint shows them in the presenter view. Multi-line notes become one paragraph per line. Blank/whitespace notes skip the part entirely.
- The Typst preview **intentionally does not render notes** — they are metadata, not visual content.
- On the fluent builder side, `ISlideBuilder.AddNotes(string)` / `ISlideBuilder.SetNotes(string)` set or replace a slide's notes (creating or updating the `NotesSlidePart`); blank input removes an existing notes part. The generation emitter and the edit path share the same `NotesSlideWriter`, so both produce byte-identical notes parts. `PptxToMarkdownConverter` reads notes back (roundtrip-tested).
- The schema stays at version `2.0`; the feature is purely additive. Known behavior: `DuplicateSlide` drops notes on the duplicate (locked by a test).

## Images and assets

The `image` primitive and the `image_card` component take a `src` (component payload key: `source`). The emitters interpret it with exact, fixed precedence (see `OoxmlEmitter.ResolveImageSource` and `ImageSourceResolver`):

1. **`data:<mime>;base64,<payload>`** — inline payloads. The media type must be one of `image/png`, `image/jpeg`, `image/gif`, `image/bmp`, `image/tiff`, `image/svg+xml`; anything else (or a payload that is not base64) throws.
2. **`http://` / `https://`** — rejected: `"Remote image URLs are not supported in v1; pass a local file path or a data URI."` Generation never fetches network assets.
3. **Absolute paths** — read as-is (`FileNotFoundException` if the file does not exist).
4. **Relative paths** — resolved against the **directory of the JSON document only** (the deck file's own directory). There is no repository-root or process-CWD fallback. Resolution is containment-checked — lexically and, for existing paths, canonically (symlink-aware): a source that escapes the document directory (`../`, or a symlink pointing outside it) is rejected with an `ArgumentException` before any file probe; a missing file inside the directory throws `FileNotFoundException`. How each surface supplies that directory: the **CLI** resolves relative `src` against the JSON document's own directory; the **API** absolutizes relative sources against the repository root (its confined base) before emitting, so relative paths keep working there too. Only the **MCP `deck_generate`** tool — a string-only surface that receives raw JSON with no document directory — rejects relative paths, with a loud, actionable error: use a data URI or an absolute path.

Supported file extensions: `.png`, `.jpg`/`.jpeg`, `.gif`, `.bmp`, `.tiff`/`.tif`, `.svg`; unknown extensions throw loudly rather than mislabeling the payload.

Image properties:

- `fit`: `fill` (default; covers the box with center cropping), `crop` (honors the authored crop; without one it behaves as `fill`), or `contain` (preserve aspect ratio inside the box). There is no `stretch` in the generation vocabulary.
- `crop`: `{ "left", "top", "right", "bottom" }` — integers from `0` to `100000`, the 1/1000ths-of-a-percent (spcPct) scale family (`100000` = 100% of that edge). All four edges are required. Only honored with `fit: "crop"`.
- `alt`: accessibility description.

## Layout-once / emit-twice

The resolved layout is a pure C# artifact (`ResolvedSlide`: absolute point geometry, colors resolved to hex, font slots resolved to family names). Both emitters consume it directly:

- **OOXML emitter** — the delivered `.pptx` (slides, shapes, notes parts, `sldSz`, and `notesSz`).
- **Typst emitter** — a `#place`-only Typst source for PDF/PNG/SVG previews. It is the preview, not a second layout engine, and the preview is the spec of record for ambiguous OOXML rendering.

A deck renders end-to-end through the same pipeline: `officeeditor generate deck.json --output deck.pdf` (or `.png`, `.svg`).

## Current limitations

- Only vocabulary version `2.0` is accepted; unknown properties are rejected everywhere.
- Slide sizes are limited to `16:9` (960 × 540 pt) and `4:3` (720 × 540 pt); square/1:1 and custom sizes are not supported.
- Layout is point-based with `row`/`column`/`grid` modes; there are no percentages, no text wrap, no z-index, and no CSS-style input.
- Lines and connectors are straight (horizontal/vertical) only.
- Gradients are linear only; shadows are a single outer drop shadow.
- Components and archetypes are the fixed v1 set; archetype slots cannot nest archetypes.
- Autofit is never emitted; text overflow is handled by the `shrink` / `error` / `clip` policies.
