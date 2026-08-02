# DOCX generation vocabulary (v1.0)

Declarative **JSON → DOCX** generation. This is the flagship DOCX deliverable
(roadmap-docx.md Phase 3): a versioned vocabulary for building new documents from scratch,
distinct from the edit instruction set in `DocxEditor.Core/Serialization`
(generate = new document from JSON; edit = ops against an existing document).

For current end-to-end usage, asset policy, API examples, and implementation limits, see
[`docs/docx-generation.md`](../../docs/docx-generation.md).

The model has **two tiers**, which is what makes it a product and not a paragraph printer:

- **Flow tier** — sections → blocks (paragraph/heading/list/table/image/callout/pageBreak/group),
  headers/footers, page setup.
- **Positioned tier** — the hard DOCX quirks, first-class: free text boxes, shapes
  (rect/line/callout), floating pictures with anchor reference, wrap mode, wrap distances,
  rotation and z-order. The converter's read side (`DocxToTypstConverter`) is the
  semantics reference: anything the positioned tier emits should round-trip through it.

## Layout

| Area | Files |
|---|---|
| Model (immutable records, no OpenXML types) | `Model/` — `DocxGenerationDocument`, `Section`, `PageSetup`, `DesignTokens`, `TextModel`, `FlowBlocks`, `TableModel`, `Images`, `Positioned`, `Enums` |
| Parser + validator | `Schema/` — `DocxGenerationDocumentParser`, `DocxGenerationIssue`, `DocxGenerationValidationResult`, `DocxGenerationValidationException` |
| OOXML emission | `Emit/Ooxml/` — flow, positioned content, styles, sections, and images |
| End-to-end API and contracts | `DocxGenerator.cs`, `Contracts/` — file, stream, and byte output plus parser/emitter contracts and `DocxGenerationResult` |

The implemented pipeline is **JSON → validate/parse → model → OOXML emit → DOCX**.
`DocxGenerator` provides file, stream, and byte output, using `DocxOoxmlEmitter` for the
package. Every block/primitive is documented below with its JSON `type` and C# record.

## Document root

```json
{
  "version": "1.0",
  "metadata": { "title": "…", "author": "…", "subject": "…", "keywords": "…", "description": "…", "language": "en-US" },
  "design": { "palette": { … }, "fonts": { … }, "typography": { … }, "spacing": { … }, "shapes": { … }, "page": { … } },
  "template": "path/to/template.docx",
  "sections": [ … ]
}
```

- `version` — **required**, must be `"1.0"`.
- `metadata` — optional `DocxMetadata` (core properties).
- `design` — optional `DesignTokens`.
- `template` — optional string path; the emitter opens it for styles/defaults and must
  **never mutate existing style definitions** (AGENTS.md). Omit → blank document.
- `sections` — **required**, at least one `Section`.

## Design tokens (optional)

```json
"design": {
  "palette": { "primary": "#1F4E79", "ink": "#1A1A1A", "paper": "#FFFFFF" },
  "fonts":  { "display": "Segoe UI", "body": "Calibri" },
  "typography": {
    "hero": { "font": "display", "size": 28, "color": "primary", "bold": true },
    "body": { "font": "body", "size": 11, "color": "ink" }
  },
  "spacing": { "tight": 4, "base": 12, "loose": 24 },
  "shapes": { "cornerRadius": 6, "defaultFill": "paper", "defaultStroke": "primary", "defaultStrokeWidth": 1 },
  "page":   { "size": "a4", "orientation": "portrait", "margins": { "top": 72, "right": 72, "bottom": 72, "left": 72 }, "defaultFont": "body", "defaultTextColor": "ink" }
}
```

- `palette` — required object, token name → `#RRGGBB`. Colors everywhere else accept a
  token name (preferred) or a raw hex literal (accepted with an **off-palette warning**).
- `fonts` — `display` / `body` font slots. `font` values may be `"display"`, `"body"` or a
  raw family name; using an undefined slot is a warning.
- `typography` — named tokens referenced by `token` on text content; unknown token → error
  with a suggestion.
- `spacing` — named pt values referenced by paragraph `spacing.before/after`.
- `shapes` — defaults for positioned box primitives.
- `page` — defaults applied when a section omits its own page setup.

## Section

```json
{
  "pageSetup": { … },
  "header": [ …flow blocks… ],
  "footer": [ …flow blocks… ],
  "blocks": [ …flow blocks… ],      // required, ≥ 1
  "positioned": [ …positioned elements… ]
}
```

### Page setup

```json
"pageSetup": {
  "size": "a4",                        // named: a3|a4|a5|b4|b5|letter|legal|executive|statement|tabloid
  "size": { "width": 595.3, "height": 841.9 },  // …or custom (mutually exclusive with named)
  "orientation": "portrait",           // portrait|landscape
  "margins": { "top": 72, "right": 72, "bottom": 72, "left": 72 },
  "columns": { "count": 2, "spacing": 18, "separator": false },
  "breakType": "nextPage"              // nextPage|continuous|oddPage|evenPage
}
```

- **Sizes are always portrait (upright) dimensions.** `orientation: "landscape"` swaps
  width/height at emit time — never pre-swap.
- Geometry is validated: margins must fit the page, columns must fit the text area;
  positioned elements overflowing the page edge produce warnings.
- `breakType` on the first section is ignored (warning).
- Omitted fields resolve to `design.page` defaults, then built-in defaults
  (**A4 portrait, 72 pt / 1 in margins**).

## Flow blocks

Every flow block carries an optional `style` (paragraph/table/list/picture style ref).

| `type` | Record | Notes |
|---|---|---|
| `paragraph` | `ParagraphBlock` | `text`/`runs`, `style`, `token`, `alignment`, `spacing` |
| `heading` | `HeadingBlock` | `level` (1–6), `text`/`runs`, `style`, `token`, `alignment` |
| `list` | `ListBlock` | `kind` (bullet|ordered), `items` (strings or text objects), `start` (ordered ≥ 1), `style` |
| `table` | `TableBlock` | see Tables |
| `image` | `ImageElement` | inline image; see Images |
| `callout` | `CalloutBlock` | `tone` (note|tip|warning|error), `text`/`runs`, `style`, `token` |
| `pageBreak` | `PageBreakBlock` | marker, no properties |
| `group` | `FlowContainerBlock` | section-safe raw grouping: nested `blocks`; emitters may flatten or wrap |

```json
{ "type": "paragraph", "text": "Hello", "style": "BodyText", "token": "body",
  "alignment": "justify", "spacing": { "before": 6, "after": 12, "line": 1.15 } }
{ "type": "list", "kind": "bullet", "items": ["one", { "runs": [{ "text": "two", "bold": true }] }] }
```

`spacing.before` / `spacing.after` accept a pt number **or** a `design.spacing` token name
(resolved at parse time); `line` is a line-spacing multiple (> 0).

## Text model

- `TextModel` — exactly one of `text` (string) or `runs` (array), plus `token`
  (typography reference), `alignment`, `spacing`. Empty cells allow neither (`text`/`runs`
  both absent → blank cell).
- `Run` — `text` (required), `style` (character style), `font` (slot or family), `size`
  (pt > 0), `color`, `bold`, `italic`, `underline`.

## Tables

```json
{ "type": "table", "widths": [120, 240], "alignment": "center",
  "rows": [
    { "header": true, "cells": [
      { "text": "Name", "fill": "primary", "alignment": "center" },
      { "text": "Value" } ] },
    { "cells": [ { "text": "A" }, { "text": "1" } ] }
  ] }
```

- Rows must form a **rectangular grid** (same cell count per row) — ragged tables are
  rejected. `widths` length must match the column count.
- `header` marks a repeating header row; cells accept `fill` (token/hex) and `alignment`.

## Images

```json
{ "type": "image", "src": "…", "fit": "contain", "width": 240, "height": 120,
  "crop": { "left": 0.05, "top": 0, "right": 0.05, "bottom": 0 }, "alt": "…" }
```

- `fit` — `fill` (default) | `contain` | `crop` | `stretch`.
- `width`/`height` optional; omit for natural size. `crop` edges are **0..1 fractions**
  (not OOXML units); a crop that removes the whole image is rejected.
- **Inline** images live in `blocks`. **Floating** images go in the section's `positioned`
  tier with the same `type: "image"` plus `x`/`y`/`width`/`height` and wrap settings.

## Positioned tier

Section-scoped floating primitives: `textBox`, `image`, `rect`, `line`, `callout`.
Common geometry:

| Field | Meaning |
|---|---|
| `x`, `y` | offsets from the `anchor` reference (≥ 0) |
| `width`, `height` | box extents (required for textBox/image/rect/callout; line uses `width` only) |
| `rotation` | degrees, default 0 |
| `zOrder` | paint order, higher = on top, default 0 |
| `anchor` | `page` \| `margin` (default) \| `column` \| `paragraph` \| `character` |
| `wrap` | `none` \| `square` (default) \| `tight` \| `through` \| `topAndBottom` \| `behindText` \| `inFrontOfText` |
| `wrapDistances` | `{ top, left, bottom, right }` ≥ 0 (square/tight/through/topAndBottom) |
| `alt` | accessibility text |

```json
{ "type": "textBox", "x": 120, "y": 96, "width": 240, "height": 80,
  "text": "Callout title", "token": "hero", "fill": "paper", "stroke": { "color": "primary", "width": 2 },
  "rotation": 0, "zOrder": 1, "anchor": "margin", "wrap": "square", "wrapDistances": { "top": 6, "left": 6, "bottom": 6, "right": 6 }, "alt": "…" }
{ "type": "rect", "x": 40, "y": 40, "width": 200, "height": 100, "fill": "primary", "cornerRadius": 6 }
{ "type": "line", "x": 40, "y": 180, "width": 400, "orientation": "horizontal", "stroke": { "color": "primary", "width": 2 } }
{ "type": "callout", "tone": "warning", "x": 80, "y": 200, "width": 320, "height": 60, "text": "Careful" }
```

Fills/strokes accept a token name or hex; strokes also take `{ "color": …, "width": … }`.

## Validation behavior

`DocxGenerationDocumentParser` (`Schema/`) — stateless, reuse-safe, deterministic:

- `Validate(string)` → `DocxGenerationValidationResult` with parsed `Document`, all
  `Errors` and `Warnings`; never throws.
- `Parse(string)` → `DocxGenerationDocument` or throws the typed aggregate
  `DocxGenerationValidationException` (an `OfficeEditorException`) listing every issue.

Issues are **path-qualified** (`$.sections[0].blocks[2].spacing.before`) with optional
"Did you mean …?" suggestions. The parser collects **all** errors instead of failing on the
first, and rejects **unknown properties at every level**. Rejected loudly: malformed JSON,
wrong root kinds, missing fields, wrong value kinds, unsupported versions, invalid enums,
bad colors, non-finite/negative dimensions, invalid page/margin/column geometry, empty
sections/blocks/lists/tables, ragged tables, and negative/non-positive positioned bounds.

## Generation API and emitter contract

- `IDocxGenerationParser` — `Validate` / `Parse` / `SupportedVersion` (implemented by the
  concrete parser; wire through this interface).
- `IDocxDocumentEmitter` — target-neutral `Emit(DocxGenerationDocument)` →
  `DocxGenerationResult`; the built-in implementation is `DocxOoxmlEmitter`.
- `DocxGenerationResult` — the generated model, emitted DOCX file output when applicable,
  and non-fatal warnings.

The current generation path emits OOXML `.docx` packages only. It does not include a Typst
preview emitter; rendering and conversion are separate operations.

## v1 limits (intentional)

- Lists are single-level (bullet glyph / decimal numbering); `list.start` for ordered.
- Tables are rectangular grids; no merged cells, no per-cell borders.
- No hyperlinks, fields, footnotes/endnotes, or multi-level numbering — see
  roadmap-docx.md out-of-scope list.
- Flow `group` is authoring sugar; emitters decide flatten vs. structural container.
