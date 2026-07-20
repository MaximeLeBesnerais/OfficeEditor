---
name: pptx-generator
description: Use when the user asks to generate a PPTX, create a presentation, make a slide deck, or "build some slides". Teaches the OfficeEditor JSON vocabulary and generation CLI. Use this skill before writing any generation JSON to ensure correct syntax.
---

# PPTX Generator

Generate beautiful PowerPoint decks from a declarative JSON vocabulary. No template required — the OfficeEditor engine resolves layout, components, and primitives into native OOXML.

## Trigger

- "generate a PPTX"
- "create a presentation"
- "make a slide deck"
- "build some slides about …"
- "I need a deck for …"

## How to generate

```bash
dotnet run --project OfficeEditor.Cli -- generate <deck.json> [--output <output.pptx>]
```

Validation errors are loud: path + "Did you mean …?" suggestions. 0 warnings = perfect output.

## JSON vocabulary

### Root structure

```json
{
  "version": "2.0",
  "design": { "palette": {...}, "fonts": {...}, "shape": {...}, "metrics": {...} },
  "slides": [ ... ]
}
```

### Design tokens

```json
"design": {
  "palette": { "primary": "#0B3D91", "accent": "#FF6B00", "ink": "#1A1A1A", "paper": "#FFFFFF", "muted": "#8A94A6" },
  "fonts": { "display": "Aptos Display", "body": "Aptos" },
  "shape": { "cornerRadius": 8, "cardStyle": "flat" },
  "metrics": { "marginPt": 40, "gutterPt": 16, "titleSizePt": 34, "bodySizePt": 14 }
}
```

- Palette tokens are referenced by name in fills and colors (e.g. `"fill": "primary"`).
- Raw hex is accepted but warned (off-palette drift).
- All fonts resolve from `fonts.display` / `fonts.body`. Per-run font is not set — the OOXML theme default (Calibri) applies universally, ensuring cross-OS compatibility.
- Token sets ready to use: `PptxEditor.Core/Generation/Design/*.tokens.json`.

### Slides

Each slide is either an **archetype** or a raw **container**.

## Archetypes (slide-level shortcuts)

Archetypes are pre-designed slide layouts. Use them for standard slide types.

| Type | Content fields | Purpose |
|---|---|---|
| `cover` | `kicker`, `title`, `subtitle` | Title slide |
| `section` | `index`, `kicker`, `title`, `subtitle` | Section divider |
| `kpi_row` | `title`, `subtitle?`, `kpis: [{value, label, delta}]` | 4 metric cards |
| `two_col` | `title`, `weights: [left, right]`, `left`, `right` | Side-by-side components |
| `table_slide` | `title`, `columns`, `rows`, `columnWeights?` | Data table |

Archetype example:

```json
{ "type": "kpi_row", "content": { "title": "Metrics", "kpis": [
  { "value": "+34%", "label": "Revenue", "delta": "vs LY" }
]}}
```

## Components (reusable content blocks)

Components go inside archetype slots or raw containers. 8 available:

| Component | Key fields | Notes |
|---|---|---|
| `card` | `title`, `subtitle?`, `body?` | Universal content block, token-driven style |
| `kpi` | `value`, `label`, `delta?` | Single metric |
| `title_block` | `kicker?`, `title`, `subtitle?` | Text stack |
| `bullet_list` | `title`, `items: [string]` | Titled list |
| `divider` | (none) | Horizontal/vertical rule |
| `badge` | `label` | Small colored pill |
| `image_card` | `src`, `fit?`, `title?`, `subtitle?` | Image with optional caption |
| `table_block` | `columns`, `rows`, `columnWeights?` | Inline table (for raw containers and archetype slots) |

Example:

```json
{ "type": "card", "content": { "title": "Revenue", "subtitle": "+34%", "body": "vs same quarter last year" }, "size": { "grow": 1 } }
```

## Primitives (building blocks in raw containers)

Use raw `container` slides or containers inside `two_col` slots for custom layouts.

| Primitive | Key fields |
|---|---|
| `text` | `text`, `fontSize`, `color`, `textAlign` |
| `rect` | `fill`, `radius` (number or `{tl,tr,br,bl}`), `stroke?` |
| `ellipse` | `fill`, `stroke?` |
| `line` | `orientation: "horizontal"|"vertical"`, `stroke` |
| `image` | `src` (path/URL/base64), `fit: "fill"|"crop"|"contain"` |
| `group` | `children` — overlapping shapes with `at` positions |
| `container` | `children`, `layout`, `fill`, `padding` |

### Fills

- Palette token: `"fill": "primary"`
- Hex color: `"fill": "#0B3D91"`
- Linear gradient: `"fill": { "angle": 135, "stops": [{ "color": "primary", "offset": 0 }, { "color": "accent", "offset": 1 }] }`

## Strict syntax rules

These rules are enforced by the validator — violations produce loud errors with paths.

### Size

```json
// Fixed dimensions — use w/h, NOT "fixed"
{ "size": { "w": 200, "h": 100 } }

// Grow to fill remaining space
{ "size": { "grow": 1 } }

// Aspect-ratio constraint
{ "size": { "aspect": "16:9" } }
```

### Layout

Containers with `layout` use positioning modes. Children inside layout containers **NEVER** use `at`.

```json
{
  "type": "container",
  "layout": { "mode": "row", "gap": 18, "justify": "space-evenly", "align": "center" },
  "padding": [40, 40, 40, 40],
  "children": [ ... ]
}
```

- Layout modes: `row`, `column`, `grid`
- Justify: `start`, `center`, `end`, `space-between`, `space-evenly`
- Align: `start`, `center`, `end`, `stretch`
- Padding: `[all]` or `[top, right, bottom, left]`
- Overflow: `"overflow": "clip"` on containers that overflow (default is `error` for layout containers)

### `at` is the escape hatch

`at` is only allowed on children of **layout-less** containers (no `layout` field) or inside `group`. Children inside a container with `layout` must NOT have `at`.

```json
// Layout-less parent — children need at + fixed size
{ "type": "group", "children": [
  { "type": "rect", "fill": "primary", "size": { "w": 80, "h": 80 }, "at": { "x": 0, "y": 10 } },
  { "type": "ellipse", "fill": "accent", "size": { "w": 60, "h": 60 }, "at": { "x": 100, "y": 0 } }
]}
```

### Two-col slots

`two_col` `left`/`right` slots accept **components only** (card, bullet_list, image_card, badge, divider, kpi, title_block, table_block). Raw elements (`image`, `text`, `rect`) are not valid in slots.

### Gradient stops

Offsets are **0.0–1.0** (not 0–100):
```json
{ "offset": 0 }   // 0%
{ "offset": 1 }   // 100%
{ "offset": 0.5 } // 50%
```

### Image fit modes

- `fill` — stretches to fill container (may distort)
- `crop` — center-crops, preserves aspect ratio
- `contain` — fits entire image, letterboxes

Image `src` accepts: absolute file paths, URLs, or base64 data URIs.

## Design principles

- **Cards first.** Use cards and image_cards as the primary visual language.
- **Shapes for decoration.** Use rects, ellipses, lines as accent bars, underlines, corner blobs, and background frames.
- **Palette tokens over raw hex.** Refer by name (`"primary"`, `"accent"`) unless you need custom colors.
- **Gradients sparingly.** One gradient slide per 5–8 slides is plenty. Gradient backgrounds on full-slide containers create impact.

## Reference files

Study these in the repo for real examples:

- `office-editor-full-deck.json` — 20-slide deck exercising every feature (archetypes, components, primitives, gradients, groups, image fits)
- `repo-intro-deck.json` — 8-slide overview deck (compact)
- `PptxEditor.Core/Generation/Design/*.tokens.json` — pre-mined design token sets (pres-pro, AetherLink)

## Output

1. Write the JSON deck to a `.json` file.
2. Generate with: `dotnet run --project OfficeEditor.Cli -- generate <deck.json>`
3. Report the `file://` link to the generated `.pptx`.
4. Check for 0 warnings and expected slide count.
