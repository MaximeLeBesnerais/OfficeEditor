# Declarative DOCX generation

OfficeEditor can generate a new `.docx` package from a loudly validated JSON document. This vocabulary is separate from the DOCX edit-instruction format: generation describes a complete document, while edit instructions mutate an existing document.

The current vocabulary version is `1.0`. See the [self-contained comprehensive example](../examples/Docx/generation/comprehensive.json).

## Generate a document

```bash
officeeditor generate examples/Docx/generation/comprehensive.json \
  --output report.docx
```

The `.docx` output extension selects DOCX generation. Always pass `--output <name>.docx`; without it, the unified command currently defaults to a `.pptx` output path.

## Document envelope and validation

A minimal document contains a version and at least one section with at least one flow block:

```json
{
  "version": "1.0",
  "sections": [
    {
      "blocks": [
        { "type": "paragraph", "text": "Hello from OfficeEditor." }
      ]
    }
  ]
}
```

The root properties are:

| Property | Required | Description |
|---|---:|---|
| `version` | Yes | Must be the string `"1.0"`. |
| `metadata` | No | Core document properties. |
| `design` | No | Palette, fonts, typography, spacing, shape, and page defaults. |
| `template` | No | Local DOCX template path. Omit it to start from a blank package. |
| `sections` | Yes | Non-empty array of section objects. |

Validation is strict and loud:

- Unknown properties are errors at every vocabulary level. Common misspellings receive suggestions.
- Required fields, JSON types, numeric ranges, enum values, token references, rectangular tables, crop bounds, and page geometry are checked.
- The validator collects errors in document order and reports each with a JSON path instead of stopping at the first contract error.
- An invalid result has no parsed document. `Parse` and all `DocxGenerator` JSON overloads throw `DocxGenerationValidationException`, whose `Issues` property contains all validation errors.
- Warnings do not invalidate a document. They are returned separately and are also included in the generation result.
- Property names, block type names, and enum values are matched case-insensitively. Palette, typography, and spacing token names are case-sensitive; use the camel-case spellings shown in this guide for portable input.

Use the validator directly when an application needs to display findings before generation:

```csharp
using DocxEditor.Core.Generation.Schema;

var validation = new DocxGenerationDocumentParser().Validate(json);
foreach (var error in validation.Errors)
    Console.Error.WriteLine(error); // $.sections[0]...: message + optional suggestion

if (!validation.IsValid)
    return;
```

## Metadata and templates

`metadata` accepts string properties `title`, `author`, `subject`, `keywords`, `description`, and `language`. They are written to the DOCX package core properties; `language` is typically a tag such as `en-US`.

`template` identifies an existing local `.docx`. The generator copies it into memory and never modifies the template file in place. Existing body content and package parts remain in the generated document, and generated sections are appended. The template's trailing body-level section properties are replaced by those of the final generated section.

Template path handling depends on the entry point:

- The CLI resolves a relative `template` against the input JSON directory, requires a local file inside that directory, rejects absolute paths, and rejects data or remote URIs.
- The public API reads the model's `template` path, or `DocxGeneratorOptions.TemplatePath`, as a filesystem path. A relative path follows the process working directory. Applications accepting untrusted input should resolve and confine template paths before calling the generator.
- `DocxGeneratorOptions.TemplatePath` overrides the JSON `template` value.

## Design tokens

If `design` is present, `design.palette` is required, though it may be empty. Palette values must be `#RRGGBB` strings.

```json
{
  "design": {
    "palette": {
      "ink": "#1F2937",
      "accent": "#2563EB",
      "paper": "#EFF6FF"
    },
    "fonts": {
      "display": "Aptos Display",
      "body": "Aptos"
    },
    "typography": {
      "reportTitle": {
        "font": "display",
        "size": 26,
        "color": "ink",
        "bold": true,
        "italic": false,
        "underline": false
      }
    },
    "spacing": {
      "small": 6,
      "section": 18
    },
    "shapes": {
      "cornerRadius": 6,
      "defaultFill": "paper",
      "defaultStroke": "accent",
      "defaultStrokeWidth": 1
    },
    "page": {
      "size": "a4",
      "orientation": "portrait",
      "margins": { "top": 72, "right": 72, "bottom": 72, "left": 72 },
      "defaultFont": "body",
      "defaultTextColor": "ink"
    }
  }
}
```

| Design object | Properties |
|---|---|
| `palette` | Arbitrary token name to `#RRGGBB` mappings. Content may reference a token or use a raw hex value; raw hex outside the palette is accepted with a warning. |
| `fonts` | `display` and `body` family names. A `font` field may reference either slot or contain a raw family name. An undefined slot falls back to the built-in font with a warning. |
| `typography.<name>` | Optional `font`, positive `size` in points, `color`, `bold`, `italic`, and `underline`. Text uses it through `token`. |
| `spacing.<name>` | Non-negative point value. Paragraph `spacing.before` and `spacing.after` may use the name. |
| `shapes` | Non-negative `cornerRadius`, `defaultFill`, `defaultStroke`, and positive `defaultStrokeWidth`. Fill defaults positioned text boxes and rectangles; stroke also defaults lines; corner radius also defaults positioned callouts. |
| `page` | Named `size`, `orientation`, `margins`, `defaultFont`, and `defaultTextColor`. Geometry applies when a section omits it. The text defaults feed positioned text and generated direct-formatting fallbacks; use `fonts.body` or typography tokens for explicit flow typography. |

Built-in defaults are A4 portrait, 72 pt (1 inch) margins, Calibri 11 pt body text, Calibri Light display text, and square shapes. A section-level page setting overrides the corresponding design page default.

## Sections and page setup

Each section accepts:

| Property | Required | Description |
|---|---:|---|
| `pageSetup` | No | Page size, orientation, margins, columns, and incoming section break. |
| `header` | No | Array of flow blocks repeated in the default header. |
| `footer` | No | Array of flow blocks repeated in the default footer. |
| `blocks` | Yes | Non-empty array of body flow blocks. |
| `positioned` | No | Array of floating primitives anchored in the section. |

`pageSetup` accepts:

- `size`: one of `a3`, `a4`, `a5`, `b4`, `b5`, `letter`, `legal`, `executive`, `statement`, or `tabloid`; or a custom object such as `{ "width": 612, "height": 792 }`. Custom dimensions are upright/portrait values and should include both positive values.
- `orientation`: `portrait` or `landscape`. Landscape swaps the upright dimensions; do not pre-swap a custom size.
- `margins`: `top`, `right`, `bottom`, and `left`, all non-negative points. Omitting the whole object inherits the design or 72 pt defaults. Within an explicit margins object, an omitted edge becomes `0`, so specifying all four edges is recommended.
- `columns`: `count` (integer at least 1), non-negative `spacing` in points, and boolean `separator`. Only counts greater than one emit a column layout.
- `breakType`: `nextPage`, `continuous`, `oddPage`, or `evenPage`. It starts this section; it is ignored with a warning on the first section. Later sections default to `nextPage`.

Margins and column gutters must fit inside the resolved page. Header and footer distance is currently fixed at 36 pt, and the vocabulary does not distinguish first, even, and default headers or footers.

## Text, runs, and spacing

Text-bearing objects use exactly one of:

```json
{ "text": "One uniformly formatted string", "token": "body" }
```

```json
{
  "runs": [
    { "text": "Mixed ", "font": "body", "size": 11 },
    { "text": "formatting", "bold": true, "color": "accent" }
  ],
  "token": "body",
  "alignment": "left",
  "spacing": { "before": 0, "after": "small", "line": 1.15 }
}
```

Text properties are:

- `token`: a name from `design.typography`.
- `alignment`: `left`, `center`, `right`, or `justify`.
- `spacing.before` and `spacing.after`: a non-negative point number or a name from `design.spacing`.
- `spacing.line`: a positive line-spacing multiple (`1`, `1.5`, `2`, and so on).

Each run requires a non-empty `text` string and may have `style`, `font`, positive `size`, `color`, `bold`, `italic`, and `underline`. Run font, size, and color override the typography token; true emphasis flags add to the token's emphasis and cannot turn a true token flag off. Leading and trailing whitespace is preserved. Newlines in positioned text become line breaks within its single paragraph.

Not every text-bearing block exposes every paragraph property: headings accept `token` and `alignment` but not `spacing`; flow callouts accept `token` but not `alignment` or `spacing`; table cells accept `token` and `alignment` but not `spacing`.

## Flow blocks

Flow blocks appear in `blocks`, `header`, `footer`, or a `group.blocks` array. They are emitted in order.

| `type` | Accepted properties | Behavior |
|---|---|---|
| `paragraph` | `text` or `runs`; optional `style`, `token`, `alignment`, `spacing` | A normal paragraph. |
| `heading` | `text` or `runs`; optional `level` (1–6, default 1), `style`, `token`, `alignment` | A heading paragraph with an outline level. |
| `list` | `items`; optional `kind`, `start`, `style` | Single-level `bullet` or decimal `ordered` list. `start` is at least 1 and only affects ordered lists. Items are strings or text objects with `text`/`runs`, `token`, `alignment`, and `spacing`. |
| `table` | `rows`; optional `widths`, `style`, `alignment` | Rectangular table. `widths` is one positive point value per column. |
| `image` | `src`; optional `fit`, `crop`, `alt`, `width`, `height`, `style` | Inline image in its own paragraph. |
| `callout` | `text` or `runs`; optional `tone`, `style`, `token` | Shaded, bordered flow callout. |
| `pageBreak` | No other properties | Hard page break. |
| `group` | `blocks`; optional `style` | Authoring container flattened into the parent flow. Nested blocks must be non-empty. |

List `kind` defaults to `bullet`; `start` on a bullet list is ignored with a warning. Each list receives fresh numbering, so ordered lists restart independently and never collide with template numbering.

A table row has a required non-empty `cells` array and optional boolean `header`; a header row repeats in Word. Every row must have the same cell count. A cell may be blank or contain `text`/`runs`, `token`, `alignment`, and `fill`. `fill` is a palette token or hex color. The generator does not support merged cells.

Callout `tone` is `note`, `tip`, `warning`, or `error` and defaults to `note`. Flow callouts use fixed tone colors. Use a positioned callout when explicit fill and stroke are needed.

## Semantic report archetypes

Beyond the raw flow blocks, the vocabulary ships five semantic report archetypes. They are authoring sugar: a pure expansion stage (`DocxGenerationExpander`) lowers them into the concrete flow blocks above (paragraphs, headings, tables, page breaks) before emission, using the theme's semantic roles and editorial surfaces — no positioned composition and no second layout engine. A document without archetypes passes through expansion unchanged, and `DocxGenerator` invokes expansion exactly once, after parse/validation and before design resolution/emission.

| `type` | Accepted properties | Expands to |
|---|---|---|
| `cover` | `title` (required); optional `eyebrow`, `subtitle`, `metadata`, `kpis`, `pageBreak` | Eyebrow/title/subtitle/metadata paragraphs on the editorial roles, an optional KPI band, and an optional trailing page break. |
| `kpiRow` | `items` (required, 2–4) | A single pale KPI band table (value row over label row). |
| `section` | `title` (required); optional `intro`, `blocks` | A level-1 heading, the intro paragraph, then the nested flow (recursively expanded). Not a DOCX page section. |
| `comparisonTable` | `columns` (required), `rows` (required); optional `emphasisFirstColumn` | A table with a repeating header row of column labels and body rows of cells. |
| `roadmap` | `phases` (required) | A Phase / Window / Action / Evidence table with tone-tinted phase numbers. |

Example cover:

```json
{
  "type": "cover",
  "eyebrow": "Q3 2026 · Editorial Edition",
  "title": "State of the Product",
  "subtitle": "A quarterly review, written for the whole team.",
  "metadata": "Prepared by the Platform Group · Reviewed 3 August 2026",
  "kpis": [
    { "value": "12.4k", "label": "Active workspaces", "tone": "positive" },
    { "value": "4.2", "label": "Incidents per month", "tone": "negative" }
  ],
  "pageBreak": true
}
```

Text fields (`title`, `eyebrow`, `subtitle`, `metadata`, `intro`, KPI `value`/`label`, roadmap `window`/`action`/`evidence`, `columns`, and comparison cells) accept either a string or a text object with `text`/`runs`, `token`, `role`, `alignment`, and `spacing`. KPI and roadmap items take a `tone` of `positive`, `neutral`, or `negative`; positive tints toward the theme teal, negative toward the theme coral, and neutral keeps the role default.

Component budgets are advisory: an item or row count outside a budget warns with its JSON path but still renders.

| Component | Budget | Warning when |
|---|---|---|
| `cover.kpis` / `kpiRow.items` | 2–4 items | 1 item, or more than 4 |
| `comparisonTable.columns` | 2–6 | fewer than 2, or more than 6 |
| `comparisonTable.rows` | 1–12 | more than 12 |
| `roadmap.phases` | 1–6 | more than 6 |

See the full report example at [`examples/Docx/generation/editorial-report.json`](../examples/Docx/generation/editorial-report.json), which exercises every archetype.

## Positioned primitives

Positioned elements are section-scoped floating objects. Geometry fields are flattened onto each element; there is no nested `position` JSON object.

Common properties are:

| Property | Description |
|---|---|
| `x`, `y` | Required non-negative offsets in points from the anchor reference. |
| `width`, `height` | Positive point dimensions. Both are required for boxes and images. A line requires only `width`, which is its length; `height` is rejected. |
| `rotation` | Clockwise degrees; default `0`. |
| `zOrder` | Integer paint order; larger values render on top. |
| `anchor` | `page`, `margin` (default), `column`, `paragraph`, or `character`. |
| `wrap` | `none`, `square` (default), `tight`, `through`, `topAndBottom`, `behindText`, or `inFrontOfText`. |
| `wrapDistances` | Non-negative point values `top`, `left`, `bottom`, and `right`. |
| `alt` | Accessibility description. |

Supported types are:

| `type` | Additional properties |
|---|---|
| `textBox` | `text` or `runs`; optional `token`, `alignment`, `spacing`, `fill`, `stroke`, `cornerRadius`. |
| `image` | Required `src`; optional `fit` and `crop`. |
| `rect` | Optional `fill`, `stroke`, and `cornerRadius`. |
| `line` | Optional `orientation` (`horizontal` or `vertical`) and `stroke`. |
| `callout` | `text` or `runs`; optional `tone`, `token`, `alignment`, `spacing`, `fill`, `stroke`, and `cornerRadius`. |

`stroke` is either a color string or `{ "color": "accent", "width": 1.5 }`. Colors may be palette tokens or `#RRGGBB`. A line with no explicit or design-default stroke uses a visible black 1 pt stroke.

Word has no exact two-axis mapping for `column`, `paragraph`, and `character` anchors, so the emitter uses the nearest WordprocessingDrawing reference and returns a warning. `tight` and `through` use the object's rectangular bounds rather than a custom contour.

## Images and assets

Both flow and positioned images use `src`. Accepted sources are:

- A `data:` URI, base64 or percent encoded.
- A local filesystem path allowed by `ImageSourceOptions`.

HTTP and HTTPS are rejected, and generation never fetches network assets. A `file:` URI is treated as an absolute local path and must pass the same absolute-path and root policy; other URI schemes are rejected. Supported payloads are PNG, JPEG, GIF, BMP, TIFF, and SVG. The payload bytes are sniffed, so a filename extension or declared data-URI media type is only a hint.

For the public API, relative local paths resolve against `ImageSourceOptions.AllowedRoot` when set, otherwise against the process working directory. Absolute paths are rejected unless `AllowAbsolutePaths` is true. When a root is set, lexical and existing-path, symlink-aware canonical checks keep assets inside it. The CLI sets the input JSON directory as the allowed root and leaves absolute paths disabled.

Default asset limits are 25 MiB encoded and decoded, 16,384 pixels on either axis, and 268,435,456 total pixels. `DocxGeneratorOptions.ImageAssetOptions` can tighten or override these values and can restrict media types. Oversized, unreadable, unsupported, or missing images fail generation rather than being dropped.

Image properties are:

- `fit`: `fill` (default, cover and center-crop), `contain` (preserve aspect ratio inside the box), `crop` (honor the author crop; without one it behaves as fill), or `stretch` (distort to the box).
- `crop`: fractions `left`, `top`, `right`, and `bottom`, each from `0` through `1`. Opposite edges must sum to less than `1`. Flow images compose this author crop with their fit. Positioned images currently apply it only with `fit: "crop"`.
- Flow `width` and `height`: optional positive point values. Supplying one derives the other from the visible source aspect ratio; supplying neither uses the DPI-derived natural size. Images without intrinsic DPI assume 96 DPI and return a warning.
- Positioned images require both `width` and `height` as part of their common geometry.
- `alt`: available on flow images directly and on positioned images through the common positioned field.

SVG is embedded without a raster fallback and produces a warning because hosts without SVG blip support may not display it.

## Template and style preservation

Templates are copied, not edited in place. Existing style definitions, document defaults, and numbering definitions are never mutated:

- A `style` reference resolves against a template style ID or style name.
- Unknown styles produce a warning and fall back to the relevant generated baseline style where one exists.
- Missing baseline styles such as Normal, Heading 1–6, Callout, and Table Grid are added lazily with collision-free IDs. Existing definitions with conventional IDs or names are reused unchanged.
- Generated lists allocate fresh numbering IDs above existing definitions.

The current emitter accepts but does not apply `style` on flow images or groups; groups are flattened. Run `style` is applied in flow content but is not currently applied inside positioned text boxes or positioned callouts.

## `DocxGenerator` API

The high-level API accepts either JSON or a parsed `DocxGenerationDocument`. A generator instance is stateless and reusable.

```csharp
using DocxEditor.Core.Generation;
using DocxEditor.Core.Generation.Assets;

var json = await File.ReadAllTextAsync("report.json");
var options = new DocxGeneratorOptions
{
    ImageSourceOptions = new ImageSourceOptions
    {
        AllowedRoot = Path.GetFullPath("assets")
    },
    // Optional: overrides "template" in the JSON.
    // TemplatePath = Path.GetFullPath("templates/report.docx")
};

var generator = new DocxGenerator();

// Path: writes through a temporary sibling and atomically replaces the destination.
var fileResult = generator.Generate(json, "report.docx", options);

// Stream: writes the complete package and leaves the caller-owned stream open.
await using var output = File.Create("report-stream.docx");
var streamResult = generator.Generate(json, output, options);

// Bytes: Content is the complete DOCX package; Result carries the model and warnings.
var generated = generator.GenerateToBytes(json, options);
await File.WriteAllBytesAsync("report-bytes.docx", generated.Content);

foreach (var warning in generated.Result.Warnings)
    Console.Error.WriteLine(warning);
```

`DocxGenerationResult` exposes the generated `Document` model, emitted file `Outputs` for the path overload, and non-fatal `Warnings`. File and asset I/O errors, invalid templates, and asset-policy failures are exceptions.

## Current limitations

- Only vocabulary version `1.0` is accepted.
- Lists are single-level bullets or decimal numbers; there are no custom glyphs or multilevel schemes.
- Tables require a rectangular grid and do not support cell merging, row/column spans, or nested flow blocks in cells.
- Headers and footers are static default variants; there are no first/even variants or field vocabulary for page numbers.
- Flow groups do not establish styling or layout; they are flattened.
- Positioned image crops are applied only in `crop` fit mode; `fill`, `contain`, and `stretch` do not compose an explicit crop in the current positioned emitter.
- Positioned primitives are boxes, images, rectangles, straight horizontal/vertical lines, and callouts. There are no connectors, arrowheads, arbitrary paths, or custom wrap contours.
- Positioned content is attached through anchor paragraphs after the section's flow content. Word's pagination and anchor rules determine the final page placement.
- Generation emits DOCX only; rendering and conversion are separate operations.
