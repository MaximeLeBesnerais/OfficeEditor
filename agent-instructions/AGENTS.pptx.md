# PPTX & PPTX→Typst Conversion - Agent Rules

> **Load this** when working on `PptxEditor.Core/**`, `PptxToTypstConverter.cs`, `StyleResolver.cs`, or any PPTX→Typst conversion code.

## Critical Conversion Rules

These are load-bearing — they prevent hours of debugging.

1. **Use regex on `OuterXml` for unreliable OOXML attributes** (`marL`, `indent`, `algn`, `val`, `char`, `type`, `idx`, `anchor`, `b`, `i`, `sz`). `OpenXmlElement.GetAttribute()` crashes on missing attributes.
   - Pattern: `Regex.Match(element.OuterXml, $@"\b{Regex.Escape(attr)}\s*=\s*""([^""]*)""")`
   - Canonical examples: `PptxEditor.Core/Converters/PptxToTypstConverter.cs` and `StyleResolver.cs` (36 such usages).

2. **`spcPct` values are 1/1000ths of a percent** — divide by `100000.0`. Example: `120000` = `120%` = `1.2`.

3. **Typst `par(leading:)` is additive** to Typst's default line advance, not a direct replacement for PPTX line spacing. Do not set it to 0 expecting no spacing.

4. **List items must be separate arguments** — `#enum[item1][item2]`, never `#enum[all text as one string]`.

5. **Never apply global auto-fit** — it breaks body paragraph wrapping on multi-line content.

## Font Handling

- **Carlito** is the Linux fallback for missing `Aptos`/`Calibri` (declare in `--font-path`).
- **Never override embedded PPTX fonts** with `--font-path`; combine paths via `Path.PathSeparator` (`:` on Linux, `;` on Windows).
- Extract embedded fonts from `ppt/fonts/` and pass via `--font-path`.

## Color & Style Parsing

- `PlaceholderValues` and `SchemeColorValues` parsing via SDK is unreliable → regex fallback on `OuterXml`.
- Table styles: start with header bold/white, avoid complex partial per-cell strokes initially.

## Reference Files (Visual Regression)

| File | Use |
|---|---|
| `examples/REF/PPTX/northwind-demo.pptx` + `.pdf` | Smoke test source (self-made; former third-party REF decks removed pre-public-release for licensing — replacements pending) |

Smoke tests must pass on this file after any change to the conversion pipeline. Generated outputs go in `examples/output/ref/`.

## PNG Mode

`--format png` generates per-slide images at the configured PPI (default 150). Use for pixel-diff regression vs reference rendering.
