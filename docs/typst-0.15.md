# Typst 0.15 — what it means for OfficeEditor

Date: 2026-07-23. Sources: typst/typst v0.15.0 release notes + https://typst.app/docs/changelog/0.15.0/ (PR numbers cited), v0.15.1 patch notes, docs.rs API pages for typst-pdf/typst-render/typst-kit 0.15.0.

**TL;DR.** 0.15 is a moderate-impact upgrade for us. The bridge upgrade is mostly mechanical but touches all three render paths: `typst-pdf` 0.15 changes `PdfOptions` (lifetime removed, new `creator`/`pretty` fields, multi-standard `PdfStandards`), `typst-render` 0.15 replaces the raw `pixel_per_pt` argument with a `RenderOptions` struct, and `typst-kit` was fully reworked — `FontSearcher` (used in `TypstBridge/native/src/fonts.rs`) is gone and the crate's feature flags were renamed. The one change that can silently shift our renders is the layout engine's expanded baseline retention (marked breaking upstream) plus new list-marker defaults — parity fixtures must be re-diffed against the REF PDFs after upgrade. Variable fonts are the headline feature: they don't remove our Aptos workaround, but they change how variable-font families are named/matched. PDF/A + PDF/UA-1 can now be combined, which is a genuine new capability we can expose through `CompileOptions`.

## 1. PDF export

- Upstream (0.15.0): "Typst can now target multiple (compatible) PDF standards at once, e.g. PDF/UA-1 and PDF/A-2a" (PR #8294). PDFs are minified by default (`--pretty` opt-out). Tagging fixes: artifact kinds in `pdf.artifact` (PR #8416), invalid structure with complex list markers (#7789), tagging order for inline content (#7861), bbox of stroked shapes (#8322). Graphics: COLRv1 compositing (krilla), gradient fixes in LinearRGB/CMYK/Luma (#8149), fixed "excessive sampling of linear gradients" (#7818).
- API diff (docs.rs): `PdfOptions<'a>` → `PdfOptions` (no lifetime); `ident: Smart<&str>` → `Smart<String>`; new fields `creator: Smart<Option<String>>` and `pretty: bool`; `standards: PdfStandards` still built via `PdfStandards::new(&[PdfStandard])` but now accepts a version + one validator together. `tagged: true` remains the default. krilla 0.6 → 0.8.2; typst-pdf now depends on a new `typst-layout` crate.
- Impact: `TypstBridge/native/src/render_pdf.rs:4` constructs `PdfOptions::default()` — compiles unchanged, but this is where new options get wired. Tagged PDF is already on for free.
- Action: **adopt (upgrade-time)** — extend the bridge ABI (`abi.rs`) + `CompileOptions` with an optional PDF-standards list (e.g. `a-2a`, `ua-1`) and pass through to `PdfStandards::new`. **Adopt (optional)**: expose `creator` metadata. **None**: `pretty` (debug only).

## 2. Layout & text

- Upstream (0.15.0): baseline info retained "in many more parts of the layout engine" (PR #8150, **breaking**) — text in an inset `box` now baseline-aligns with surrounding text; list/enum markers baseline-align with first line of item. New `list.marker-align` property; default is now baseline-vertical + `end` horizontal (#7895). Centering inside lists now uses full available width (#7895). Fixed justified text protruding into margin after e.g. zero-width space (#8415). Fixed uneven CJK-Latin spacing in justified paragraphs (#7606). Unicode update fixed guillemet line breaking (#8406). Layout convergence failures now produce detailed diagnostics (#7364).
- Upstream (0.15.1): fixed gaps in multi-page lists when `marker-align`/`number-align` has a vertical component; math-only fixes (`lr` alignment, `op` vertical alignment).
- Impact: our emitters (`PptxEditor.Core/Generation/Emit/Typst`, `PptxEditor.Core/Converters/*ToTypstConverter.cs`) lean on `#place` with precise boxes; `place`-only slides are largely unaffected by the baseline change, but any inline `box(inset: ...)` usage and list emission can shift. Expect pixel-level diffs.
- Action: **watch/verify** — after the crate bump, regenerate the parity fixtures and re-run visual-diff on the license-clean REF corpus (`examples/REF/PPTX/`: sales_acceleration_deck, AetherLink, northwind-launch-review, northwind-demo) before claiming the upgrade done. **Adopt (later)**: convergence diagnostics improve our error surfacing in `TypstCompilerService`.

## 3. Fonts

- Upstream (0.15.0): variable font support (PR #8425) — axes `ital/slnt/wght/wdth/opsz` auto-set from `text` params; custom axes via `text(variations:)`; family names trimmed of "Variable"/"Var"/"VF" suffixes (**minor breaking**, #8444). Stricter `text.features` tag parsing (minor breaking). SimSun-ExtB family-merge exception. CLI: Adobe CC font discovery, lazy font scan (CLI only).
- Upstream (dev): "The `typst-kit` crate was completely reworked" (#7710, #8026). 0.15 `typst_kit::fonts` exposes `FontStore`, `FontSource`, `FontPath`, and functions `embedded()` / `scan()` / `system()`; feature flags renamed (`embedded-fonts`, `scan-fonts` — our `fonts`/`embed-fonts` no longer exist). `FontSearcher` is removed.
- Impact: `TypstBridge/native/src/fonts.rs:8` uses `FontSearcher` for embedded fallback fonts — must be ported to `fonts::embedded()`; `Cargo.toml` feature list must change. Our Aptos→Carlito substitution (`PptxEditor.Core/Services/FontMetricsCatalog.cs`, `PptxToTypstConverter.cs:4636`) is still needed: Aptos is not embedded and nothing about fallback chains changed. Note: if a user supplies the real Aptos (a variable font) via font-path, 0.15 will instantiate it properly, but `FontMetricsCatalog` reads default-instance metrics — metric drift vs the used instance is possible.
- Action: **adopt (upgrade-time)** — rewrite `fonts.rs` against the new `typst-kit` fonts API. **Watch** — variable-font support + metrics parity in `FontMetricsCatalog` if we ever ship Aptos VF. **Check** — any emitted font name ending in Variable/Var/VF and any `features` tags.

## 4. Performance

- Upstream: no headline compile/render speedups in 0.15. Incremental wins: fixed "excessive sampling of linear gradients" in PDF and SVG export (#7818, relevant to gradient-heavy decks like AetherLink); smaller minified PDF/SVG output; deterministic cross-platform floating-point (#7712, helps fixture reproducibility); CLI lazy font discovery (CLI path only).
- 0.15.1: multi-page PNG/SVG error now raised only after the final iteration — fewer spurious failures in iterative compiles.
- Impact: ms/slide claims in demos are safe; gradient slides may export slightly faster/smaller.
- Action: **none** beyond re-measuring after upgrade.

## 5. PNG/SVG rendering

- Upstream (0.15.0): `typst-render` gains `RenderOptions { pixel_per_pt: Scalar, render_bleed: bool }` — `render()`/`render_merged()` now take options instead of a bare float (docs.rs). Deps: tiny-skia 0.12 (we pin 0.11.4 in `Cargo.toml`), resvg 0.47, pixglyph. PNG fixes: conic gradient angles (#7952), color bitmap glyph positioning/sizing (#8440), invisible negatively-scaled text (#8111). SVG: minified by default, smaller output, `typst-frame`/`typst-doc`/`typst-group`/`typst-shape`/`typst-text` classes removed (**minor breaking**, #7680), non-determinism fixed (#7680).
- Impact: `TypstBridge/native/src/render_png.rs:15` (`typst_render::render(page, pixel_per_pt)`) needs the `RenderOptions` form; PPI semantics unchanged (still `ppi/72.0` → pixel_per_pt, validated in `compiler.rs:378`). SVG class removal is irrelevant unless anything post-processes our SVGs by class.
- Action: **adopt (upgrade-time)** — port `render_png.rs`, bump tiny-skia to 0.12.x, verify `typst-svg` 0.15 entry-point signature during the bump.

## 6. Everything else notable

- New `typst-layout` crate split out (internal restructure; affects Cargo.lock only).
- MSRV raised to Rust 1.92 (was 1.89 on 0.14.x) — check CI toolchain.
- File paths in Typst source may no longer contain backslashes (**breaking**) — emitters must emit forward-slash paths for images/fonts (Windows paths passed via the bridge ABI are unaffected).
- Removals (**breaking**): `path` element (use `curve`), `pattern` (use `tiling`), `pdf.embed` (use `pdf.attach`), scoped `*.decode` functions — grep emitters before upgrade; we use none of these today.
- Bundle export (experimental) can emit PDF+PNG+SVG+assets from one compile — **watch** as a future single-pass multi-output path, not usable via the crates yet without feature flags.
- Multiple bibliographies, `divider`, `within` selector, spot colors, MathML/HTML work — no impact on our paged pipeline.
- Note: our current pin (0.14.2) already includes the wasmi security fix; plugins are unused.

## Breaking changes checklist

| Change | Affects us? | Where handled |
|---|---|---|
| `typst-pdf` `PdfOptions` API (lifetime removed, `creator`, `pretty`, multi-standard) | Yes | `TypstBridge/native/src/render_pdf.rs`, `abi.rs`, `CompileOptions` |
| `typst-render` `RenderOptions` (was bare `pixel_per_pt`) | Yes | `TypstBridge/native/src/render_png.rs:15` |
| `typst-kit` rework (`FontSearcher` removed, features renamed) | Yes | `TypstBridge/native/src/fonts.rs:8`, `Cargo.toml` |
| tiny-skia 0.11 → 0.12 | Yes | `TypstBridge/native/Cargo.toml` pin |
| Baseline retention in layout engine (render shifts) | Yes | parity fixtures in `Generation/Fixtures/`, visual-diff vs `examples/REF/` |
| New `list.marker-align` default | If lists emitted | `Emit/Typst` list emission |
| No backslashes in source paths | Check | path emission in `Emit/Typst`, converters |
| Variable/Var/VF family-name trimming | Check | font names in decks/fixtures |
| Stricter `text.features` parsing | Check | any emitted `features` args |
| Removals (`path`, `pattern`, `pdf.embed`, `*.decode`) | No (grep to confirm) | — |
| HTML `box`/`block` semantics, `html.script/style` strings-only | No (no HTML output) | — |
| SVG `typst-*` classes removed, minified SVG | No (no class post-processing) | — |
| MSRV 1.92 | Yes | Rust toolchain in CI/build |
