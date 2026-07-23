# Typst Backend & Compilation - Agent Rules

> **Load this** when working on `OfficeEditor.Core/Services/TypstCompilerService.cs`, `TypstBridge/**`, PPTX→Typst PDF/PNG/SVG output, or any code that calls the Typst compiler.

## Backend Selection (in order)

`TypstCompilerService.Compile()` (in `OfficeEditor.Core/Services/`) tries backends in this order:

1. **TypstBridge** (primary) — native wrapper, fastest, full feature set (PDF, PNG, SVG, fonts, working-dir).
2. **Typst CLI** (safety net) — only if TypstBridge fails or is unavailable.

The legacy managed-wrapper PDF fallback dependency has been removed; the backend chain is TypstBridge → Typst CLI only.

## Common Pitfalls (active)

- **Native vs CLI output may differ slightly** (e.g., font metrics, kerning). Always verify reference outputs after backend changes.
- **`CompileOptions.ProcessTimeout` is 2 minutes by default.** Large documents may need it raised.

## Output Formats

| Format | Constant | Notes |
|---|---|---|
| PDF | `OutputFormat.Pdf` | Default. Supports all backends. |
| PNG | `OutputFormat.Png` | Per-slide images at `CompileOptions.Ppi` (default 150). |
| SVG | `OutputFormat.Svg` | Per-slide vector output. |

## Verification

After any change to the compile path, run a smoke test on `examples/REF/PPTX/northwind-demo.pptx`. Diff generated PDFs against the reference PDFs in `examples/REF/PPTX/` (license-clean corpus: `sales_acceleration_deck` = primary 16-slide deck with SmartArt, `AetherLink-Glass-Shareholder-Overview`, `northwind-investor-40`, `northwind-launch-review`, `northwind-demo`).
