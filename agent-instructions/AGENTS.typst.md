# Typst Backend & Compilation - Agent Rules

> **Load this** when working on `OfficeEditor.Core/Services/TypstCompilerService.cs`, `TypstBridge/**`, PPTX→Typst PDF/PNG/SVG output, or any code that calls the Typst compiler.

## Backend Selection (in order)

`TypstCompilerService.Compile()` (in `OfficeEditor.Core/Services/`) tries backends in this order:

1. **TypstBridge** (primary) — native wrapper, fastest, full feature set (PDF, PNG, SVG, fonts, working-dir).
2. **`typstsharp` (legacy PDF fallback)** — only attempted if TypstBridge fails AND output is PDF AND no `--font-path` was supplied. Probed at static init; gracefully absent if NuGet type can't load.
3. **Typst CLI** (safety net) — only if both above fail or are unavailable.

Treat `typstsharp` removal as a separate dependency cleanup, **not** part of TypstBridge work. It is still referenced from `PptxEditor.Core.csproj` and `OfficeEditor.Core.csproj`.

## Common Pitfalls (active)

- **Legacy `typstsharp` can fail on hardened Linux** (e.g., missing glibc, restricted containers). TypstBridge is the primary path for this reason; CLI is the safety net.
- **Native vs CLI/legacy output may differ slightly** (e.g., font metrics, kerning). Always verify reference outputs after backend changes.
- **Don't pass `--font-path` when relying on `typstsharp` fallback** — the legacy path doesn't support custom fonts and will be skipped silently.
- **`CompileOptions.ProcessTimeout` is 2 minutes by default.** Large documents may need it raised.

## Output Formats

| Format | Constant | Notes |
|---|---|---|
| PDF | `OutputFormat.Pdf` | Default. Supports all backends. |
| PNG | `OutputFormat.Png` | Per-slide images at `CompileOptions.Ppi` (default 150). |
| SVG | `OutputFormat.Svg` | Per-slide vector output. |

## Verification

After any change to the compile path, run a smoke test on `examples/REF/PPTX/pres-pro.pptx` and `AetherLink-Glass-Shareholder-Overview.pptx`. Diff generated PDFs against the reference PDFs in `examples/REF/PPTX/`.
