> Note: .pptx and the other third-party REF decks referenced below were removed (licensing). Historical numbers retained for reference.

# PPTX Render Benchmark — Baseline (Pre-Optimization)

Generated: 2026-07-18 11:38:31 UTC by `tools/pptx-benchmark`

Ground-truth latency capture **before** the converter optimization work (W3). 
Product target for the instant-preview pipeline: **<500 ms per slide** (whole-deck PNG @150ppi path).

## Environment

| Property | Value |
|---|---|
| OS | Darwin 25.5.0 Darwin Kernel Version 25.5.0: Mon Apr 27 20:41:12 PDT 2026; root:xnu-12377.121.6~2/RELEASE_ARM64_T6050 |
| Architecture | Arm64 |
| .NET runtime | .NET 9.0.17 |
| .NET SDK | 10.0.301 |
| TypstBridge | 0.1.0 |
| LibreOffice | not available on this machine |

## .pptx (16 slides)

### This pipeline (PptxEditor → Typst)

| Stage | Cold (ms) | Warm median (ms, N=5) | Warm per-slide (ms) |
|---|---|---|---|
| PresentationBuilder.Open | 156.6 | 16.2 | 1.0 |
| ExportThumbnails (whole-deck PNG @150ppi) | 477.8 | 267.5 | 16.7 |
| ↳ png.compile (backend: bridge) | 214.9 | 103.9 | — |
| ExportToPdf | 273.3 | 149.5 | 9.3 |
| ↳ pdf.compile (backend: bridge) | 136.7 | 7.6 | — |
| **Total preview path (Open + PNG)** | **634.4** | **283.7** | **17.7** |

`*.compile` rows are the Typst compile (TypstBridge backend) inside the stage above,
captured via the `OFFICEEDITOR_TIMING` hooks in `TypstCompilerService`.

### LibreOffice (soffice --headless --convert-to pdf)

> LibreOffice not available on this machine (soffice not found on PATH or at /Applications/LibreOffice.app/Contents/MacOS/soffice) — leg skipped.

## Cold vs warm summary (preview path: Open + whole-deck PNG @150ppi)

| Deck | Slides | Cold total (ms) | Warm median total (ms, N=5) | Derived per-slide (ms) | LibreOffice warm median (ms) |
|---|---|---|---|---|---|
| .pptx | 16 | 634.4 | 283.7 | 17.7 | n/a (not installed) |

## Methodology

- Each deck is measured in a **fresh child process**: iteration 0 is the cold run,
  iterations 1..5 are the warm runs (reported as the median of N=5).
- Deck bytes are read from disk once, untimed; `PresentationBuilder.Open(byte[])` is the measured open stage.
- Stages per iteration: `PresentationBuilder.Open` → `ExportThumbnails` (whole-deck PNG @150ppi)
  → `ExportToPdf`. Wall-clock `Stopwatch` per stage.
- Typst compile time inside each export is captured with the opt-in `OFFICEEDITOR_TIMING=1`
  hooks in `TypstCompilerService` (backend, total compile ms per call).
- Per-slide times are derived as stage total ÷ slide count (the current pipeline renders the whole deck).
- LibreOffice leg: `soffice --headless --norestore --convert-to pdf --outdir <tmp> <deck>`,
  timed from process start to exit. Cold run uses a fresh `-env:UserInstallation` profile
  directory; warm runs reuse one profile (median of N=5). When `soffice` is not on PATH
  or at `/Applications/LibreOffice.app/Contents/MacOS/soffice`, the leg is skipped without failing.
- No third-party benchmark dependencies: `Stopwatch` + `Process` only.

## Limitations

- Single machine, single pass — absolute numbers carry normal OS/JIT noise; treat as orders of magnitude.
- Cold runs include one-time costs (assembly load/JIT, TypstBridge probe,
  system-font discovery) exactly as a fresh API process would experience them.
- `soffice --convert-to png` historically renders only slide 1 for Impress decks, so the LibreOffice
  leg converts to PDF and per-slide LO numbers are derived from the PDF total ÷ slide count.
- LibreOffice numbers (when present) include full process start + profile cost, not just rendering.
- Converter-internal stages (font extraction, system-font discovery, per-slide conversion) are not
  split out yet — `ExportThumbnails`/`ExportToPdf` are measured end to end, with the Typst compile
  portion attributed via the timing hooks.

