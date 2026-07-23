# pptx-benchmark

Pre-optimization latency ground truth for the PPTX render pipeline (PptxEditor → Typst),
compared against headless LibreOffice. Product target: **<500 ms per slide** whole-deck PNG @150ppi.

## Usage

```bash
dotnet run --project tools/pptx-benchmark
dotnet run --project tools/pptx-benchmark -- --runs 5 --update-baseline
```

| Flag | Default | Meaning |
|---|---|---|
| `--runs N` | 5 | Warm iterations per deck (median reported); one cold iteration always runs first in a fresh child process |
| `--out <path>` | `examples/output/benchmark/report.md` | Report output path (gitignored working copy) |
| `--update-baseline` | off | Also copy the report to `baselines/baseline-pre-optimization.md` (tracked) |

## What it measures

Per deck (the license-clean REF corpus under `examples/REF/PPTX/`: sales_acceleration_deck, AetherLink-Glass-Shareholder-Overview, northwind-launch-review, northwind-investor-40, northwind-demo), in a fresh child process:

- **Cold**: iteration 0 — `PresentationBuilder.Open`, `ExportThumbnails` (whole-deck PNG @150ppi), `ExportToPdf`.
- **Warm**: median of N repeat iterations in the same process.
- Typst compile time inside each export, via the opt-in `OFFICEEDITOR_TIMING=1` hooks in
  `OfficeEditor.Core/Services/TypstCompilerService.cs` (zero behavior/performance change when off).
- **LibreOffice leg** (optional): `soffice --headless --norestore --convert-to pdf` with a fresh
  `-env:UserInstallation` profile (cold) vs a reused profile (warm, median of N). Probes `PATH` and
  `/Applications/LibreOffice.app/Contents/MacOS/soffice`; if LibreOffice is absent the leg is skipped
  cleanly and the absence is recorded in the report. LibreOffice is **not** a build dependency.

## Outputs

- `examples/output/benchmark/report.md` — gitignored working copy.
- `baselines/baseline-pre-optimization.md` — **tracked** ground truth, committed before the
  converter optimization branch (W3) lands. Regenerate with `--update-baseline` only when
  deliberately re-baselining.

No third-party dependencies: `Stopwatch` + `Process` only.
