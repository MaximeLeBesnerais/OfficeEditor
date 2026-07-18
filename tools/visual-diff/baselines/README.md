# visual-diff baselines

Threshold-checked visual regression gates compare a fresh `metrics.json` against a
baseline stored here. **Baselines are machine-dependent.** RMSE values shift with
installed fonts, the rendering backend (TypstBridge vs CLI fallback), poppler and
ImageMagick versions, and DPI. Do not hand-edit numbers, and do not reuse a
baseline produced on a different machine or font environment — regenerate it.

A baseline file is a verbatim copy of a suite run's `metrics.json` (same shape;
`GeneratedAt`/`Dpi` in the file record where it was produced).

## Layout

```text
tools/visual-diff/baselines/
├── README.md           (this file)
└── pptx/
    └── metrics.json    (baseline for --suite pptx; regenerate per environment)
```

## Regenerating the PPTX baseline

Run from the repository root on a machine with poppler + ImageMagick installed
(`brew install poppler imagemagick` on macOS, `apt install poppler-utils imagemagick`
on Debian/Ubuntu):

```bash
# 1. Produce the generated PDFs (once, or after converter changes)
dotnet run --project tools/convert-pptx -- examples/REF/PPTX/Presentation1.pptx \
  examples/output/ref/pptx/Presentation1.pdf --format pdf
dotnet run --project tools/convert-pptx -- examples/REF/PPTX/pres-pro.pptx \
  examples/output/ref/pptx/pres-pro.pdf --format pdf

# 2. Run the suite (or add --generate to run step 1 automatically)
dotnet run --project tools/visual-diff -- --suite pptx

# 3. Promote the fresh metrics to the baseline
mkdir -p tools/visual-diff/baselines/pptx
cp examples/output/visual-diff/pptx/metrics.json tools/visual-diff/baselines/pptx/metrics.json
```

## Enforcing the baseline

```bash
dotnet run --project tools/visual-diff -- --suite pptx \
  --baseline tools/visual-diff/baselines/pptx/metrics.json \
  --margin 0.05
```

Exit codes: `0` within baseline + margin, `2` on any per-page or per-document
average breach, `1` on operational errors (missing tools, missing baseline).
`--margin` is in normalized RMSE units: `0.05` = 5 percentage points.
