# Visual Diff Tool

> Note: the  /  REF fixtures and the committed `baselines/pptx/` baselines were removed (licensing). Regenerate baselines with `--update-baseline` once replacement REF fixtures land.

Compares reference and generated output page-by-page and writes an HTML report plus JSON metrics. Inputs can be **PDFs** (rendered to PNG with poppler first) or **pre-rendered PNGs** (compared directly). Use it to review visual regressions in generated output, or — with `--baseline` — to enforce an RMSE threshold in a gate.

## Requirements

- .NET 9 SDK
- ImageMagick: `compare` (or ImageMagick 7's `magick`, invoked as `magick compare`) — required for all comparisons
- Poppler tools: `pdftocairo` preferred, `pdftoppm` supported as a fallback — **only required for PDF inputs**; PNG-pair mode skips poppler entirely
- typst CLI — **only required for `--suite gen --render`**; the gen suite itself diffs pre-rendered PNGs and needs nothing beyond ImageMagick

External tools are probed in `/usr/bin`, `/opt/homebrew/bin` (Homebrew on Apple Silicon), then `PATH`. Run `--probe` to see what was found:

```bash
dotnet run --project tools/visual-diff -- --probe
```

Install the native tools if they are missing:

```bash
# macOS (Homebrew)
brew install poppler imagemagick

# Ubuntu/Debian
sudo apt install poppler-utils imagemagick

# Arch/Manjaro
sudo pacman -S poppler imagemagick
```

When a required tool is missing, the tool exits 1 before doing any work and prints an install hint plus the probed locations.

## Compare the built-in suites

Run from the repository root:

```bash
# DOCX suite: examples/REF/DOCX/*.pdf vs examples/output/ref/docx/*.pdf
dotnet run --project tools/visual-diff -- --suite docx

# PPTX suite: examples/REF/PPTX/northwind-demo.pdf vs examples/output/ref/pptx/*.pdf
dotnet run --project tools/visual-diff -- --suite pptx

# GEN suite (Phase 5 parity fixtures): PowerPoint ground truth vs Typst preview,
# per primitive, with per-primitive RMSE thresholds as the gate
dotnet run --project tools/visual-diff -- --suite gen --generate
```

Default report output is `examples/output/visual-diff/<suite>/`.

### The gen suite (Phase 5 parity harness)

`--suite gen` is the parity harness for the from-scratch generation vocabulary (plan.md §5, rule 2): every Tier-1 primitive + linear gradient has one fixture deck (catalog: `PptxEditor.Core/Generation/Fixtures/FixtureCatalog`), emitted by BOTH emitters from the same resolved layout, then diffed as **PowerPoint render (ground truth) vs Typst render (spec of record)** per fixture.

The pipeline has three steps; the two render steps are external and opt-in by design:

```bash
# 1. Generate fixture decks + Typst sources (deterministic, in-process, no external tools)
dotnet run --project tools/visual-diff -- --suite gen --generate

# 2. Render the Typst preview pages (needs the typst CLI — probed)
dotnet run --project tools/visual-diff -- --suite gen --render

# 3. Produce the PowerPoint ground truth manually per fixture
#    (open examples/output/gen/parity/<fixture>/fixture.pptx in PowerPoint, export PDF,
#     pdftocairo -png -r 150 ground-truth.pdf ground-truth/page)

# 4. Diff + gate (thresholds from baselines/gen/thresholds.json apply by default)
dotnet run --project tools/visual-diff -- --suite gen
```

Missing renders are loud per-fixture skips with the exact production steps; if nothing can be compared the suite exits 1. The gate is `ThresholdCheck`: any page or fixture-average normalized RMSE above its per-primitive ceiling exits 2. See [baselines/gen/README.md](baselines/gen/README.md) for threshold calibration.

### The PPTX suite and generated PDFs

`--suite pptx` diffs the committed reference PDF (`examples/REF/PPTX/northwind-demo.pdf`) against generated PDFs under `examples/output/ref/pptx/`. Generated PDFs are produced by the repo's existing converter, `tools/convert-pptx`. If one is missing, the suite prints the exact command and skips that deck:

```bash
dotnet run --project tools/convert-pptx -- examples/REF/PPTX/northwind-demo.pptx \
  examples/output/ref/pptx/northwind-demo.pdf --format pdf
```

Or let visual-diff invoke the converter itself for any missing deck:

```bash
dotnet run --project tools/visual-diff -- --suite pptx --generate [--font-path /usr/share/fonts]
```

## Compare an arbitrary pair

### PDF pair (needs poppler + ImageMagick)

```bash
dotnet run --project tools/visual-diff -- \
  --ref reference.pdf \
  --gen generated.pdf \
  --out output-dir \
  --name "my comparison" \
  --dpi 150
```

### PNG pair (needs ImageMagick only — no poppler)

`--ref`/`--gen` accept a single `.png` or a directory of PNGs. Directories are paired up in natural page order (`page-2` before `page-10`); a single PNG is a one-page input. Inputs are copied into the report tree, then fed straight to the ImageMagick compare logic — the poppler render step is skipped entirely.

```bash
dotnet run --project tools/visual-diff -- \
  --ref docs/assets/pptx-comparison/official/ \  # NOTE: dir removed (licensing)
  --gen examples/output/pptx-comparison/generated/ \
  --out examples/output/visual-diff/pptx-comparison/ \
  --name 
```

`--ref` and `--gen` must be the same kind of input (both PDFs, both PNG files, or both directories).

## Threshold checks (RMSE gate)

Two gates exist, mutually exclusive per run:

- **`--baseline <metrics.json>`** — machine-dependent baseline + margin gate (REF-deck suites). Exits 2 when any per-page or per-document average normalized RMSE exceeds `baseline + margin`.
- **`--thresholds <thresholds.json>`** — per-primitive absolute RMSE ceilings (gen parity suite). Exits 2 when any page or fixture average exceeds its fixture's threshold.

The core tool is **report-only**: without a gate it always exits 0 on success, regardless of RMSE values. The threshold wrapper engages only when you ask for it, and exits non-zero on regression:

```bash
# Standalone: check an existing metrics.json against a baseline
dotnet run --project tools/visual-diff -- \
  --check examples/output/visual-diff/pptx/metrics.json \
  --baseline tools/visual-diff/baselines/pptx/metrics.json \
  --margin 0.05

# Standalone: check an existing metrics.json against per-primitive thresholds
dotnet run --project tools/visual-diff -- \
  --check examples/output/visual-diff/gen/metrics.json \
  --thresholds tools/visual-diff/baselines/gen/thresholds.json

# Combined: run the suite, then check the fresh metrics
dotnet run --project tools/visual-diff -- --suite pptx \
  --baseline tools/visual-diff/baselines/pptx/metrics.json \
  --margin 0.05
```

A check fails (exit 2) when any **per-page** or any **per-document average** normalized RMSE exceeds `baseline + margin`, when the page count mismatches, or when a page/document has no usable metric. `--margin` is absolute in normalized RMSE units: `0.05` = 5 percentage points (default: `0.05`).

Exit codes:

| Code | Meaning |
|------|---------|
| 0 | Success; thresholds (if checked) satisfied |
| 1 | Operational or usage error (missing tools/files, bad arguments) |
| 2 | Threshold check ran and found at least one breach |

### Baselines

Baselines live in [`baselines/`](baselines/) as metrics.json-shaped files. **They are machine-dependent** (fonts, rendering backend, poppler/ImageMagick versions) and must be regenerated per environment — see [baselines/README.md](baselines/README.md) for the regeneration commands. A missing baseline is an operational error (exit 1) with instructions, never a silent pass.

## Options

| Option | Description |
|--------|-------------|
| `--suite <docx\|pptx\|gen>` | Runs a built-in comparison suite. |
| `--ref <path>` | Reference PDF, single PNG, or directory of PNGs. |
| `--gen <path>` | Generated PDF, single PNG, or directory of PNGs (same kind as `--ref`). |
| `--out <path>` | Report directory. Required for arbitrary pairs; defaults to `examples/output/visual-diff/<suite>/` for suites. |
| `--name <name>` | Display/report name for an arbitrary pair. |
| `--dpi <number>` | Rasterization DPI for PDF inputs and `--render`. Default: `150`. Higher values are more precise but slower and larger. |
| `--generate` | (`--suite pptx`) Build missing generated PDFs via `tools/convert-pptx`. (`--suite gen`) Generate fixture decks + Typst sources in-process. |
| `--font-path <dir>` | Extra font directory passed to `tools/convert-pptx` when using `--generate`. |
| `--render` | (`--suite gen`) Render fixture `.typ` sources to PNG pages via the typst CLI (probed; opt-in). |
| `--check <file>` | Threshold-check an existing metrics.json against `--baseline` or `--thresholds`. |
| `--baseline <file>` | Baseline metrics.json; after a run, check the fresh metrics against it. Mutually exclusive with `--thresholds`. |
| `--thresholds <file>` | Per-primitive thresholds.json (gen suite); after a run, check the fresh metrics against the per-fixture RMSE ceilings. Defaults for `--suite gen` to `tools/visual-diff/baselines/gen/thresholds.json` when present. |
| `--margin <number>` | Absolute margin on normalized RMSE (`0.05` = 5 points) for `--baseline` checks. Default: `0.05`. |
| `--probe` | Print external-tool availability (poppler, ImageMagick, typst) and exit. |
| `--help` | Show usage. |

## Reading the report

- **Reference** and **Generated** images show the rendered pages.
- **Diff** images highlight pixel-level differences between the two renders.
- **RMSE** is the ImageMagick root-mean-square error. Lower is closer; `0` means the page images are pixel-identical.
- **Normalized/Percent RMSE** are scaled versions of the same measurement, useful for comparing pages of different content. Threshold checks use the normalized value.
- A **page count mismatch** means only pages present in both inputs were compared.

Pixel diffs can be noisy. Small RMSE values or faint diff pixels may come from antialiasing, font substitution, renderer differences, or DPI changes rather than meaningful layout regressions. Confirm suspicious pages visually before treating them as failures.

## Output files

Reports under `examples/output/visual-diff/` are generated artifacts and are ignored by git via `examples/output/`. Do not commit them unless you intentionally add a curated fixture or reference artifact.
