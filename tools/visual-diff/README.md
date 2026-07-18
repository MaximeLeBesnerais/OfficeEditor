# Visual Diff Tool

Compares reference and generated output page-by-page and writes an HTML report plus JSON metrics. Inputs can be **PDFs** (rendered to PNG with poppler first) or **pre-rendered PNGs** (compared directly). Use it to review visual regressions in generated output, or — with `--baseline` — to enforce an RMSE threshold in a gate.

## Requirements

- .NET 9 SDK
- ImageMagick: `compare` (or ImageMagick 7's `magick`, invoked as `magick compare`) — required for all comparisons
- Poppler tools: `pdftocairo` preferred, `pdftoppm` supported as a fallback — **only required for PDF inputs**; PNG-pair mode skips poppler entirely

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

# PPTX suite: examples/REF/PPTX/{REMOVED,pres-pro}.pdf vs examples/output/ref/pptx/*.pdf
dotnet run --project tools/visual-diff -- --suite pptx
```

Default report output is `examples/output/visual-diff/<suite>/`.

### The PPTX suite and generated PDFs

`--suite pptx` diffs the two committed reference PDFs (`examples/REF/PPTX/REMOVED.pdf`, `pres-pro.pdf`) against generated PDFs under `examples/output/ref/pptx/`. Generated PDFs are produced by the repo's existing converter, `tools/convert-pptx`. If one is missing, the suite prints the exact command and skips that deck:

```bash
dotnet run --project tools/convert-pptx -- examples/REF/PPTX/REMOVED.pptx \
  examples/output/ref/pptx/REMOVED.pdf --format pdf
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
  --ref docs/assets/pptx-comparison/official/REMOVED \
  --gen examples/output/pptx-comparison/generated/REMOVED \
  --out examples/output/visual-diff/pptx-comparison/REMOVED \
  --name REMOVED
```

`--ref` and `--gen` must be the same kind of input (both PDFs, both PNG files, or both directories).

## Threshold checks (RMSE gate)

The core tool is **report-only**: without `--baseline` it always exits 0 on success, regardless of RMSE values. The threshold wrapper engages only when you ask for it, and exits non-zero on regression:

```bash
# Standalone: check an existing metrics.json against a baseline
dotnet run --project tools/visual-diff -- \
  --check examples/output/visual-diff/pptx/metrics.json \
  --baseline tools/visual-diff/baselines/pptx/metrics.json \
  --margin 0.05

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
| `--suite <docx\|pptx>` | Runs a built-in comparison suite. |
| `--ref <path>` | Reference PDF, single PNG, or directory of PNGs. |
| `--gen <path>` | Generated PDF, single PNG, or directory of PNGs (same kind as `--ref`). |
| `--out <path>` | Report directory. Required for arbitrary pairs; defaults to `examples/output/visual-diff/<suite>/` for suites. |
| `--name <name>` | Display/report name for an arbitrary pair. |
| `--dpi <number>` | Rasterization DPI for PDF inputs. Default: `150`. Higher values are more precise but slower and larger. |
| `--generate` | (`--suite pptx`) Build missing generated PDFs via `tools/convert-pptx`. |
| `--font-path <dir>` | Extra font directory passed to `tools/convert-pptx` when using `--generate`. |
| `--check <file>` | Threshold-check an existing metrics.json against `--baseline`. |
| `--baseline <file>` | Baseline metrics.json; after a run, check the fresh metrics against it. |
| `--margin <number>` | Absolute margin on normalized RMSE (`0.05` = 5 points). Default: `0.05`. |
| `--probe` | Print external-tool availability and exit. |
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
