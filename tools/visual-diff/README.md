# PDF Visual Diff Tool

Renders reference and generated PDFs to PNG images, compares each page with ImageMagick, and writes an HTML report plus JSON metrics. Use it to review visual regressions in generated PDF output.

## Requirements

- .NET 9 SDK
- Poppler tools: `pdftocairo` preferred, `pdftoppm` supported as a fallback
- ImageMagick: `compare`

Install the native tools if they are missing:

```bash
# Ubuntu/Debian
sudo apt install poppler-utils imagemagick

# Arch/Manjaro
sudo pacman -S poppler imagemagick
```

## Compare the DOCX reference suite

Run from the repository root:

```bash
dotnet run --project tools/visual-diff -- --suite docx
```

This compares the built-in DOCX reference PDFs under `examples/REF/DOCX/` with generated PDFs under `examples/output/ref/docx/`.

Default report output:

```text
examples/output/visual-diff/docx/
├── index.html
├── metrics.json
└── <document>/
    ├── reference/page-*.png
    ├── generated/page-*.png
    ├── diff/page-*.png
    └── metrics.json
```

Open `index.html` in a browser to inspect page thumbnails and diff images.

## Compare an arbitrary PDF pair

```bash
dotnet run --project tools/visual-diff -- \
  --ref reference.pdf \
  --gen generated.pdf \
  --out output-dir \
  --name "my comparison" \
  --dpi 150
```

Options:

| Option | Description |
|--------|-------------|
| `--suite docx` | Runs the built-in DOCX comparison suite. |
| `--ref <path>` | Reference PDF for a one-off comparison. |
| `--gen <path>` | Generated PDF for a one-off comparison. |
| `--out <path>` | Report directory. Required for arbitrary pairs; defaults to `examples/output/visual-diff/docx/` for `--suite docx`. |
| `--name <name>` | Display/report name for an arbitrary pair. |
| `--dpi <number>` | Rasterization DPI. Default: `150`. Higher values are more precise but slower and larger. |

## Reading the report

- **Reference** and **Generated** images show the rendered PDF pages.
- **Diff** images highlight pixel-level differences between the two renders.
- **RMSE** is the ImageMagick root-mean-square error. Lower is closer; `0` means the page images are pixel-identical.
- **Normalized/Percent RMSE** are scaled versions of the same measurement, useful for comparing pages of different content.
- A **page count mismatch** means only pages present in both PDFs were compared.

Pixel diffs can be noisy. Small RMSE values or faint diff pixels may come from antialiasing, font substitution, renderer differences, or DPI changes rather than meaningful layout regressions. Confirm suspicious pages visually before treating them as failures.

## Output files

Reports under `examples/output/visual-diff/docx/` are generated artifacts and are ignored by git via `examples/output/`. Do not commit them unless you intentionally add a curated fixture or reference artifact.
