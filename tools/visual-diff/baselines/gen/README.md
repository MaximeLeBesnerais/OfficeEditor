# gen parity baselines (Phase 5, P7)

The `gen` suite is the parity harness for the from-scratch generation vocabulary
(plan.md §5, rule 2): every Tier-1 primitive + linear gradient has one fixture deck,
emitted by BOTH emitters from the same resolved layout, then diffed as

```
PowerPoint render (ground truth)  vs  Typst render (spec of record)
```

per fixture, with a **per-primitive normalized-RMSE threshold** as the gate.

## Files

```text
tools/visual-diff/baselines/gen/
├── README.md           (this file)
├── thresholds.json     (per-primitive RMSE ceilings; mirrors FixtureCatalog,
│                        drift-guarded by FixtureCatalogTests)
└── metrics.json        (optional per-environment metrics baseline for drift
                         tracking; regenerate locally, see below)
```

`thresholds.json` is authored and machine-independent. The optional `metrics.json`
baseline (same shape as the pptx suite baseline) is **machine-dependent** — fonts,
PowerPoint version, typst version, ImageMagick — regenerate it per environment.

## The pipeline (render steps are opt-in)

```bash
# 1. Generate the fixture decks + Typst sources (deterministic, in-process, no
#    external tools — works in any environment)
dotnet run --project tools/visual-diff -- --suite gen --generate

# 2. Render the Typst preview pages (opt-in; needs the typst CLI)
dotnet run --project tools/visual-diff -- --suite gen --render

# 3. Produce the PowerPoint ground truth per fixture (manual, needs PowerPoint):
#    - open examples/output/gen/parity/<fixture>/fixture.pptx in PowerPoint
#    - File > Export > PDF to examples/output/gen/parity/<fixture>/ground-truth.pdf
#    - pdftocairo -png -r 150 ground-truth.pdf ground-truth/page

# 4. Run the parity gate (thresholds apply by default for --suite gen)
dotnet run --project tools/visual-diff -- --suite gen
```

Exit codes: `0` all fixtures within threshold, `2` threshold breach, `1` operational
error (missing tools/renders/thresholds). Without renders present the suite skips
fixtures loudly and exits 1 with these instructions — never a silent pass.

## Calibrating thresholds

The committed threshold values are provisional starting points. On a machine with
PowerPoint + typst + poppler + ImageMagick, run the full pipeline above, inspect
`examples/output/visual-diff/gen/metrics.json` per fixture, then set each threshold in
`FixtureCatalog` (and regenerate this file) to the observed worst page RMSE plus
headroom for renderer noise (antialiasing, hinting). Then review the diff images
before accepting. Regenerate `thresholds.json` from the catalog with the drift-guard
test's update path (`OE_UPDATE_SNAPSHOTS=1` on the catalog tests).

## Regenerating the optional metrics baseline

```bash
dotnet run --project tools/visual-diff -- --suite gen
cp examples/output/visual-diff/gen/metrics.json tools/visual-diff/baselines/gen/metrics.json
# then enforce drift vs that environment's baseline with:
dotnet run --project tools/visual-diff -- --suite gen \
  --baseline tools/visual-diff/baselines/gen/metrics.json --margin 0.05
```

Note `--baseline` and `--thresholds` are mutually exclusive per run; run the tool
twice to enforce both gates.
