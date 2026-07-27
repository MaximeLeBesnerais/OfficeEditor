# Deck-Fidelity Batch 1 — Final Report

Branch: `fix/deck-batch1` (worktree `DocxEditor-fix-b1`) · 2026-07-28

## Commits (one per fix + one follow-up)

| Commit | Fix |
|---|---|
| `6f27ab1` | fix: resolve a:grpFill so group children inherit the parent group's fill |
| `8f2c405` | fix: resolve p:style fillRef/lnRef against the theme format scheme |
| `95a100e` | fix: emit small-caps for cap="small" text (master titleStyle inheritance) |
| `7ddbab4` | fix: synthesize small caps instead of Typst #smallcaps (fonts lack smcp) |

Test file (new, no existing test files touched):
`DocxEditor.Tests/Unit/PptxToTypstConverterDeckB1Tests.cs` (9 tests).

## Fix 1 — grpFill inheritance (`6f27ab1`)

**Root cause (verified):** `ExtractShapeFillColor` (PptxToTypstConverter.cs) only
read `a:solidFill`. FISHBONE/FRANCE group children declare `a:grpFill` +
`a:ln/a:noFill`, so fill resolved empty and the fill/stroke gate in `ConvertShape`
dropped every such shape — whole fishbones and the France map silhouette vanished.
`ConvertGroupShape` never read the group's own fill at all (only its transform).
Both decks' groups carry an explicit `a:solidFill` on `grpSpPr` (FISHBONE:
schemeClr tx2; FRANCE: srgb 68BC6C), so solid-fill inheritance suffices.

**Fix:** `ResolveGroupShapeFill` resolves the group's fill (solidFill → color;
noFill → none; grpFill/absent → pass through the inherited fill for nested
groups) and threads it ConvertGroupShape → ConvertElement → ConvertShape →
ExtractShapeGeometry → ExtractShapeFillColor, where `a:grpFill` resolves to it.

## Fix 2 — custGeom-in-group dropped (`8f2c405`)

**Drop path (verified, NOT the audit's "roundRect-first-child" guess):**
`ConvertGroupShape` recurses children through `ConvertElement` → `ConvertShape` →
`ExtractShapeGeometry`, which returns a valid shape element (roundRect/custGeom
are recognized) but with empty FillColor/StrokeColor — Opposites s5 Group 10/11
children carry NO fill marker and NO `a:ln` in spPr; their fill/stroke come only
from `p:style` → `a:fillRef idx=3/accent5` etc. (theme fmtScheme), which the
converter never resolved. The fill/stroke gate then discarded each child. The
same slide's Group 48 rendered because its children have explicit spPr solidFill.
FRANCE's map groups were a different mechanism (grpFill) and were fixed by Fix 1.

**Fix:** `StyleResolver.ResolveStyleFillReference` / `ResolveStyleLineReference`
resolve ref idx 1..3 against the theme fmtScheme (fillStyleLst/lnStyleLst),
substituting the reference's own color into the style's phClr placeholder while
preserving the style's transforms (tint/shade/satMod; gradients read via
GradientFillReader). `ApplyStyleReferenceFillAndStroke` in ConvertShape applies
them when spPr has no explicit fill marker / no a:ln; placeholders are excluded
(they inherit via the layout/master chain). REF decks audited: zero
placeholder+p:style collisions → no regression surface.

## Fix 3 — cap="small" (`95a100e` + `7ddbab4`)

**Root cause (verified):** detection and inheritance already worked
(`ExtractCapAttribute` returns "small"; master txStyles → DefaultTextStyle.Caps →
paragraph/run formatting). Only emission was missing: all six sites in
PptxToTypstConverter.SourceGeneration.cs branched solely on `caps == "all"`.
FRANCE: slideMaster2 titleStyle defRPr cap="small" → slides 1–4 titles.

**Fix, part 1:** shared `CapsOpenTag` helper; all six sites emit `all→#upper[`,
`small→#smallcaps[`.

**Fix, part 2 (approach correction):** native `#smallcaps` silently no-ops in
this pipeline — TypstBridge compiles with system fonts DISABLED and none of the
resolvable fonts (Open Sans, Carlito, Calibri, Arial…) carry OpenType smcp
(verified by byte-identical renders through libtypst_bridge). Final emission
synthesizes small caps with a scoped show rule:
`#[#show regex("\p{Ll}"): it => text(size: 0.8em)[#upper(it)]; …]`
— verified via the native bridge with Carlito + Open Sans: correct small-caps
look, block-scoped (no leak to siblings), survives embedded `#linebreak()`.

## RMSE before → after (scripts/rmse.py, --force renders)

| Deck | Mean | Key slides |
|---|---|---|
| FISHBONE | 11.0% → **8.9%** | s2 18.6→7.7 · s3 19.3→7.9 · s4 20.2→8.1 · s9 23.9→9.9 (all <15% ✓) |
| FRANCE | 13.7% → **9.2%** | s1 16.6→3.9 · s2 19.0→14.3 · s3 17.3→9.7 (all <15% ✓) |
| Opposites | 9.0% → **8.5%** | s5 21.8→8.9 (<15% ✓) |

No slide in any deck regressed by >1pp at any step. FISHBONE and Opposites were
byte-neutral under the cap fix, as expected.

## Tests

`DocxEditor.Tests/Unit/PptxToTypstConverterDeckB1Tests.cs`:
1. `Convert_GrpFillFreeformInsideGroup_InheritsParentGroupSolidFill`
2. `Convert_GrpFillFreeformInsideNestedGroup_InheritsOuterGroupFill`
3. `Convert_GrpFillFreeform_GroupHasNoFill_StillNotInvented`
4. `Convert_StyleFillReference_ResolvesThemeSolidFillAndLine`
5. `Convert_StyleFillReferenceGradient_ResolvesThemeGradientStops`
6. `Convert_StyleReference_DoesNotOverrideExplicitFill`
7. `GenerateTypstSource_RunCapSmall_EmitsSmallCaps`
8. `GenerateTypstSource_RunCapAll_StillEmitsUpper`
9. `GenerateTypstSource_MasterTitleStyleCapSmall_TitleInheritsSmallCaps`

Full `dotnet test` green: 1493 DocxEditor.Tests + 332 Api + 40 Mcp + 20 Cli + 36
TypstBridge — 0 failed, 0 warnings (TreatWarningsAsErrors).

## Visual verification (own eyeball, ours vs ref PNGs)

- FISHBONE s9: fishbone body/spine/head/tail fully restored, matches ref.
- FRANCE s1: green France silhouette + Corsica restored.
- FRANCE s3: map restored; header "MAP OF FRANCE W/ STATISTICS" in proper small
  caps (M/F/S full height, remainder smaller caps) matching ref.
- Opposites s5: red minus / green plus icons restored with correct theme
  gradients and white bars, matching ref.

## Delegation summary

- Speed agents (blocking): baseline RMSE runs; custGeom-in-group drop-path trace;
  grpFill code-path trace; cap-handling trace + FRANCE master/layout chain;
  test-harness pattern survey; all post-fix render+RMSE runs; Typst font/smcp
  investigation (byte-hash render comparisons through libtypst_bridge); two rounds
  of synthesized-small-caps snippet verification (compile + PNG inspection).
- Done by lead: root-cause confirmation against deck XML (FISHBONE grpSpPr,
  Opposites slide5 groups, theme fmtScheme/objectDefaults), fix design, all code
  and tests, full test-suite runs, commits, and all final eyeball checks.

## Not fixed (out of scope, for the record)

- FISHBONE number badges (rect with no fill/style/theme-default) still dropped —
  theme has no objectDefaults; small RMSE contribution.
- FRANCE s3: teardrop preset pins render as rotated diamonds; cxnSp connectors
  missing; s4 percentStacked/doughnut charts unsupported (audit classes 2/3/6).
- Opposites custGeom silhouette drift (class B) and Typst-vs-PPT baseline noise.
