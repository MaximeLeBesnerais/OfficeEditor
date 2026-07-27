# Deck-Fidelity Batch 1 — Progress

Branch: `fix/deck-batch1` (worktree `DocxEditor-fix-b1`)

## Baselines (measured 2026-07-28, `scripts/rmse.py`)

| Deck | Mean | Worst slides |
|---|---|---|
| FISHBONE | 11.0% | s9 23.9%, s4 20.2%, s3 19.3%, s2 18.6% |
| FRANCE | 13.7% | s2 19.0%, s3 17.3%, s1 16.6%, s4 14.6% |
| Opposites | 9.0% | s5 21.8% |

Note: `rmse.py --no-render` requires the sibling PDFs (`local-ref/<deck>/<deck>.pdf`);
they were copied from the main checkout (gitignored, never committed). Ours renders
cached under `/tmp/rmse-<deck>/ours/`.

## Fix 1 — grpFill inheritance (FISHBONE s2/3/4/9, FRANCE s1–3 map)

- Verified root cause: `ExtractShapeFillColor` (PptxToTypstConverter.cs:1113) only
  reads `a:solidFill`; `a:grpFill` children resolve to empty fill and are dropped by
  the gate at :840. Group's own fill (grpSpPr) is never read (`ConvertGroupShape`
  :2217 only uses the transform).
- FISHBONE groups: grpSpPr has explicit `a:solidFill` (schemeClr tx2) → solid-fill
  inheritance suffices. FRANCE Groups 23/40/115 same pattern (green 68BC6C).
- Plan: thread resolved group fill through ConvertGroupShape → ConvertElement →
  ConvertShape → ExtractShapeGeometry → ExtractShapeFillColor.
- Status: DONE (commit 6f27ab1). FISHBONE mean 11.0→8.9% (s2 18.6→7.7, s3 19.3→7.9,
  s4 20.2→8.1, s9 23.9→9.9 — all <15%, no regressions >1pp). FRANCE mean 13.7→10.2%
  (s1 16.6→3.9, s2 19.0→14.8, s3 17.3→13.0 — all <15%). Eyeballed FISHBONE s9 +
  FRANCE s1: fishbone and map silhouette fully restored.
- Tests: Convert_GrpFillFreeformInsideGroup_InheritsParentGroupSolidFill,
  Convert_GrpFillFreeformInsideNestedGroup_InheritsOuterGroupFill,
  Convert_GrpFillFreeform_GroupHasNoFill_StillNotInvented.

## Fix 2 — custGeom-in-group dropped (Opposites s5 Groups 10/11)

- Verified drop path (NOT "roundRect-first-child" as audit guessed): per-child gate
  failure in `ConvertShape` (:840). Group 10/11 children have NO fill in spPr and no
  usable stroke; their fill/stroke come only from `p:style` → `a:fillRef`/`a:lnRef`
  (theme fmtScheme), which the converter never resolves. Group 48 renders because its
  children carry explicit `a:solidFill` in spPr.
- FRANCE map groups are fixed by Fix 1 instead (children use grpFill).
- Plan: resolve `p:style` fillRef/lnRef against theme fmtScheme (phClr substitution)
  in StyleResolver; apply in ConvertShape when spPr lacks explicit fill/line and the
  shape is not a placeholder. REF decks checked: no placeholder+p:style collisions,
  no regression risk.
- Status: DONE (commit 8f2c405). Opposites s5 21.8→8.9% (mean 9.0→8.5%),
  zero regressions. Eyeballed s5: minus/plus icons render with correct
  gradient fills + white bars, matching ref.
- Tests: Convert_StyleFillReference_ResolvesThemeSolidFillAndLine,
  Convert_StyleFillReferenceGradient_ResolvesThemeGradientStops,
  Convert_StyleReference_DoesNotOverrideExplicitFill.

## Fix 3 — cap="small" not emitted (FRANCE headers s1–4)

- Verified: detection + inheritance of cap="small" already work (ExtractCapAttribute
  returns "small"; master txStyles loaded via StyleResolver). Emission sites in
  PptxToTypstConverter.SourceGeneration.cs only branch on `caps == "all"` (6 sites).
- FRANCE: slideMaster2 titleStyle defRPr cap="small"; slides 1–4 titles map to it.
  Theme minor font = Calibri (→ Carlito, has OpenType smcp).
- Plan: emit `#smallcaps[...]` for caps == "small" alongside the `#upper[...]` path.
- Status: DONE (commits 95a100e + 7ddbab4). IMPORTANT: native `#smallcaps` no-ops
  in this pipeline — TypstBridge disables system fonts and none of the resolvable
  fonts (Open Sans, Carlito, Calibri, ...) carry OpenType smcp. Final emission
  synthesizes small caps via a scoped show rule:
  `#[#show regex("\p{Ll}"): it => text(size: 0.8em)[#upper(it)]; …]`
  (verified via native bridge compiles: works with Carlito + Open Sans, scoped to
  the block, survives embedded #linebreak()).
- FRANCE after fix: s1 3.9, s2 14.3, s3 9.7, s4 13.4, s5 7.0, s6 6.9 — mean 9.2%
  (was 13.7% baseline). Eyeballed s3 header: proper small caps vs ref.
- Tests: GenerateTypstSource_RunCapSmall_EmitsSmallCaps,
  GenerateTypstSource_RunCapAll_StillEmitsUpper,
  GenerateTypstSource_MasterTitleStyleCapSmall_TitleInheritsSmallCaps.

## Final state (all targets met)

- FISHBONE mean 11.0→8.9%; s2/3/4/9 all <15% ✓; no regression >1pp ✓
- FRANCE mean 13.7→9.2%; s1/2/3 all <15% ✓; no regression >1pp ✓
- Opposites mean 9.0→8.5%; s5 21.8→8.9% <15% ✓; no regression >1pp ✓
- Full `dotnet test` green (1493 DocxEditor.Tests + 332 Api + 40 Mcp + 20 Cli + 36 TypstBridge), 0 warnings.
