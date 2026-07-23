# PPTX roadmap

PPTX is the proof-carrier of the OfficeEditor suite: it is the only format with the full arc —
fluent builder (`PptxEditor.Core/Builders`), edit instructions (`PptxEditor.Core/Instructions`),
variable/mail-merge, and the generation pipeline (JSON v2.0 → loud schema validator → archetypes →
components → layout → dual emit OOXML + Typst), wrapped by the API, the MCP host, and the web demo.
What ships well here defines the product. The arc of this roadmap:
**fixture regeneration & trust → converter fidelity → generation vocabulary depth → ecosystem hardening → v1.**

Item sizes: **S** ≲ 1 day, **M** ≲ 1 week, **L** multi-week. Phases are dependency-ordered, not
date-ordered: nothing in Phase 1+ is trustworthy until Phase 0 lands, and the schema freeze
(Phase 2) gates the v0.5 release.

Standing disciplines (apply to every item, per AGENTS.md):

- Layout once, emit twice — every new primitive ships with both the OOXML and the Typst emitter plus a parity fixture.
- The Typst preview is the spec of record for ambiguous OOXML rendering; never mutate existing style definitions in template decks.
- 0-warnings policy holds; `dotnet test` after every change; smoke tests must pass on `examples/REF/PPTX/northwind-demo.pptx` (and its successors).

---

## Phase 0 — Fixture regeneration & trust restoration

**Blocks:** everything. Phases 1–2 measure success against visual parity; without a licensed corpus there is nothing to measure against.

The REF corpus is self-made and lives in `examples/REF/PPTX/`: `sales_acceleration_deck` (primary,
16 slides, SmartArt on slide 15), `AetherLink-Glass-Shareholder-Overview`, `northwind-investor-40`,
`northwind-launch-review`, plus the OfficeEditor-generated `northwind-demo`;
`tools/visual-diff/baselines/pptx/` is generated per machine and not committed.
be regenerated against the new corpus. The showeet-licensed SmartArt corpus stays local-only under
`local-ref/smartarts/` (gitignored). Remaining work from this phase:
  Why: `baselines/gen` exists but `baselines/pptx` does not; the pptx suite has no gate. Baselines
  are machine-dependent — regenerate per the baselines README rules. REF PDFs double as ground
  truth for the `LibreOfficeCompareService` path.
  **S** — Acceptance: `--suite pptx` threshold-checked in CI on the new corpus.
- **Un-skip the remaining guarded tests** and make missing REF decks fail loudly instead of silently
  shrinking theory data. (The SmartArt / solid-fill / AetherLink image tests were un-guarded on the
  new corpus; `PresPro_Slide1_TitleFontSize_FromMaster` stays guarded — no license-clean deck uses
  master-inherited title placeholders.)
  **S** — Acceptance: smoke + parity tests run against all new decks; zero tests skipped for missing files.
- **Re-mine design tokens** — `Generation/Design/pres-pro.tokens.json` and
  `aetherlink.tokens.json` predate the current corpus. Re-run `BrandProfileExtractor`
  against the new styled deck and commit replacement token sets.
  **S** — Acceptance: token JSONs traceable to a deck in the repo; mining test round-trips.
- **WebApplicationFactory smoke tests for `OfficeEditor.Api/Program.cs`** (938 lines, ~0% covered — the largest single coverage hole; no WebApplicationFactory usage today). Cover the 20 mapped endpoints at the happy-path + error-contract level:
  - Deck lifecycle: `POST/GET /api/decks`, anatomy, file, slides/{n}/preview, instructions.
  - Generation: `POST /api/decks/generate`, `/api/convert`, `/api/samples`.
  - Demo/compare: `/api/demo/render*`, `/api/demo/compare/{typst,libreoffice}*`, `/api/demo/decks`, deck-template.
  - Misc: `/api/health`, `/api/preview/{id}`, `/api/download/{id}`.
  **M** — Acceptance: Program.cs exercised at integration level; API tests no longer unit-only; error payloads match the documented contract.
- **OfficeEditor.Cli instrumentation** — no test project references `OfficeEditor.Cli.csproj`
  today (verified: only self-reference repo-wide).
  Why: the multi-format CLI is a shipped surface with zero regression net.
  **M** — Acceptance: CLI test project with command-level tests (`create`, `edit`, `detect`,
  `merge`, `generate`) against temp files.

## Phase 1 — Converter fidelity (`PptxToTypstConverter`)

**Depends on:** Phase 0 corpus. Work happens behind parity fixtures so fixes are provable, not anecdotal.

The converter is the largest class in the suite (4,696 lines — it grew substantially during the
0.15 upgrade) and the biggest class-level coverage gap (~280 uncovered lines). Known artifacts are
documented in `docs/typst-0.15.md`. Per AGENTS.md: **every fix ships with a parity fixture** (OOXML +
Typst emitters + RMSE fixture) — no half-tested features. The Typst preview is the spec of record.

- **Split the converter first** — `PptxToTypstConverter.cs` at 4,696 lines is too big to safely
  absorb a phase of fidelity surgery. Break it into partials/strategies by content type (text,
  tables, images, shapes, charts, SmartArt) before the items below.
  Why: everything else in this phase touches this file; doing it first keeps each fix small and reviewable.
  **M** — Acceptance: same public API, no behavior change, full suite green, converter ≥ current coverage in every new file.

- **Z-order / panel-overlap fixes on styled panels.**
  Why: visible artifact on real decks; erodes trust in the preview-as-spec.
  **M** — Acceptance: new fixture deck reproducing the overlap passes the RMSE gate; artifact gone
  across the REF corpus.
- **Big-numeral caption overlap** (KPI-style slides).
  **S–M** — Acceptance: fixture + regression test; no overlap at any configured PPI.
- **Table style depth** — current state is "start header bold/white, avoid complex partial
  strokes" (AGENTS.pptx.md). Add banded rows, partial strokes, cell margins.
  **M** — Acceptance: one fixture per style feature; visual-diff delta vs PowerPoint within threshold.
- **SmartArt beyond approximation** (currently approximated with warnings) — or honest documented
  limits listing which diagram types degrade.
  **L** — Acceptance: top-N diagram types render structurally; the rest warn loudly against a
  documented support matrix in `docs/`.
- **Chart rendering improvements** — currently simplistic.
  **M** — Acceptance: bar/line/pie render with axis labels and legend; fixture-backed. (Distinct
  from the Phase 2 *generation* chart component — this is conversion of existing charts.)
- **Image fit edge cases** — extend `PptxImageFrameSizingTests` coverage: crop, stretch, tile, aspect-locked.
  **S** — Acceptance: fixture per fit mode.
- **Gradient fidelity spot-check** — typst 0.15 fixed excessive sampling of linear gradients and
  LinearRGB/CMYK gradient bugs (`docs/typst-0.15.md` §1); gradient-heavy slides should now export
  faster and smaller. Re-diff any gradient fixtures.
  **S** — Acceptance: gradient slides within threshold.
- **Groups & connectors** — group shapes with zero child extents already handled defensively
  (`Convert_GroupShapeWithZeroChildExtents_DoesNotThrow...`); extend to nested groups, rotations,
  and connector routing.
  **M** — Acceptance: fixture deck with nested/rotated groups passes the RMSE gate.
- **Aptos→Carlito metric drift check** — `FontMetricsCatalog` reads default-instance metrics; if a
  user supplies real Aptos (a variable font) via font-path, typst 0.15 instantiates it properly and
  metrics can drift (`docs/typst-0.15.md` §3).
  **S** — Acceptance: documented behavior + test pinning current substitution.
- **Coverage burn-down** on the ~280 uncovered lines, driven by the new REF corpus.
  **M** — Acceptance: converter ≥ 93% line coverage.

## Phase 2 — Generation vocabulary depth

**Depends on:** Phase 1 far enough along that new primitives land on a trustworthy emitter base. Can overlap with late Phase 1.

The vocabulary is young: five archetypes (`cover`, `section`, `kpi_row`, `two_col`, `table_slide`
— `ArchetypeSlides.Names`), eight components (`Badge`, `BulletList`, `Card`, `Divider`,
`ImageCard`, `Kpi`, `TableBlock`, `TitleBlock`), and a theme model of palette/fonts/metrics/shape
only (`Generation/Model/DesignTokens.cs` — `MetricTokens` is four values). The schema is version
`"2.0"` (`GenerationDocumentParser.SupportedVersion`) but not declared stable.

- **More archetypes**: quote, image+text splits, comparison, timeline, stats row, section-divider
  variants.
  Why: these cover the bulk of real business decks. Each must expand to existing
  components/primitives — archetypes never nest and own slot geometry (`ArchetypeSlides` rules).
  **M each** — Acceptance: dual-emit + parity fixture per archetype; validator round-trip tests.
- **Chart component** — bar/line from data → native OOXML chart, with a Typst-rendered
  approximation for preview; dual-emit discipline.
  **L** — Acceptance: OOXML chart editable in PowerPoint; preview approximates layout; fixture passes the RMSE gate.
- **Richer theme model** — typography scale, spacing scale, corner presets, possibly shadows.
  Why: brand decks need more than palette + two font slots.
  **M** — Acceptance: token JSONs in `Generation/Design/` regenerated from the new styled REF deck; schema documented.
- **Schema stability pass → declare v1 freeze of the 2.0 vocabulary** + migration notes.
  Why: the repo's own roadmap lists "Stable v1 generation schema"; freeze gates external authoring
  (MCP `deck_generate`, web client generate tab, external AI generators).
  **M** — Acceptance: versioned schema doc, additive-change policy, validator rejects unknown
  versions with a clear message.
- **Validator UX** — the loud validator is a differentiator; add did-you-mean suggestions and
  path-precise fixes.
  **S–M** — Acceptance: fuzzed malformed docs produce actionable diagnostics; MCP error passthrough verified.

## Phase 3 — Edit engine depth + ecosystem

**Depends on:** Phase 2 schema freeze (edit ops reference archetypes; MCP parity references the frozen vocabulary). Ecosystem items (CI platforms, `CompileOptions`, web client) can start anytime.

`PptxInstructionEngine` supports six ops: `PptxReplaceTextInstruction`,
`PptxReplaceImageInstruction`, `PptxReplaceTableInstruction`, `PptxMoveSlideInstruction`,
`PptxDuplicateSlideInstruction`, `PptxDeleteSlideInstruction` — anything else throws
`NotSupportedException`. Generation is deeper than editing today.

- **Edit-instruction completeness audit + gap fill**: add slide (from archetype), restyle (apply a
  theme to an existing deck), richer table ops (insert/delete row/column).
  **M–L** — Acceptance: op catalog documented; each op has engine + builder + API/MCP tests.
- **MCP tool parity with API** — the host exposes 4 tools (`deck_anatomize`,
  `deck_replace_element`, `deck_render_slide`, `deck_generate` in `ToolSchemas`); map the remaining
  API surface deliberately, or document why not.
  **S–M** — Acceptance: parity matrix in `OfficeEditor.Mcp/README.md`.
- **Expose typst 0.15 capabilities via `CompileOptions`** — combined PDF/A + PDF/UA-1 standards and
  `PdfOptions.creator` metadata (adoption plan already in `docs/typst-0.15.md` §1: extend `abi.rs`
  + `CompileOptions`, pass through to `PdfStandards::new`).
  **S** — Acceptance: bridge ABI round-trip; rendered PDF validates against the requested standards.
- **Windows/Linux runtime verification in CI** — developed on macOS; the win-x64 native lib is
  built in CI but never run. Add a smoke job on `windows-latest` + `ubuntu-latest` exercising the
  native libs (render one fixture deck end-to-end). Also verify the typst 0.15 MSRV (Rust 1.92).
  Why: the "pure managed code + native bridge" portability claim is untested off macOS.
  **M** — Acceptance: smoke render green on both runners; CI fails on regression.
- **Web client hardening** — error states on all four demo tabs (render / generate / any / compare
  in `DemoPage.tsx`), accessibility pass, `npm run build` + lint gate in CI.
  **M** — Acceptance: CI gate green; axe-clean on main flows.

## Phase 4 — v1 polish

**Depends on:** Phases 0–3. This phase is packaging, not capability.

- **NuGet package READMEs + icons + SourceLink** — 0.1.0 is published unlisted (`PptxEditor.Core`
  carries `<Version>0.1.0</Version>`); listing requires storefront hygiene.
  **S** — Acceptance: packages pass NuGet validation, deterministic + source-linked.
- **API versioning strategy** — URL or header versioning before external consumers depend on `/api/*`.
  **S** — Acceptance: versioned routes with the current surface as v1; unversioned redirects documented as temporary.
- **Performance pass** — establish a render-timings baseline (ms/slide via TypstBridge sessions)
  against a stated goal; profile the converter hot path. typst 0.15 brought no headline speedups,
  so current numbers stand (see `docs/typst-0.15.md` §4).
  **M** — Acceptance: numbers published in docs; >10% regression gate in CI or the benchmark tool.
- **Docs site or expanded `docs/`** — generation vocabulary reference, converter limits, theming
  guide.
  **M** — Acceptance: every public JSON vocabulary key documented with an example.
- **Migration guide 0.x → 1.0.**
  **S** — Acceptance: covers schema freeze deltas and any API renames.

## Out of scope (and why)

- **Animations/transitions** — Typst preview cannot represent them; dual-emit discipline breaks. Revisit if the preview path changes.
- **Video/audio embeds** — same reason; also outside the static-render mission.
- **PowerPoint template marketplace** — distribution problem, not an engineering one; brand profiles (`BrandProfileExtractor`) + generation themes cover the need.
- **Full chart-parity with PowerPoint rendering** — PowerPoint's chart engine is a moving target; we commit to editable native charts + approximate previews, not pixel-parity.
- **Collaborative editing** — the API deck-session model is single-writer; collaboration is a different product.

## Release gates

Everything in this roadmap targets **v0.5** — one version train. Within it: Phase 0 lands first
(nothing ships on unverifiable rendering), Phase 1 deepens what users touch (converter output),
Phase 2 deepens the authoring vocabulary, Phases 3–4 complete ecosystem and packaging (frozen
schema, verified platforms, listed packages). The schema freeze (Phase 2) is the release's
hard gate — v0.5 does not ship without it.

Post-v0.5: SmartArt depth beyond the documented matrix, additional chart types, web-client features beyond the demo, MCP surface expansion driven by user demand, typst bundle export (single-pass PDF+PNG+SVG) once usable via the crates.
