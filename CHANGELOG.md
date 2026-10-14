# Changelog

## [0.10.1](https://github.com/MaximeLeBesnerais/OfficeEditor/compare/v0.10.0...v0.10.1) — 2026-10-14

**Not released.** This section describes the `dev` branch after v0.10.0; nothing here is tagged or published. To cut it, merge the tested commit into `main`, bump the README release/NuGet markers and `TypstBridge/native/Cargo.toml` to 0.10.1, then tag `v0.10.1` — `release.yml` now refuses to publish a tag whose tree does not document that version (the v0.10.0 tag shipped a README that still told consumers to install 0.7.1).

### Added

- `officeeditor create <out.docx|out.pptx|out.xlsx> --instructions <generation.json>` generates the document from a generation source instead of a blank file, through the same validation and rendering path as `generate`. The instruction file is never rewritten, and a rejected source cannot leave a blank document behind.
- `officeeditor generate --no-write-ids` generates from the in-memory normalized copy and leaves the source byte-identical; `--backup` saves the pre-normalization bytes to `<file>.bak` before missing IDs are written back (the two flags are mutually exclusive).
- `generate --output <out.pdf|out.png|out.svg>` now renders by source vocabulary: the root collection (`slides` for PPTX, `sections` for DOCX, `worksheets` for XLSX) selects the generator before rendering, so PPTX deck JSON is no longer routed to the XLSX renderer.
- Symbol packages: every `.nupkg` ships a matching `.snupkg`, and each package embeds the repository README and third-party notices.
- CI: fail-closed change detection (a diff is only treated as docs-only once positively proven), an opt-in Typst render-path leg, and an `fr_FR.UTF-8` locale leg.
- Release gates in `release.yml`: tag format, `main` ancestry, README/CHANGELOG/`TypstBridge/native/Cargo.toml` version agreement, one `.snupkg` per `.nupkg`, and a smoke test of the packaged CLI.

### Changed

- `make dev` now starts the API (http://localhost:5001) and the Vite web client (http://localhost:5173) together through the client's `dev:all` script; `make api` and `make web` remain single-surface targets.
- `Directory.Build.props` is the single source of `<Version>` (0.10.1) for every project; projects no longer declare their own, and the release workflow stamps the tag version over it at pack time.
- The MCP `initialize` result and startup log report the assembly version instead of a hardcoded literal, so the server cannot drift from the package again.
- Diagnostics that quote numbers render culture-invariantly — layout overflow/shrink warnings, image DPI and text-fit warnings, positioned-element geometry errors, XLSX range/width/height errors, parity threshold errors, and the LibreOffice timeout message — so a non-en-US locale no longer changes their text. Regression tests pin the behaviour with a deliberately non-English culture instead of relying on the runner's locale.

### Fixed

- MCP server version drift: the host reported `0.7.0` on a 0.10.0 tree.
- CLI help text now documents `create --instructions`, `generate --no-write-ids`/`--backup`, and PDF/PNG/SVG outputs.

### Documentation

- Chart converter README: clustered, stacked and percent-stacked bar/column charts plus pie charts and doughnuts are implemented (line charts remain placeholders, doughnut holes stay transparent); the stale "stacked is a placeholder" table and the pie/doughnut claims were corrected.
- `examples/Xlsx/README.md`: `create --instructions` and the unified `edit` command are documented as implemented instead of stubs.
- Visual-diff suite lists now name the real set (`docx|pptx|xlsx|gen`) and the DOCX suite's reference-output requirement; AGENTS.md lists all six test projects.
- Documented the Rust/`cargo` build prerequisite: a fresh clone's first `dotnet build` compiles the native TypstBridge (the host-RID asset under `TypstBridge/runtimes/` is gitignored), and `build-native.sh` fails without cargo.

## [0.10.0](https://github.com/MaximeLeBesnerais/OfficeEditor/compare/v0.9.0...v0.10.0) — 2026-09-16

### Added

- Shared sequential `GenerationJsonEditor.Apply` batches for stable-ID property edits and renames across DOCX, PPTX and XLSX generation sources. Failed batches return no partial result.
- CLI `inspect` and source JSON `edit` commands; edit output must be a new file.
- Stateless `POST /api/documents/edit-source` endpoint and `document_edit_source` MCP tool, returning edited source and an element inventory.
- `ids --dry-run` and `ids --backup`, changed-file reporting and existing-backup protection.
- Source editing, generation round-trip, transport and migration regression tests.

### Changed

- PPTX generation API/MCP responses include `normalizedJson` so callers can persist stable identities.
- Release workflow smoke-tests the packaged CLI before publishing.

### Compatibility

- Source editors validate operation syntax and IDs; format generators validate vocabulary and rendering.
- Source IDs do not become OOXML IDs. Existing numeric `deck_replace_element` behavior is unchanged.
- ID coverage is unchanged; see [source editing](docs/source-editing.md) for coverage and copy/move semantics.
- All six NuGet packages share version 0.10.0.

### Validation

- 3,368 local tests passed, including opt-in Typst compilation.
- All seven reference-document PDF conversions passed.
- Installed 0.10.0 CLI package passed migration, inspect, edit and XLSX generation smoke checks.

## [0.9.0](https://github.com/MaximeLeBesnerais/OfficeEditor/compare/v0.8.0...v0.9.0) — 2026-09-15

### Added

- Stable, editable `id` fields on supported elements in DOCX, PPTX, and XLSX generation JSON. Missing IDs receive readable defaults such as `paragraph-0` and `cell-0`; existing IDs are preserved and reserved before defaults are assigned.
- Shared `GenerationJsonIds` APIs to normalize documents, inspect elements, find an element by ID, rename it, and update its properties. Duplicate or invalid IDs produce explicit errors.
- `NormalizedJson` on successful JSON parsing and generation results, so callers can persist automatically assigned IDs.
- `officeeditor ids <file.json|directory> [--recursive]` to upgrade existing generation sources. Unrelated JSON is skipped, and recognized documents are validated for ID conflicts before the batch is written.
- [Element ID documentation](docs/element-ids.md) and an [implementation plan](docs/plans/stable-element-identity.md).

### Changed

- CLI `generate` writes missing IDs back to its input generation JSON before generating the document. Running it again leaves complete IDs unchanged.
- Migrated 28 example, demo, and fixture JSON documents with 828 IDs, preserving their existing values.

### Compatibility and migration

- Existing generation JSON without IDs remains accepted. IDs use lowercase letters, digits, underscores, and hyphens, start with a letter, and must be unique within the document.
- Persist normalized JSON to keep automatically assigned identities across edits. Renaming an ID changes its identity; copied elements need unique IDs.
- IDs identify supported source JSON elements. They do not add identity metadata to exported Office files or identify every nested formatting property, text run, or primitive value. See the documentation for coverage by format.

### Validation

- 3,351 tests passed locally, including the opt-in Typst compilation tests.
- All five PPTX and both DOCX reference-document smoke checks passed.

All six NuGet packages share version **0.9.0**, including `MaximeLB.OfficeEditor.Cli` and `MaximeLB.TypstBridge.Managed`.
