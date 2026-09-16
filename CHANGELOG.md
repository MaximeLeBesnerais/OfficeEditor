# Changelog

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
