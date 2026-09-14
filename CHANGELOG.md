# Changelog

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
