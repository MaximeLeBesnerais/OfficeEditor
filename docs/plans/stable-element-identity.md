# Element IDs in generation JSON

Scope updated after user clarification: add a small, shared JSON identity layer,
not a Studio integration or a document/session persistence system.

## Contract

- Add `id` to authored elements: PPTX slides and their element trees; DOCX sections,
  header/footer/body blocks and positioned elements; XLSX worksheets and explicit
  cells, columns and tables. Design/style dictionaries and primitive row values
  are not elements and do not acquire IDs.
- Preserve every explicit ID. Generate readable defaults (`text-0`, `paragraph-0`,
  `worksheet-0`, etc.) using per-type counters, reserving all explicit names first.
- Use the existing PPTX syntax `[a-z][a-z0-9_-]*` and document-wide uniqueness.
  Renaming in JSON is supported; invalid/duplicate names fail loudly.
- Persist generated IDs in source JSON. Once saved, insertion/reorder/content edits
  do not change existing IDs. Starting again from ID-less JSON cannot preserve
  identities across reorder; file normalization prevents that common mistake.
- Do not change rendering or conflate source IDs with Office numeric drawing IDs.

## Implementation outline

1. Shared normalization and ID lookup/rename/property-edit helpers in
   OfficeEditor.Core. Preserve unrelated JSON and make repeated normalization a
   no-op. Expose explicit file normalization with atomic writes.
2. DOCX/XLSX parser and model support matching existing PPTX behavior: every parsed
   element has an ID, and errors remain in each format's validation contract.
3. File-based CLI generation adds missing IDs back to its input. Add an `ids`
   command for individual files or recursive migration of a directory, skipping
   unrelated JSON. Update recognized generation JSON already tracked in this repo.
4. Tests: defaults, explicit-name collisions, duplicate/malformed IDs, rename,
   ID-based edits, reorder stability, nested content, all three formats, idempotent
   files and CLI behavior. Run the full .NET suite and reference smoke checks.

## Deliberate boundaries

String APIs cannot modify the caller's string: use the normalizer's returned JSON
and persist it, or use file normalization/CLI generation. Read-only render, inspect
and validation calls do not unexpectedly write files. IDs are initially authored
JSON identity; automatic recovery from Office files, arbitrary external Office
edits, undo/redo and GUI/MCP integrations are outside this change. Duplicate/copy
operations must assign fresh IDs to copied elements; duplicate IDs are not repaired
silently. Renaming deliberately changes identity and callers must update selections.
