# Editing generation sources (0.10)

Stable string IDs identify elements in saved DOCX, PPTX and XLSX generation JSON.
Save normalized source before reordering it. `deck_generate` and
`POST /api/decks/generate` now return `normalizedJson` for this purpose.

## CLI

```sh
officeeditor ids ./documents --recursive --dry-run
officeeditor ids ./documents --recursive --backup
officeeditor inspect document.json
officeeditor edit document.json --instructions edits.json --output edited.json
officeeditor generate edited.json --output edited.docx
```

`inspect` prints an array with each element's ID, kind, JSON path, parent ID and
properties. For ID-less input these are the IDs normalization would assign.
`edit` requires a new output path and never overwrites an existing output file.
The source is unchanged, including when a later operation fails.

`ids --dry-run` lists files needing IDs without writing files or backups.
`--backup` saves the original bytes to `<file>.bak` for each changed file and
refuses an existing backup. The batch checks IDs and backup conflicts before
writing; filesystem errors during writing can still leave some files updated.
Completed files are left untouched. Directory scans exclude build folders and
symlinks, as described in [element IDs](element-ids.md).

## Edit operations

`edits.json` is an array, evaluated in order:

```json
[
  {"type":"rename", "id":"paragraph-0", "newId":"intro"},
  {"type":"set", "id":"intro", "properties":{"text":"Updated introduction"}}
]
```

`set` replaces the supplied properties and retains all others. Objects and arrays
are replaced whole, not merged recursively. `rename` changes identity; subsequent
operations must use the new ID. Unknown operations, unknown operation fields,
missing targets, duplicate IDs and invalid IDs fail the entire batch. Errors
identify the operation index. An empty array normalizes IDs and returns the source.

The editor validates operation syntax and identity. It does **not** validate the
format vocabulary or render the document. Run the format generator to validate
properties, layout and assets. For example, changing a text property's value is
valid for a paragraph, but adding an unknown property fails at generation time.

## HTTP and MCP

Send this envelope to `POST /api/documents/edit-source` or call the MCP tool
`document_edit_source` with the same arguments:

```json
{
  "document": {
    "version":"1.0",
    "sections":[{"blocks":[{"type":"paragraph","id":"intro","text":"Before"}]}]
  },
  "operations":[{"type":"set","id":"intro","properties":{"text":"After"}}]
}
```

Both return `document` (the edited JSON object) and `elements` (the inventory).
Persist `document` and pass it to the format generator. For PPTX, use
`POST /api/decks/generate` or `deck_generate` to create a new deck session.
The source endpoint is stateless and does not mutate an existing binary/session.
Invalid requests return HTTP 400 or MCP InvalidParams.

## Identity contract and coverage

Coverage is unchanged from 0.9: PPTX slides and recursive children; DOCX sections,
headers, footers, blocks and positioned elements; XLSX worksheets, cells, columns
and tables. Formatting objects, text runs and primitive values have no IDs.
See [element IDs](element-ids.md) for details.

Reordering saved elements retains their IDs. Deleting an element removes its
identity. When copying an element, assign fresh IDs to it and all copied children;
copying IDs unchanged fails uniqueness checks. When moving elements across source
documents, resolve destination ID collisions explicitly. There are no automatic
copy/delete/move operations in this release.

Source IDs are distinct from numeric OOXML IDs. Existing `deck_replace_element`
operations retain their numeric, slide-local contract. Source identities cannot
be reconstructed from an arbitrary uploaded Office file.

## .NET

Use `GenerationJsonEditor.Apply(sourceJson, operationsJson)` from
`OfficeEditor.Core.Generation`. It returns normalized source only after the whole
batch succeeds and never modifies caller input. Inspect the result using
`GenerationJsonIds.Inspect` and generate it with the relevant format library.
