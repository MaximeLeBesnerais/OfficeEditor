# Element IDs

Generation JSON accepts an `id` on each authored element. Use IDs to find and edit
content without relying on text matches, array positions or Office drawing numbers.

```json
{
  "type": "text",
  "id": "experience-heading",
  "text": "Professional experience",
  "size": { "w": 300, "h": 40 },
  "at": { "x": 40, "y": 100 }
}
```

IDs start with a lowercase letter and contain lowercase letters, digits, hyphens
or underscores (`[a-z][a-z0-9_-]*`). They are unique across the entire source
document, including slides/sections/worksheets. Invalid and duplicate IDs fail
validation. IDs can be renamed in JSON; renaming deliberately changes identity.

Missing IDs are generated using readable type names: `slide-0`, `text-0`,
`paragraph-0`, `textbox-0`, `worksheet-0`, `cell-0`, etc. Existing explicit names
are reserved before defaults are allocated. No existing ID is renumbered when
another element is inserted, deleted or reordered in saved JSON.

## Persisting IDs

`officeeditor generate` automatically writes missing IDs back to its input JSON.
It does not otherwise change the document's values. To update existing sources
without generating output:

```sh
officeeditor ids deck.json
officeeditor ids ./documents --recursive
```

The directory form updates recognized OfficeEditor generation documents and skips
unrelated configuration JSON and malformed JSON. It excludes `.git`, `bin`, `obj`,
`node_modules` and symlinks. A duplicate/invalid ID in a recognized document stops
batch preflight before any file is written. Re-running the command is a no-op for
files that already have IDs. Missing IDs are inserted without reformatting other
properties, text, numeric literals or newlines.

String-based APIs cannot change their caller's string. Their successful validation
results and JSON-to-bytes generation results expose `NormalizedJson`; persist that
value when updating a source document. Read-only validation/render calls do not
write files. Programmatic model inputs can carry explicit `Id` values, but do not
have original JSON to return.

## Lookup, rename and edit from C#

```csharp
using System.Text.Json.Nodes;
using OfficeEditor.Core.Generation;

string json = GenerationJsonIds.Normalize(File.ReadAllText("deck.json")).Json;
var element = GenerationJsonIds.GetElement(json, "experience-heading");
// element.Id, Kind, ParentId, Path and detached Properties are available.

json = GenerationJsonIds.SetProperties(json, element.Id, new JsonObject
{
    ["text"] = "Professional Experience",
    ["at"] = new JsonObject { ["x"] = 60, ["y"] = 120 }
});
json = GenerationJsonIds.Rename(json, element.Id, "career-heading");
File.WriteAllText("deck.json", json);
```

`Inspect(json)` returns all authored elements in tree order. Property edits replace
each supplied property whole: edit `runs` explicitly to change rich text, and use
the format's own geometry vocabulary (`at`/`size` for PPTX, `x`/`y`/`width`/`height`
for positioned DOCX, `address` for an XLSX cell). The helper checks identity;
the format parser/generator still validates the edited content and geometry.

## Coverage and limits

- PPTX: slide roots and nested containers, groups, primitives and authored
  components/archetypes. Synthesized component children keep the existing layout
  pipeline's derived IDs; they are not independent authored JSON elements.
- DOCX: sections and their header, footer, body/group/semantic-section blocks and
  positioned primitives. Tables/lists are editable blocks; individual runs, list
  entries and table cells are not separate source IDs in this version.
- XLSX: worksheets, explicit cells, column definitions and table definitions.
  Primitive `rows` values and header strings remain unchanged; use explicit
  `cells` when individual ID-based selection is needed.

Design tokens, metadata, styles and other property dictionaries are not elements.
Starting afresh from ID-less JSON after reordering cannot preserve prior identity:
keep the normalized source. Copying an element requires a new ID; duplicate names
are deliberately rejected. This field identifies source JSON, not an arbitrary
Office file reopened without its source. Native-file identity recovery and GUI
integration are separate concerns.

Implementation outline: [stable-element-identity.md](plans/stable-element-identity.md).
