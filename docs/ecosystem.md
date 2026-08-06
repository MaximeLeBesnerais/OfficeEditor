# OfficeEditor ecosystem

OfficeEditor is the engine. Around it sit two satellite projects, both out of this repository, plus this repo's own docs. This page maps the ecosystem and cross-links the documentation.

## In this repo

| Doc | Covers |
|---|---|
| [Declarative PPTX generation](pptx-generation.md) | The `version: "2.0"` deck vocabulary — design tokens, container tree, primitives, components, archetypes, speaker notes, image sources. |
| [Declarative DOCX generation](docx-generation.md) | The `version: "1.0"` document vocabulary — flow + positioned tiers, themes, semantic report archetypes. |

The canonical deck schema lives at [`PptxEditor.Core/Generation/Schema/deck.schema.json`](../PptxEditor.Core/Generation/Schema/deck.schema.json) and is published as `https://officeeditor.dev/schemas/deck-2.0.json` so `$id` references resolve (see officeeditor.dev below).

## OfficeEditorStudio

**Location:** `~/Work/OfficeEditorStudio` (separate git repo, not part of this repository)

A local-first, self-sufficient **desktop document studio** for decks (PPTX) and docs (DOCX) built on the OfficeEditor engine. The document **is** the generation JSON — visual editor, JSON editor, and AI agent all manipulate the same model.

Key facts:

- **Stack:** Avalonia 12.1.1 (MVVM via CommunityToolkit.Mvvm), Skia-rendered canvas, `net10.0`. Cross-platform Win/macOS/Linux, no webview shell.
- **Engine consumption:** in-proc via the published NuGet packages only — `MaximeLB.PptxEditor.Core`, `MaximeLB.DocxEditor.Core`, `MaximeLB.OfficeEditor.Core`, `MaximeLB.TypstBridge.Managed`, all `0.7.1`. **No project/repo references**; the studio is a consumer of the published product. XLSX packages are deliberately excluded (no sheets in the product).
- **Document format:** the generation JSON — `.deck.json` (PPTX vocabulary v2.0) and `.doc.json` (DOCX vocabulary v1.0). Every UI action is a validated patch to it; undo/redo is patch history; save is the JSON file itself (versionable, diffable).
- **Rendering:** fully in-app via embedded TypstBridge — SVG (interactive preview, warm whole-deck ≈ 2–4 ms/slide), PNG on demand, PDF export. No Office, no LibreOffice, no browser, no cloud dependency.
- **AI-native (BYOK):** OpenAI-compatible chat completions (OpenAI / OpenRouter / Ollama / LM Studio). The model emits schema-conforming JSON patches; validation errors (field path + suggestion) are fed back verbatim for self-correction — validation-as-feedback is treated as an AI feature. A **vision loop** renders the current slide/page PNG (8–48 ms), attaches it to the model turn, and lets the model critique its own layout — only offered for vision-capable models (flagged from the cached models.dev catalog).
- **Planned MCP server mode:** expose the open document to external agents with the same tool family as this repo's MCP host (`deck_anatomize` / `deck_replace_element` / `deck_render_slide` / `deck_generate`) pointed at the live session — human and agent share one document, one model, one validation path.

### SPEC.md summary

`~/Work/OfficeEditorStudio/SPEC.md` (approved direction, **pre-implementation scaffold** — Avalonia template + engine packages restored, no product code yet) defines:

- **Two studios on one skeleton:** a **deck studio** (slide list / canvas / inspector three-pane) and a **doc studio** (page-based preview, flow blocks + positioned tier).
- **Structured mode** (default): archetype & block palette; drag reorders children (document order is paint order).
- **Canvas mode** (per slide): layout-less container; drag moves/resizes via `at`/`size` patches; full direct manipulation (selection, multi-select, handles, snap-to-grid).
- **Variables panel:** detect `{{var}}` (incl. `{{name|default}}`) → fill-in form → live merged preview.
- **JSON view:** full source with schema validation, two-way synced with the visual editor (patch log is the merge mechanism).
- **Validation panel:** field-path errors + suggestions, click-to-locate.
- Packaging: self-contained per-OS bundles (win-x64, osx-arm64, linux-x64) with bundled TypstBridge runtime assets; macOS signing + notarization in CI.
- Non-goals (v1): XLSX studio, animations/transitions, SmartArt editing, charts editor, in-canvas text editing, real-time collaboration, PowerPoint/Word feature parity.

## officeeditor.dev

**Location:** `~/Work/officeeditor.dev` (separate repo)

The landing site for the `officeeditor.dev` domain.

- **Stack:** Vite 7 + React 19 + TypeScript + Tailwind CSS v4 (`@tailwindcss/vite`), Shiki syntax highlighting, self-hosted Inter / JetBrains Mono fonts; built with `bun`.
- **Schema hosting:** serves the PPTX generation vocabulary schema at `public/schemas/deck-2.0.json` so its `$id` (`https://officeeditor.dev/schemas/deck-2.0.json`) resolves from anywhere. A post-build checksum guard (`scripts/verify-schema.mjs`) fails the build if the schema does not survive the Vite build verbatim (public → dist). Note: the published copy may lag the repo's canonical `deck.schema.json` (it currently predates the `notes` speaker-notes property).
- **Real renders:** the slide images under `public/assets/slides/` are genuine OfficeEditor output — the demo deck → PPTX → PNG via TypstBridge, converted to WebP.
- **Docs planned:** the site currently introduces the product and hosts schemas/renders; documentation sections are planned.

## Related projects

- This repo (the engine): README, [Declarative PPTX generation](pptx-generation.md), [Declarative DOCX generation](docx-generation.md).
- `OfficeEditorStudio` — desktop studio consuming the `MaximeLB.*` 0.7.1 packages (see above).
- `officeeditor.dev` — landing site + schema hosting (see above).
