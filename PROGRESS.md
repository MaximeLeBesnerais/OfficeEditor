# PPTX Roadmap — Progress Log

Branch: `agent/pptx-roadmap` | Worktree: `~/Work/DocxEditor-agent-pptx`

Pre-existing condition: 16 DocxEditor.Tests failures (TypstBridge native lib not compiled in this worktree). These are out of scope and documented here for transparency. All tests in scope (OfficeEditor.Api.Tests, PptxEditor.Core, OfficeEditor.Mcp.Tests) must stay green.

| Item | Status | Commit | Note |
|------|--------|--------|------|
| Phase 0: Fix OfficeEditor.Api.Tests image path resolution | DONE | TBD | `OoxmlEmitter.ResolveImageSource` now handles worktree `.git` file detection for resolving relative image paths |
