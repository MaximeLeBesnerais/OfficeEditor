# PPTX Roadmap — Progress Log

Branch: `agent/pptx-roadmap` | Worktree: `~/Work/DocxEditor-agent-pptx`

Pre-existing condition: 16 DocxEditor.Tests failures (TypstBridge native lib not compiled in this worktree). These are out of scope and documented here for transparency. All tests in scope (OfficeEditor.Api.Tests, PptxEditor.Core, OfficeEditor.Mcp.Tests) must stay green.

| Item | Status | Commit | Note |
|------|--------|--------|------|
| Phase 0: Fix OfficeEditor.Api.Tests image path resolution | DONE | 7160eea | `OoxmlEmitter.ResolveImageSource` handles worktree `.git` file detection for relative image paths |
| Phase 0: Regenerate pptx baselines | BLOCKED | — | Typst CLI and TypstBridge native lib unavailable in this environment. Cannot produce PDFs from PPTX. Needs typst binary or `--font-path` + native lib. |
| Phase 0: Un-skip remaining guarded tests | DONE | — | No file-based guards found. All env-gated tests use correct `OE_RUN_TYPST_COMPILE_TESTS` pattern for Typst-dependent tests. |
| Phase 0: Re-mine design tokens | DONE | 0ae199c | TokenMiner created; aetherlink.tokens.json regenerated from refreshed corpus. Round-trip test added. |
| Phase 0: OfficeEditor.Cli instrumentation | DONE | 0cdf946 | OfficeEditor.Cli.Tests project created with 20 command-level tests. Invokes Main() via reflection. |
| Phase 0: WebApplicationFactory smoke tests | DONE | 52cf85d | 22 integration tests via `WebApplicationFactory<Program>` covering all endpoint categories. |
| Phase 1: Split PptxToTypstConverter | DONE | 299e6cf | Moved Typst source generation methods (~1036 lines) to partial file `PptxToTypstConverter.SourceGeneration.cs`. Main file reduced from 4696 to 3660 lines. No behavior change, all tests green. Further splits (text extraction, table extraction, fonts) remain for future work. |
