# TypstBridge Remaining Phases

## Goal

Track the remaining work after the standalone TypstBridge rendering bridge reached native PDF, SVG, and PNG support through the Rust cdylib and managed P/Invoke wrapper.

## Current progress

Implemented in the standalone bridge:

- Native Rust `cdylib` under `TypstBridge/native`.
- ABI exports for version, probe, compile, result free, string free, and last error.
- Source string compilation with working-directory asset resolution.
- Explicit font file/directory paths.
- Native PDF, SVG, and PNG rendering.
- One PDF output; ordered per-page SVG/PNG outputs.
- PNG PPI support.
- Structured diagnostics and managed error mapping.
- Managed .NET 9 P/Invoke wrapper and managed tests.
- Packaging scripts and runtime asset layout foundation.

## Scope boundaries

- OfficeEditor/PptxEditor integration is still pending.
- Keep the external Typst CLI fallback as the safety net until native parity is demonstrated in the integration layer.
- Do not remove `typstsharp` until native integration passes reference PPTX conversion checks.
- Treat the package/platform matrix as preliminary until each RID is built and verified.
- Generated runtime assets under `TypstBridge/runtimes/` are local build outputs and should not be committed unless packaging policy changes.
- Do not commit unless explicitly requested by the user.

## Verification commands

Build the native runtime asset for the current RID:

```bash
TypstBridge/packaging/build-native.sh linux-x64
```

Run managed bridge tests:

```bash
dotnet test TypstBridge/tests/TypstBridge.Managed.Tests/TypstBridge.Managed.Tests.csproj
```

Optional native checks:

```bash
cargo test --manifest-path TypstBridge/native/Cargo.toml
```

## Remaining milestones

### 1. Integration readiness review

- Confirm native and managed tests pass from a clean checkout after building the runtime asset.
- Confirm packaging scripts produce the expected `runtimes/<rid>/native` layout.
- Review font discovery and font precedence behavior for PPTX-generated Typst.
- Review diagnostics and fallback expectations before touching OfficeEditor/PptxEditor.

### 2. OfficeEditor/PptxEditor backend plan

- Design backend selection for `TypstCompilerService`.
- Prefer TypstBridge when the native library loads and probes successfully.
- Preserve the external Typst CLI fallback.
- Define error messages that identify attempted backends.
- Keep existing callers unchanged.

### 3. Integration implementation and reference verification

- Wire the managed bridge into the OfficeEditor/PptxEditor compile path.
- Do not remove `typstsharp` during initial integration.
- Run reference conversions for `examples/REF/Presentation1.pptx` and `examples/REF/pres-pro.pptx`.
- Verify PDF and PNG outputs before considering dependency removal.

### 4. Dependency and packaging cleanup

- Remove `typstsharp` only after native integration is accepted.
- Verify publish output includes the correct native runtime asset.
- Expand and confirm the platform matrix before distributing packages.
