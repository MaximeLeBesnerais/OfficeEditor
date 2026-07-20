# TypstBridge Remaining Phases

## Goal

Track the remaining work after TypstBridge reached native PDF, SVG, and PNG support and became the primary backend for OfficeEditor.Core's `TypstCompilerService`.

## Current progress

Implemented in the bridge:

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
- OfficeEditor.Core `TypstCompilerService` prefers TypstBridge before fallback paths.
- External Typst CLI fallback remains available.
- Linux x64 native runtime asset generation can be triggered automatically by the managed project when the asset is missing.

## Scope boundaries

- Keep the external Typst CLI fallback as the safety net.
- Keep `typstsharp` as fallback/legacy unless a separate dependency-removal decision is made.
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

### 1. Integration validation review

- Confirm native and managed tests pass from a clean checkout after building the runtime asset.
- Confirm packaging scripts produce the expected `runtimes/<rid>/native` layout.
- Review font discovery and font precedence behavior for PPTX-generated Typst.
- Review diagnostics and fallback behavior in OfficeEditor.Core.

### 2. Backend hardening

- Keep TypstBridge as the preferred `TypstCompilerService` backend when the native library loads and probes successfully.
- Preserve the external Typst CLI fallback.
- Keep error messages clear when fallback paths are attempted.
- Keep existing callers unchanged.

### 3. Reference verification

- Run reference conversions for `examples/REF/PPTX/pres-pro.pptx` and `examples/REF/PPTX/AetherLink-Glass-Shareholder-Overview.pptx`.
- Verify PDF and PNG outputs before considering any fallback/dependency changes.

### 4. Dependency and packaging cleanup

- Remove `typstsharp` only if a separate dependency-removal effort is approved and fallback behavior remains acceptable.
- Verify publish output includes the correct native runtime asset.
- Expand and confirm the platform matrix before distributing packages.
