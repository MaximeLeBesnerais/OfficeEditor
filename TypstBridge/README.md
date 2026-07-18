# TypstBridge

TypstBridge is a first-party native bridge around Typst's Rust crates, exposed to .NET through a stable C ABI and a managed P/Invoke wrapper.

It is not a C# reimplementation of Typst. OfficeEditor.Core's `TypstCompilerService` now tries TypstBridge first and falls back to the legacy paths when the bridge is unavailable or compilation fails.

## Current status

- Native Rust `cdylib` implemented under [`native/`](native/).
- Managed .NET 9 P/Invoke wrapper implemented under [`managed/`](managed/).
- Native rendering is implemented for PDF, SVG, and PNG.
- Source strings are compiled through a native Typst world.
- Relative assets resolve from the request working directory.
- Explicit font file/directory paths are supported.
- PDF returns one output; SVG and PNG return one output per page.
- PNG rendering accepts a PPI value.
- Diagnostics and error messages are exposed through the ABI and managed wrapper.
- OfficeEditor.Core uses TypstBridge as the primary `TypstCompilerService` backend.
- The external Typst CLI remains the safety-net fallback.
- `typstsharp` remains present as a legacy PDF-only fallback path and has not been removed.
- Linux x64 builds can auto-generate the native runtime asset when `TypstBridge.Managed` is built and the asset is missing.
- Managed tests cover native loading, PDF/SVG/PNG rendering, multi-page outputs, PPI, and diagnostics. Asset/font behavior and repeated compile/free stability are covered by native-side checks or remain future managed-test coverage where gaps exist.
- **ABI v3** adds persistent compile sessions (warm worlds), a process-wide parsed-font cache, and bounded comemo memoization eviction for long-lived processes; see below and [`docs/abi.md`](docs/abi.md).

## Performance behavior (ABI v3)

- **Native font cache.** Parsed font sets (`FontBook` + `Font` handles) are memoized process-wide, keyed by the sorted font-path list plus working directory and font-flag state. A cache hit skips directory scanning, `fs::read`, and font parsing entirely. The cache is never evicted; its size is bounded by the number of distinct font-path lists the process compiles with.
- **Persistent sessions.** `TypstBridgeCompiler.CreateSession(workingDirectory, fontPaths)` returns a `TypstBridgeCompileSession` that keeps the native world (library, font set, working directory) warm. `UpdateSource(source, rootFileName)` swaps the source in place; `Compile(outputFormat, ppi)` recompiles with the same result layout as a cold compile; `Dispose()` frees the native handle (finalizer-safe via `SafeHandle`). Session output is byte-identical to a cold `Compile` of the same inputs. Sessions are independent and compile in parallel; a single session serializes its own calls.
- **Memoization eviction.** Every 10 finished compiles the bridge runs one `comemo::evict(30)` pass (typst-cli watch-mode pattern), reclaiming memoized entries unused for roughly 300 compiles so server memory stays bounded. `TypstBridgeCompiler.EvictCache(maxAge)` forces a pass on demand; `EvictCache(0)` evicts everything unreferenced. Output bytes are deterministic across eviction boundaries (`World::today` is pinned at 1970-01-01).

Current limits:

- `typstsharp` is still in the repository for fallback/legacy behavior.
- Package/platform support is preliminary; the runtime asset layout exists, but the full platform matrix still needs verification.

## Layout

```text
TypstBridge/
├── README.md
├── docs/
│   ├── implementation-plan.md
│   └── abi.md
├── native/      # Rust cdylib project
├── managed/     # C# P/Invoke wrapper
├── packaging/   # Native build and runtime asset scripts
└── tests/       # Managed bridge tests and fixtures
```

## Build native runtime asset

From the repository root, build and copy the native library into the .NET runtime asset layout:

```bash
TypstBridge/packaging/build-native.sh linux-x64
```

On Windows:

```powershell
TypstBridge\packaging\build-native.ps1 -Rid win-x64
```

The generated asset is copied to:

```text
TypstBridge/runtimes/<rid>/native/<native-library>
```

For Linux x64 this is:

```text
TypstBridge/runtimes/linux-x64/native/libtypst_bridge.so
```

Linux builds use `-C link-arg=-Wl,-z,noexecstack` so the bridge can load on hardened kernels.

On Linux x64, building `TypstBridge.Managed` or a project that references it can run the native build automatically when `TypstBridge/runtimes/linux-x64/native/libtypst_bridge.so` is missing. Clean source builds on Linux x64 therefore require Rust and `cargo` unless the runtime asset is already present.

If the native library has already been built, copy it without rebuilding:

```bash
TypstBridge/packaging/pack-runtime-assets.sh linux-x64
```

Generated runtime assets under `TypstBridge/runtimes/` are build outputs for local test/package validation and should not be committed unless a packaging decision explicitly changes that policy.

## Run managed tests

Build the native runtime asset for your RID first, then run:

```bash
dotnet test TypstBridge/tests/TypstBridge.Managed.Tests/TypstBridge.Managed.Tests.csproj
```

The managed project includes `TypstBridge/runtimes/**/*` as runtime assets, so tests should load the native library without manually setting `LD_LIBRARY_PATH` when the asset exists for the current RID.

## Supported compile request surface

- `source`: Typst source string.
- `workingDirectory`: base directory for relative imports/assets such as `assets/...`.
- `rootFileName`: virtual root file name used by the source-backed world.
- `fontPaths`: explicit font files or directories.
- `outputFormat`: `Pdf`, `Svg`, or `Png`.
- `ppi`: PNG rasterization density.
- diagnostics: structured diagnostics plus human-readable messages.

Expected outputs:

| Format | Output behavior |
| --- | --- |
| PDF | One output item containing the PDF bytes. |
| SVG | One output item per page, ordered by page index. |
| PNG | One output item per page, ordered by page index, rendered at the requested PPI. |

## Packaging

Native runtime asset helpers live in [`packaging/`](packaging/):

```bash
TypstBridge/packaging/build-native.sh linux-x64
TypstBridge/packaging/pack-runtime-assets.sh linux-x64
```

See [`packaging/README.md`](packaging/README.md) for script details and the preliminary RID matrix. See [`docs/abi.md`](docs/abi.md) for the C ABI shape.
