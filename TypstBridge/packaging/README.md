# TypstBridge Packaging

This directory contains build helpers for packaging the TypstBridge native Rust
library as .NET runtime assets.

The scripts prepare the native runtime layout consumed by the managed P/Invoke
wrapper, managed tests, and OfficeEditor.Core through its `TypstBridge.Managed`
reference. They do not provide full platform packaging or
replace the external Typst CLI fallback.

TypstBridge renders PDF, SVG, and PNG through the native Rust cdylib and is the
primary backend for OfficeEditor.Core's `TypstCompilerService`. The
platform/package matrix is still preliminary. Release packages currently build and
publish **osx-arm64, linux-x64, and win-x64**. The other layouts below are local-build
or future-matrix conveniences and are not release-verified promises.

## Runtime asset layout

Native libraries are copied under `TypstBridge/runtimes/<rid>/native`:

```text
TypstBridge/runtimes/
├── linux-x64/native/libtypst_bridge.so
├── linux-arm64/native/libtypst_bridge.so
├── win-x64/native/typst_bridge.dll
├── win-arm64/native/typst_bridge.dll
├── osx-x64/native/libtypst_bridge.dylib
└── osx-arm64/native/libtypst_bridge.dylib
```

The Bash build script supports Linux and macOS x64/arm64 targets; the PowerShell
build script supports `win-x64` and `win-arm64`. The release workflow publishes
only the three release-verified RIDs listed above.

Generated files under `TypstBridge/runtimes/` are local build outputs for test
and package validation. Do not commit them unless the packaging policy is
explicitly changed.

On a recognized host RID, building `TypstBridge.Managed` (or a project that references
it) can build the matching native asset when it is missing and the build is not
cross-targeting another RID. Clean source builds therefore require Rust, `cargo`, and
the host linker when the runtime asset is not already present; otherwise the managed
service probes the bridge and can fall back to the external Typst CLI.

## Build native library

From the repository root or any other directory:

```bash
TypstBridge/packaging/build-native.sh linux-x64
```

The script:

1. Locates `TypstBridge/native/Cargo.toml` relative to the script path.
2. Maps the RID to an explicit Rust target (`x86_64-unknown-linux-gnu` or
   `aarch64-unknown-linux-gnu`) and runs `cargo build --release --target`.
3. Copies `native/target/<rust-target>/release/libtypst_bridge.so` to
   `TypstBridge/runtimes/<rid>/native/`.

On Windows PowerShell:

```powershell
TypstBridge\packaging\build-native.ps1 -Rid win-x64
```

This maps the RID to an explicit Rust target (`x86_64-pc-windows-msvc` or
`aarch64-pc-windows-msvc`) and copies
`native\target\<rust-target>\release\typst_bridge.dll` to
`TypstBridge\runtimes\<rid>\native\`.

Both build scripts fail with clear messages if `cargo` is not available, if the
expected native crate is missing, or if the expected artifact is missing. The
required Rust target and linker toolchain must be installed for the selected RID.

After building the runtime asset for the current RID, run managed verification
with:

```bash
dotnet test TypstBridge/tests/TypstBridge.Managed.Tests/TypstBridge.Managed.Tests.csproj
```

## Pack an existing artifact

If the native library has already been built, copy it into the runtime layout
without rebuilding:

```bash
TypstBridge/packaging/pack-runtime-assets.sh linux-x64
```

This script only copies already-built Linux artifacts from the explicit Rust
target output directory and does not require `cargo` to be installed.

## Linux noexecstack requirement

Linux builds must not request an executable stack. The Bash build script adds:

```text
-C link-arg=-Wl,-z,noexecstack
```

This is required for hardened Linux kernels and avoids native
loading failures on restricted systems. When available, verify a Linux
artifact with:

```bash
readelf -W -l TypstBridge/runtimes/linux-x64/native/libtypst_bridge.so
```

The `GNU_STACK` program header should not include the executable flag.
