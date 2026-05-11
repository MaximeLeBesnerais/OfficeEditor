# TypstBridge

TypstBridge is the planned replacement for the current `typstsharp` dependency.

The goal is to provide a controlled native bridge around Typst's Rust crates, exposed through a stable C ABI and consumed from .NET through P/Invoke. This is not a C# reimplementation of Typst.

## Goals

- Replace `typstsharp` with a bridge we can build, package, debug, and harden ourselves.
- Support the complete existing compile surface: PDF, PNG, SVG, multi-page output, working directories, asset paths, font paths, PNG PPI, diagnostics, and safe memory ownership.
- Keep the external `typst` CLI fallback as a safety net.
- Compile Linux binaries with non-executable stack flags so the bridge works on hardened kernels.

## Planned Layout

```text
TypstBridge/
├── README.md
├── docs/
│   ├── implementation-plan.md
│   └── abi.md
├── native/      # Future Rust cdylib project
├── managed/     # Future C# P/Invoke wrapper
├── packaging/   # Native build and runtime asset scripts
└── tests/       # Future Rust, ABI, and managed tests
```

## Status

Foundation work only. Do not remove `typstsharp` until the native bridge supports PDF, PNG, and SVG end-to-end and passes reference PPTX conversion checks.

The first native implementation may only expose ABI/probe functionality while the full compile surface is built out.

## Packaging

Native runtime asset helpers live in [`packaging/`](packaging/):

```bash
TypstBridge/packaging/build-native.sh linux-x64
TypstBridge/packaging/pack-runtime-assets.sh linux-x64
```

On Windows:

```powershell
TypstBridge\packaging\build-native.ps1 -Rid win-x64
```

Scripts build or copy the Rust cdylib from `TypstBridge/native` into the .NET runtime asset layout:

```text
TypstBridge/runtimes/<rid>/native/<native-library>
```

Linux builds must use a non-executable stack (`-C link-arg=-Wl,-z,noexecstack`) so the bridge can load on hardened kernels.

See `docs/implementation-plan.md` for the full migration plan.
