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
├── packaging/   # Future native build and runtime asset scripts
└── tests/       # Future Rust, ABI, and managed tests
```

## Status

Planning only. Do not remove `typstsharp` until the native bridge supports PDF, PNG, and SVG end-to-end and passes reference PPTX conversion checks.

See `docs/implementation-plan.md` for the full migration plan.
