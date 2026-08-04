# Security Policy

## Supported versions

OfficeEditor is pre-1.0. Security fixes are applied to the latest code on `main` and, when appropriate, released in the next NuGet version. The latest listed NuGet release is `0.7.1`; older snapshots and unlisted builds are not supported.

| Version | Supported |
|---|---|
| Latest `main` | Yes |
| `0.7.x` NuGet packages | Yes |
| Older or private builds | No |

## Reporting a vulnerability

Please **do not open a public issue** for a suspected vulnerability.

1. Email **maxime@lebesnerais.com** with the subject `OfficeEditor security report`. After private vulnerability reporting is enabled for the public repository, you may instead use [GitHub Security Advisories](https://github.com/MaximeLeBesnerais/OfficeEditor/security/advisories/new).
2. Include the affected package and version or commit, impact, reproduction steps, and any suggested mitigation.
3. Avoid attaching confidential Office documents; use a minimal synthetic file whenever possible.

You should receive an acknowledgement after the report is reviewed. A fix, advisory, and release will be coordinated according to severity and exploitability. Please allow time for a patch before publishing details.

## Product trust model

The product is the **OfficeEditor library suite**. The libraries run in the caller's process and are intended to create, edit, inspect, and render documents supplied by that caller. They validate malformed models and many malformed OpenXML package structures, but they are **not a sandbox for hostile documents**.

Applications that accept content from untrusted users are responsible for their deployment boundary, including:

- request, archive, decompressed-size, page/slide, memory, and CPU limits;
- cancellation and operation deadlines;
- worker-process or container isolation when a native crash or forced timeout must be contained;
- restricting filesystem access and the image/font paths made available to generation;
- authentication, authorization, quotas, HTTPS, logging, and tenant isolation.

TypstBridge is an in-process native Rust backend. Its ABI validates inputs and fences unwind panics, but an in-process native library cannot isolate aborts, out-of-memory conditions, or infinite work. Use a killable worker process when processing hostile content.

OfficeEditor does not execute Office macros or VBA. File paths explicitly supplied to builder or generation APIs may be read by the hosting process under its own operating-system permissions.

## Demo and integration surfaces

`OfficeEditor.Api`, `OfficeEditor.Web.Client`, the LibreOffice comparison service, and the MCP stdio host are local demos and integration conveniences, not hosted multi-tenant products. They do not provide a complete production security boundary. Do not expose them directly to the public internet without adding the controls listed above.

LibreOffice and Poppler are optional external comparison/rendering tools used by the demo and benchmarks; they are not embedded in the OfficeEditor libraries or NuGet packages.

## In scope

Examples of issues worth reporting include:

- unintended file access or path-confinement bypasses;
- command, JSON/YAML, shell, XML, or template injection;
- memory-safety defects in TypstBridge or its managed/native boundary;
- malformed input that corrupts output, leaks resources, or violates a documented failure contract;
- disproportionate denial of service that bypasses documented/configured limits;
- dependency or release-pipeline compromise.

Ordinary resource consumption within caller-selected inputs and limits, missing authentication in the explicitly local demo, and unsupported Office rendering fidelity are not library vulnerabilities by themselves.
