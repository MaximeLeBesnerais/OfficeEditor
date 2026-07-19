# Phase 5 — Shared Behavior Contract & Conventions

> MANDATORY. Read fully before your workstream file. Apply with no discussion.

## Behavior contract (acute thinking)

Structure ALL your work exactly like this:

1. **What is asked** — restate your task in ≤3 lines.
2. **How to do it** — the shortest viable path. No alternatives essay.
3. **Do it** — execute decisively. No hedging, no "wait, what if…". You decide, you execute.

Then report:

1. **Is it done?** — verified with evidence (build/test/diff output).
2. **Deviations?** — list them, or "none".
3. **Report** — what changed (files), why, how verified.

Efficiency rule: **What, why, how — report.** You are purpose-oriented; you complete the task. Self-hesitation is failure.

## Hard rules (from AGENTS.md / plan.md)

- Run `dotnet test` after every change. All tests must pass inside your worktree.
- Never commit on your worktree branch unless the orchestrator explicitly says so.
- `agent-instructions/AGENTS.pptx.md` rules apply to ALL OOXML emission (regex attribute reads, spcPct units, no global autofit, per-shape normAutofit only).
- Load-bearing rules from plan.md §2 apply to everyone:
  1. Layout is resolved exactly once, in C#. Typst is a dumb renderer (`#place` only).
  2. Every primitive ships with both emitters + one parity fixture in the same PR.
  3. Coordinates are the escape hatch, not the API.
  4. Preview tricks never enter the PPTX.
- v1 non-goals (plan.md §1): no percentages (pt only), no wrap, no z-index, no CSS input, no Tier-3 rendering.
- Conventions: nullable reference types on, PascalCase public APIs, `_camelCase` private fields, validate all inputs, never swallow exceptions.

## Worktree conventions

- You work inside a git worktree at the path given in your dispatch prompt. The repo there IS the real repo (shared `.git`).
- All paths in your P-file are relative to your worktree root.
- **Strict file ownership**: touch ONLY files under "Owns" in your P-file.
- Need a change outside your ownership? Do NOT edit it. Produce `integration/<name>.patch` (a `git diff` patch against your worktree HEAD) at your worktree root and flag it in your report. The orchestrator applies it between waves.
- Build/test from your worktree root: `dotnet build DocxEditor.sln` / `dotnet test`. Your `bin/obj` are isolated — parallel worktrees do not conflict.
- If you need `rtk`, read `agent-instructions/RTK.md` first.
