# Phase 5 — Orchestration (worktree-per-agent)

Master spec: `plan.md` (repo root). This file is the dispatch/merge machinery.

## Mechanism

- Each workstream P1–P10 has a prompt file here (`P<N>.md`) + the shared contract (`SHARED.md`).
- Dispatch prompt is one line: "Read `agent-tasks/phase5/SHARED.md`, then `agent-tasks/phase5/P<N>.md`, in your worktree. Apply with no discussion."
- Each agent gets a git worktree branched from current `main` at dispatch time.
- Cross-file needs flow through `integration/*.patch` in the agent's worktree; the orchestrator applies them between waves.

## Dependency graph & waves

| Wave | Workstreams | Gate to enter |
|---|---|---|
| 1 | P1, P8 | none (P8's dep F9/brand-extractor is already merged) |
| 2 | P2 | P1 merged |
| 3 | P3 | P2 merged |
| 4 | P4 ∥ P5 | P2 merged |
| 5 | P7 | P4+P5 merged |
| 6 | P6 | P2 merged; P4/P5 available for fixtures |
| 7 | P9, P10 | P1–P5 merged (P9); P6 merged (P10) |

Merge order (plan.md §6): **P1 → P2 → P3 → (P4 ∥ P5) → P7 → P6 → (P8 ∥ P9 ∥ P10).**

## Worktree naming

- Branch: `feat/phase5-p<N>-<slug>` (e.g. `feat/phase5-p1-model`).
- Path: `$TMPDIR/opencode/phase5/wt-p<N>`.

## Wave merge gates

Each merge requires, in the main checkout:
1. `dotnet test` green.
2. Smoke decks green (examples/REF: REMOVED, pres-pro).
3. No files outside the workstream's ownership (except orchestrator-applied integration patches).

## Phase-level definition of done

plan.md §7, verbatim. Do not re-negotiate.
