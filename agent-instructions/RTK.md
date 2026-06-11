# RTK - Shell Token Filter

**Usage**: Token-optimized CLI proxy for shell commands. Despite the name ("Rust Token Killer"), it is a generic shell-output filter and works with any toolchain — this project is .NET 9 / C#.

## Rule

Always prefix shell commands with `rtk`.

Examples:

```bash
rtk git status
rtk dotnet test
rtk dotnet build
```

## Setup Verification (mandatory first step)

```bash
rtk --version
```

If `rtk` is **not installed**:

- Fall back to running the plain command: `git status`, `dotnet test`, etc.
- Or use the explicit proxy form: `rtk proxy <cmd>` (bypasses filtering for that command only).
- Do **not** retry the filtered form repeatedly — it will keep failing and waste time.

## Meta Commands

```bash
rtk gain            # Token savings analytics
rtk gain --history  # Recent command savings history
rtk proxy <cmd>     # Run raw command without filtering
```

## Project-Local Config

Custom filters live in `.rtk/filters.toml` (committed to the repo) and override user-global / built-in filters. See that file for the schema and examples.

## Verification

```bash
rtk --version       # confirm install
rtk gain            # see accumulated savings
which rtk           # confirm location
```
