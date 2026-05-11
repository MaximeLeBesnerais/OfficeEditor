#!/usr/bin/env bash
set -euo pipefail

fail() {
  printf 'build-native.sh: %s\n' "$*" >&2
  exit 1
}

SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
BRIDGE_ROOT="$(cd "$SCRIPT_DIR/.." && pwd)"
NATIVE_DIR="$BRIDGE_ROOT/native"
RID="${1:-linux-x64}"

command -v cargo >/dev/null 2>&1 || fail "cargo was not found. Install Rust from https://rustup.rs/ and retry."

case "$RID" in
  linux-x64)
    LIB_NAME="libtypst_bridge.so"
    RUST_TARGET="x86_64-unknown-linux-gnu"
    ARTIFACT="$NATIVE_DIR/target/$RUST_TARGET/release/$LIB_NAME"
    ;;
  linux-arm64)
    LIB_NAME="libtypst_bridge.so"
    RUST_TARGET="aarch64-unknown-linux-gnu"
    ARTIFACT="$NATIVE_DIR/target/$RUST_TARGET/release/$LIB_NAME"
    ;;
  osx-x64|osx-arm64)
    fail "$RID packaging is planned but not wired yet. Build on macOS and copy libtypst_bridge.dylib to runtimes/$RID/native."
    ;;
  win-x64|win-arm64)
    fail "$RID packaging is handled by build-native.ps1 on Windows."
    ;;
  *)
    fail "unsupported RID '$RID'. Supported by this script: linux-x64, linux-arm64. Planned elsewhere: osx-x64, osx-arm64, win-x64, win-arm64."
    ;;
esac

[[ -f "$NATIVE_DIR/Cargo.toml" ]] || fail "native crate not found at $NATIVE_DIR/Cargo.toml. Create TypstBridge/native before running this script."

mkdir -p "$BRIDGE_ROOT/runtimes/$RID/native"

EXISTING_RUSTFLAGS="${RUSTFLAGS:-}"
if [[ -n "$EXISTING_RUSTFLAGS" ]]; then
  export RUSTFLAGS="$EXISTING_RUSTFLAGS -C link-arg=-Wl,-z,noexecstack"
else
  export RUSTFLAGS="-C link-arg=-Wl,-z,noexecstack"
fi

printf 'Building TypstBridge native crate for %s (%s)...\n' "$RID" "$RUST_TARGET"
cargo build --release --target "$RUST_TARGET" --manifest-path "$NATIVE_DIR/Cargo.toml"

[[ -f "$ARTIFACT" ]] || fail "expected build artifact was not found: $ARTIFACT"

cp "$ARTIFACT" "$BRIDGE_ROOT/runtimes/$RID/native/$LIB_NAME"
printf 'Copied %s to %s\n' "$LIB_NAME" "$BRIDGE_ROOT/runtimes/$RID/native/$LIB_NAME"
