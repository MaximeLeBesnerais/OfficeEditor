#!/usr/bin/env bash
set -euo pipefail

fail() {
  printf 'pack-runtime-assets.sh: %s\n' "$*" >&2
  exit 1
}

SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
BRIDGE_ROOT="$(cd "$SCRIPT_DIR/.." && pwd)"
NATIVE_DIR="$BRIDGE_ROOT/native"
RID="${1:-linux-x64}"

case "$RID" in
  linux-x64)
    LIB_NAME="libtypst_bridge.so"
    ARTIFACT="$NATIVE_DIR/target/x86_64-unknown-linux-gnu/release/$LIB_NAME"
    ;;
  linux-arm64)
    LIB_NAME="libtypst_bridge.so"
    ARTIFACT="$NATIVE_DIR/target/aarch64-unknown-linux-gnu/release/$LIB_NAME"
    ;;
  osx-x64|osx-arm64)
    fail "$RID runtime packing is planned but not wired yet. Expected library name: libtypst_bridge.dylib."
    ;;
  win-x64|win-arm64)
    fail "$RID runtime packing should be performed on Windows after build-native.ps1 produces typst_bridge.dll."
    ;;
  *)
    fail "unsupported RID '$RID'. Supported by this script: linux-x64, linux-arm64. Planned elsewhere: osx-x64, osx-arm64, win-x64, win-arm64."
    ;;
esac

[[ -f "$NATIVE_DIR/Cargo.toml" ]] || fail "native crate not found at $NATIVE_DIR/Cargo.toml. Create TypstBridge/native before packing runtime assets."
[[ -f "$ARTIFACT" ]] || fail "expected native artifact was not found: $ARTIFACT. Run build-native.sh $RID first."

mkdir -p "$BRIDGE_ROOT/runtimes/$RID/native"
cp "$ARTIFACT" "$BRIDGE_ROOT/runtimes/$RID/native/$LIB_NAME"
printf 'Copied %s to %s\n' "$LIB_NAME" "$BRIDGE_ROOT/runtimes/$RID/native/$LIB_NAME"
