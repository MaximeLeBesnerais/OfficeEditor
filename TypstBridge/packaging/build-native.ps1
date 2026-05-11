param(
    [string]$Rid = "win-x64"
)

$ErrorActionPreference = "Stop"

function Fail([string]$Message) {
    Write-Error "build-native.ps1: $Message"
    exit 1
}

$ScriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$BridgeRoot = Resolve-Path (Join-Path $ScriptDir "..")
$NativeDir = Join-Path $BridgeRoot "native"

if (-not (Get-Command cargo -ErrorAction SilentlyContinue)) {
    Fail "cargo was not found. Install Rust from https://rustup.rs/ and retry."
}

switch ($Rid) {
    "win-x64" {
        $LibName = "typst_bridge.dll"
        $RustTarget = "x86_64-pc-windows-msvc"
        $Artifact = Join-Path $NativeDir "target\$RustTarget\release\$LibName"
    }
    "win-arm64" {
        $LibName = "typst_bridge.dll"
        $RustTarget = "aarch64-pc-windows-msvc"
        $Artifact = Join-Path $NativeDir "target\$RustTarget\release\$LibName"
    }
    "linux-x64" { Fail "linux-x64 packaging is handled by build-native.sh on Linux." }
    "linux-arm64" { Fail "linux-arm64 packaging is handled by build-native.sh on Linux." }
    "osx-x64" { Fail "osx-x64 packaging is planned but not wired yet. Build on macOS and copy libtypst_bridge.dylib to runtimes/osx-x64/native." }
    "osx-arm64" { Fail "osx-arm64 packaging is planned but not wired yet. Build on macOS and copy libtypst_bridge.dylib to runtimes/osx-arm64/native." }
    default {
        Fail "unsupported RID '$Rid'. Supported by this script: win-x64, win-arm64. Planned elsewhere: linux-x64, linux-arm64, osx-x64, osx-arm64."
    }
}

$CargoToml = Join-Path $NativeDir "Cargo.toml"
if (-not (Test-Path $CargoToml)) {
    Fail "native crate not found at $CargoToml. Create TypstBridge/native before running this script."
}

$RuntimeDir = Join-Path $BridgeRoot "runtimes\$Rid\native"
New-Item -ItemType Directory -Force -Path $RuntimeDir | Out-Null

Write-Host "Building TypstBridge native crate for $Rid ($RustTarget)..."
cargo build --release --target $RustTarget --manifest-path $CargoToml
if ($LASTEXITCODE -ne 0) {
    Fail "cargo build failed with exit code $LASTEXITCODE."
}

if (-not (Test-Path $Artifact)) {
    Fail "expected build artifact was not found: $Artifact"
}

$Destination = Join-Path $RuntimeDir $LibName
Copy-Item -Force $Artifact $Destination
Write-Host "Copied $LibName to $Destination"
