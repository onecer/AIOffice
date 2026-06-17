<#
.SYNOPSIS
    AIOffice installer for Windows (PowerShell).

.DESCRIPTION
    Downloads the matching prebuilt aioffice.exe from a GitHub release, verifies
    its SHA256 against the release SHA256SUMS, installs it, and offers to add the
    install directory to your user PATH.

    Run from an elevated or normal PowerShell prompt:

      irm https://raw.githubusercontent.com/onecer/AIOffice/main/dist/install.ps1 | iex

    or, with a pinned version / custom directory:

      $env:VERSION = "v1.9.0"; $env:AIOFFICE_BIN = "C:\tools\aioffice"
      irm https://raw.githubusercontent.com/onecer/AIOffice/main/dist/install.ps1 | iex

.PARAMETER Version
    Release tag to install (e.g. v1.9.0). Defaults to $env:VERSION, then the
    latest GitHub release, then the hardcoded fallback.

.PARAMETER InstallDir
    Install directory. Defaults to $env:AIOFFICE_BIN, then
    "$env:LOCALAPPDATA\aioffice".
#>
[CmdletBinding()]
param(
    [string]$Version    = $env:VERSION,
    [string]$InstallDir = $env:AIOFFICE_BIN
)

$ErrorActionPreference = "Stop"
Set-StrictMode -Version Latest

$Repo            = "onecer/AIOffice"
$FallbackVersion = "v1.9.0"

function Write-Info($msg) { Write-Host "aioffice-install: $msg" }
function Die($msg) { Write-Error "aioffice-install: $msg"; exit 1 }

# --- detect architecture -> asset name -------------------------------------
function Get-Asset {
    $arch = $env:PROCESSOR_ARCHITECTURE
    if (-not $arch) { $arch = "AMD64" }
    switch ($arch.ToUpper()) {
        "ARM64" { return "aioffice-win-arm64.exe" }
        "AMD64" { return "aioffice-win-x64.exe" }
        "X86"   { Die "32-bit x86 is not supported (need x64 or arm64)" }
        default { Die "unsupported CPU architecture: $arch" }
    }
}

# --- resolve version -------------------------------------------------------
function Resolve-Version {
    if ($Version) { return $Version }
    try {
        $api = "https://api.github.com/repos/$Repo/releases/latest"
        $resp = Invoke-RestMethod -Uri $api -Headers @{ "User-Agent" = "aioffice-install" }
        if ($resp.tag_name) { return $resp.tag_name }
    } catch {
        Write-Info "could not resolve latest release from GitHub API; falling back to $FallbackVersion"
    }
    return $FallbackVersion
}

# --- main ------------------------------------------------------------------
$asset      = Get-Asset
$releaseTag = Resolve-Version

if (-not $InstallDir) { $InstallDir = Join-Path $env:LOCALAPPDATA "aioffice" }
$dest = Join-Path $InstallDir "aioffice.exe"

$base    = "https://github.com/$Repo/releases/download/$releaseTag"
$binUrl  = "$base/$asset"
$sumsUrl = "$base/SHA256SUMS"

Write-Info "$releaseTag $asset -> $dest"

# Force TLS 1.2+ on older PowerShell hosts (Windows PowerShell 5.1 defaults to
# SSL3/TLS1.0). The enum type is [Net.SecurityProtocolType]; add Tls13 in its own
# try so a host that lacks it still gets Tls12.
try { [Net.ServicePointManager]::SecurityProtocol = [Net.ServicePointManager]::SecurityProtocol -bor [Net.SecurityProtocolType]::Tls12 } catch {}
try { [Net.ServicePointManager]::SecurityProtocol = [Net.ServicePointManager]::SecurityProtocol -bor [Net.SecurityProtocolType]::Tls13 } catch {}

$tmp = Join-Path ([System.IO.Path]::GetTempPath()) ("aioffice-install-" + [System.Guid]::NewGuid().ToString("N"))
New-Item -ItemType Directory -Path $tmp -Force | Out-Null
try {
    $tmpBin  = Join-Path $tmp "aioffice.exe"
    $tmpSums = Join-Path $tmp "SHA256SUMS"

    # 1. download binary + checksums
    try {
        Invoke-WebRequest -Uri $binUrl  -OutFile $tmpBin  -UseBasicParsing
        Invoke-WebRequest -Uri $sumsUrl -OutFile $tmpSums -UseBasicParsing
    } catch {
        Die "download failed from $base (check that release $releaseTag exists and has asset $asset): $($_.Exception.Message)"
    }

    # 2. verify SHA256
    $expected = $null
    foreach ($line in (Get-Content $tmpSums)) {
        $parts = $line -split '\s+', 2
        if ($parts.Count -eq 2 -and $parts[1].Trim() -eq $asset) {
            $expected = $parts[0].Trim().ToLower()
            break
        }
    }
    if (-not $expected) { Die "no SHA256 entry for $asset in SHA256SUMS" }

    $actual = (Get-FileHash -Path $tmpBin -Algorithm SHA256).Hash.ToLower()
    if ($expected -ne $actual) {
        Write-Error "aioffice-install: SHA256 mismatch for $asset"
        Write-Error "  expected: $expected"
        Write-Error "  actual:   $actual"
        Die "refusing to install a binary that failed integrity verification"
    }
    Write-Info "sha256 verified ($actual)"

    # 3. install. On Windows a running aioffice.exe is LOCKED against overwrite but
    #    can still be renamed, so move any existing copy aside first, then install the
    #    new one. A stale ".old-*" left behind by a still-running process is cleaned
    #    up on a later upgrade, once the process holding it exits.
    New-Item -ItemType Directory -Path $InstallDir -Force | Out-Null
    if (Test-Path $dest) {
        $old = "$dest.old-$([System.Guid]::NewGuid().ToString('N'))"
        try { Move-Item -Path $dest -Destination $old -Force } catch {}
    }
    Copy-Item -Path $tmpBin -Destination $dest -Force
    Get-ChildItem -Path $InstallDir -Filter "aioffice.exe.old-*" -ErrorAction SilentlyContinue |
        ForEach-Object { try { Remove-Item -Path $_.FullName -Force -ErrorAction Stop } catch {} }
}
finally {
    Remove-Item -Path $tmp -Recurse -Force -ErrorAction SilentlyContinue
}

# 4. PATH hint / add to user PATH
$userPath = [Environment]::GetEnvironmentVariable("Path", "User")
if (-not $userPath) { $userPath = "" }
$onPath = ($userPath -split ';') -contains $InstallDir
if (-not $onPath) {
    Write-Info "NOTE $InstallDir is not on your PATH."
    Write-Info "  Adding it to your user PATH (takes effect in new terminals)."
    $newPath = if ($userPath.TrimEnd(';')) { "$($userPath.TrimEnd(';'));$InstallDir" } else { $InstallDir }
    [Environment]::SetEnvironmentVariable("Path", $newPath, "User")
    # Make it usable in the current session too.
    $env:Path = "$env:Path;$InstallDir"
}

# 5. confirm install
Write-Info "installed."
try {
    & $dest version
} catch {
    Write-Info ("installed to $dest (run 'aioffice version' to confirm).")
}
