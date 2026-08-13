<#
.SYNOPSIS
    Sync the vendored OfficeCLI engine from the local upstream checkout.

.DESCRIPTION
    Copies src/officecli (minus bin, obj and sync-exclude.txt entries), skills/ and
    schemas/help/ from the upstream OfficeCLI repository into this project, keeping
    the vendored files byte-identical to upstream (fidelity rule: only DELETE is
    allowed — see VENDOR.md).

    After the copy it prints `git diff --stat` as the change report.

.PARAMETER UpstreamPath
    Path to the upstream OfficeCLI checkout (defaults to .\..\OfficeCLI).

.EXAMPLE
    .\sync-from-upstream.ps1
    .\sync-from-upstream.ps1 -UpstreamPath D:\src\OfficeCLI
#>
param(
    [string]$UpstreamPath = (Join-Path (Split-Path $PSScriptRoot -Parent) 'OfficeCLI')
)

$ErrorActionPreference = 'Stop'

function Test-Dir([string]$Path, [string]$What) {
    if (-not (Test-Path $Path -PathType Container)) {
        throw "Upstream $What not found at: $Path"
    }
}

$upstream = (Resolve-Path $UpstreamPath).Path
$root = $PSScriptRoot

Test-Dir (Join-Path $upstream 'src\officecli') 'src\officecli'
Test-Dir (Join-Path $upstream 'skills') 'skills'
Test-Dir (Join-Path $upstream 'schemas\help') 'schemas\help'

# --- exclude list from sync-exclude.txt --------------------------------
$excludes = Get-Content (Join-Path $root 'sync-exclude.txt') |
    Where-Object { $_ -and -not $_.TrimStart().StartsWith('#') }

function Match-Exclude([string]$RelPath) {
    foreach ($pattern in $excludes) {
        $p = $pattern.Trim().Replace('\', '/').TrimEnd('/')
        $r = $RelPath.Replace('\', '/')
        if ($p -like '*') {
            # glob relative to upstream root
            if ($r -like $p) { return $true }
        }
        if ($r -eq $p) { return $true }
        if ($r.StartsWith($p + '/')) { return $true }
    }
    return $false
}

# --- copy src\officecli ------------------------------------------------
Write-Host "==> Syncing src\officecli from $upstream"
$src = Join-Path $upstream 'src\officecli'
$dst = Join-Path $root 'src\officecli'
$copied = 0
Get-ChildItem $src -Recurse -File | ForEach-Object {
    $rel = $_.FullName.Substring($src.Length + 1)
    if ($rel -like 'bin\*' -or $rel -like 'obj\*') { return }
    if (Match-Exclude ('src\officecli\' + $rel)) {
        Write-Verbose "excluded: $rel"
        return
    }
    $target = Join-Path $dst $rel
    New-Item -ItemType Directory -Force -Path (Split-Path $target -Parent) | Out-Null
    Copy-Item $_.FullName $target -Force
    $copied++
}
Write-Host "  $copied files copied"

# --- copy skills and schemas/help ---------------------------------------
foreach ($sub in @('skills', 'schemas')) {
    Write-Host "==> Syncing $sub"
    robocopy (Join-Path $upstream $sub) (Join-Path $root $sub) /E /NFL /NDL /NJH /NJS | Out-Null
    if ($LASTEXITCODE -ge 8) { throw "robocopy failed for $sub (exit $LASTEXITCODE)" }
}

# --- report -------------------------------------------------------------
Write-Host ""
Write-Host "==> git diff --stat (vendored changes vs last sync):"
git -C $root diff --stat -- src skills schemas
Write-Host ""
Write-Host "Next: update version/commit in VENDOR.md, build, run OfficeTool.Tests,"
Write-Host "      and re-check --output-schema-crc parity with the upstream CLI."
