<#
.SYNOPSIS
    Semi-automatic update of OfficeCliEngine (+ the OfficeTool agent surface)
    from the upstream STABLE release.

.DESCRIPTION
    Runs the full update pipeline against the latest (or pinned) OfficeCLI release:

      1. sync-from-upstream.ps1  — download the release Source code (zip), vendor the
                                   engine byte-identical, update VENDOR.md;
      2. Gap analysis            — CLI commands vs OfficeTool methods (every documented
                                   command must exist as a method), view modes, and
                                   embedded-resource parity between the upstream csproj
                                   and ours (a new embedded file missing here breaks the
                                   engine at runtime);
      3. Build engine            — surfaces engine API / package drift;
      4. Build AIOrchestrator    — surfaces OfficeTool breakage against the new engine;
      5. Run OfficeTool.Tests    — the deterministic harness (docx/xlsx/pptx);
      6. Optional: pack          — verify the NuGet package builds.

    The report is written to sync-gap-report.md at the repo root (gitignored) and
    printed. Nothing is committed or pushed: review, fix the reported gaps, then
    commit (the CI workflow publishes NuGet on push to master).

.PARAMETER Tag
    Release tag to sync (default: latest). Passed through to sync-from-upstream.ps1.

.PARAMETER UpstreamPath
    Offline sync from a local source tree. Passed through to sync-from-upstream.ps1.

.PARAMETER SkipSync / SkipBuild / SkipTests / SkipPack
    Skip individual pipeline stages (e.g. -SkipSync to re-analyze an already-synced tree).

.EXAMPLE
    .\update-officecli.ps1                 # full update to the latest release
    .\update-officecli.ps1 -Tag v1.0.144   # update to a pinned release
    .\update-officecli.ps1 -SkipSync       # gap analysis + build + tests on current tree
#>
param(
    [string]$Tag,
    [string]$UpstreamPath,
    [switch]$SkipSync,
    [switch]$SkipBuild,
    [switch]$SkipTests,
    [switch]$SkipPack
)

$ErrorActionPreference = 'Stop'
$root = $PSScriptRoot
$reportPath = Join-Path $root 'sync-gap-report.md'
$report = New-Object System.Text.StringBuilder

function Write-Report([string]$Line) {
    Write-Host $Line
    [void]$report.AppendLine($Line)
}
function Read-Utf8([string]$Path) { return [IO.File]::ReadAllText($Path, [Text.Encoding]::UTF8) }

# --- prereqs -----------------------------------------------------------------
foreach ($tool in @('git', 'dotnet')) {
    if (-not (Get-Command $tool -ErrorAction SilentlyContinue)) { throw "Required tool '$tool' not found on PATH." }
}
if (-not $UpstreamPath -and -not $SkipSync -and -not (Get-Command gh -ErrorAction SilentlyContinue)) {
    throw "gh CLI not found on PATH (needed to resolve the release). Pass -UpstreamPath for offline sync."
}

Write-Report "=== OfficeCliEngine update pipeline ==="
Write-Report ("Start: " + (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'))

# --- 1. sync -----------------------------------------------------------------
if (-not $SkipSync) {
    $syncArgs = @{}
    if ($Tag) { $syncArgs['Tag'] = $Tag }
    if ($UpstreamPath) { $syncArgs['UpstreamPath'] = $UpstreamPath }
    & (Join-Path $root 'sync-from-upstream.ps1') @syncArgs
    if ($LASTEXITCODE -ne 0) { throw "sync-from-upstream.ps1 failed (exit $LASTEXITCODE)." }
}
else {
    Write-Report "(sync skipped)"
}

# --- 2. gap analysis ---------------------------------------------------------
Write-Report ""
Write-Report "=== Gap analysis ==="

$toolPath = Join-Path (Split-Path $root -Parent) 'AIOrchestrator\API\OfficeTool.cs'
if (-not (Test-Path $toolPath)) { throw "OfficeTool.cs not found at: $toolPath" }

function Get-CliCommands([string]$EngineRoot) {
    $cmds = @{}
    # CommandBuilder.Plugins.cs is entirely subcommands of the excluded `plugins`
    # command (environment ops) — skip the file so list/info/lint never flag.
    Get-ChildItem (Join-Path $EngineRoot 'src\officecli\CommandBuilder*.cs') -File |
        Where-Object { $_.Name -ne 'CommandBuilder.Plugins.cs' } | ForEach-Object {
        $text = Read-Utf8 $_.FullName
        foreach ($m in [regex]::Matches($text, 'new Command\("([A-Za-z0-9_\-]+)"')) {
            $cmds[$m.Groups[1].Value] = $true
        }
    }
    $viewFile = Join-Path $EngineRoot 'src\officecli\CommandBuilder.View.cs'
    if (Test-Path $viewFile) {
        # Only the comma-separated mode list right after "View mode:" — the rest
        # of the description explains text mode and must not be parsed.
        $m = [regex]::Match((Read-Utf8 $viewFile), 'View mode:\s*([a-zA-Z]+(?:,\s*[a-zA-Z]+)*)')
        if ($m.Success) {
            $m.Groups[1].Value -split ',' | ForEach-Object {
                $v = $_.Trim()
                if ($v) { $cmds["view $v"] = $true }
            }
        }
    }
    return $cmds
}

function Get-OfficeToolMethods([string]$Path) {
    $text = Read-Utf8 $Path
    $methods = @{}
    foreach ($m in [regex]::Matches($text, 'public\s+(?!class\s)(?:static\s+)?(?:override\s+)?[\w<>,\[\]\.\? ]+\s+(\w+)\s*\(')) {
        $methods[$m.Groups[1].Value] = $true
    }
    return $methods
}

# CLI command -> OfficeTool method. EXCLUDED = intentionally not ported (see OfficePorting.md §3).
$map = @{
    'open'              = 'EXCLUDED' # resident infra — instance state replaces it (Open()/Create())
    'close'             = 'EXCLUDED' # resident infra — Save()/Dispose() replace it
    '__resident-serve__' = 'EXCLUDED' # internal CLI plumbing
    'watch'             = 'Watch'
    'unwatch'           = 'Unwatch'
    'mark'              = 'Mark'
    'unmark'            = 'Unmark'
    'get-marks'         = 'GetMarks'
    'goto'              = 'Goto'
    'get'               = 'Get'
    'get selected'      = 'GetSelected'
    'query'             = 'Query'
    'set'               = 'Set'
    'add'               = 'Add'
    'remove'            = 'Remove'
    'move'              = 'Move'
    'swap'              = 'Swap'
    'refresh'           = 'EXCLUDED' # Word TOC/page refresh — backend candidate, not yet ported
    'raw'               = 'Raw'
    'raw-set'           = 'RawSet'
    'add-part'          = 'AddPart'
    'validate'          = 'Validate'
    'save'              = 'Save'
    'batch'             = 'Batch'
    'dump'              = 'Dump'
    'import'            = 'Import'
    'create'            = 'Create'
    'merge'             = 'Merge'
    'plugins'           = 'EXCLUDED' # environment op (installer/plugins), not agent-facing
    'install'           = 'EXCLUDED' # environment op
    'config'            = 'EXCLUDED' # environment op
    'mcp'               = 'EXCLUDED' # environment op
    'help'              = 'Help'
    'load_skill'        = 'LoadSkill'
    'view'              = 'EXCLUDED' # container command — its modes are mapped individually below
    'view text'         = 'ViewText'
    'view annotated'    = 'ViewAnnotated'
    'view outline'      = 'ViewOutline'
    'view stats'        = 'ViewStats'
    'view issues'       = 'ViewIssues'
    'view html'         = 'ViewHtml'
    'view svg'          = 'ViewSvg'
    'view screenshot'   = 'ViewScreenshot'
    'view pdf'          = 'EXCLUDED' # exporter plugin (installed plugin + subprocess), not agent-facing
    'view forms'        = 'ViewForms'
}

$cmds = Get-CliCommands $root
$methods = Get-OfficeToolMethods $toolPath
$newCmds = @(); $missingMethods = @(); $excluded = @()

foreach ($cmd in $cmds.Keys) {
    if (-not $map.ContainsKey($cmd)) { $newCmds += $cmd; continue }
    $target = $map[$cmd]
    if ($target -eq 'EXCLUDED') { $excluded += $cmd; continue }
    if (-not $methods.ContainsKey($target)) { $missingMethods += "$cmd -> $target" }
}

$orphans = @($methods.Keys | Where-Object { -not ($map.Values -contains $_) })

Write-Report ("CLI commands found:    " + $cmds.Count)
Write-Report ("Intentional exclusions: " + ($excluded -join ', '))
Write-Report ("OfficeTool methods:    " + $methods.Count)

if ($newCmds.Count -eq 0) { Write-Report "OK — no new CLI commands." }
else {
    Write-Report ("WARNING — NEW CLI command(s) without a mapping: " + ($newCmds -join ', '))
    Write-Report "         Add an OfficeTool method (or document an intentional exclusion in the script)."
}
if ($missingMethods.Count -eq 0) { Write-Report "OK — every mapped command has its method." }
else {
    Write-Report ("WARNING — mapped command(s) whose method is missing: " + ($missingMethods -join ', '))
}
if ($orphans.Count -eq 0) { Write-Report "OK — no orphan methods." }
else { Write-Report ("Info — methods without a CLI counterpart (backup/restore, skill files, props): " + ($orphans -join ', ')) }

# --- embedded-resource parity (upstream csproj vs ours) ----------------------
Write-Report ""
Write-Report "=== Embedded-resource parity ==="

function Resolve-Glob([string]$Base, [string]$Pattern) {
    $norm = $Pattern.Replace('/', '\')
    if ($norm -notmatch '[\*\?]') {
        $p = [IO.Path]::GetFullPath((Join-Path $Base $norm))
        return @(if (Test-Path $p) { $p })
    }
    $result = @()
    if ($norm -match '\*\*') {
        $idx = $norm.IndexOf('\**')
        $prefix = [IO.Path]::GetFullPath((Join-Path $Base $norm.Substring(0, $idx)))
        $rest = $norm.Substring($idx + 3).TrimStart('\')
        if (Test-Path $prefix) {
            if ($rest) { Get-ChildItem $prefix -Recurse -File | Where-Object { $_.Name -like $rest } | ForEach-Object { $result += $_.FullName } }
            else { Get-ChildItem $prefix -Recurse -File | ForEach-Object { $result += $_.FullName } }
        }
    }
    else {
        $sep = $norm.LastIndexOf('\')
        $dirPart = if ($sep -ge 0) { $norm.Substring(0, $sep) } else { '' }
        $filePart = if ($sep -ge 0) { $norm.Substring($sep + 1) } else { $norm }
        $dir = if ($dirPart) { [IO.Path]::GetFullPath((Join-Path $Base $dirPart)) } else { [IO.Path]::GetFullPath($Base) }
        if (Test-Path $dir) { Get-ChildItem $dir -File | Where-Object { $_.Name -like $filePart } | ForEach-Object { $result += $_.FullName } }
    }
    return $result
}

function Get-EmbeddedFiles([string]$CsprojPath, [string]$ProjectDir) {
    $text = Read-Utf8 $CsprojPath
    $files = New-Object 'System.Collections.Generic.HashSet[string]' ([StringComparer]::OrdinalIgnoreCase)
    foreach ($m in [regex]::Matches($text, '<EmbeddedResource\s+Include="([^"]+)"')) {
        Resolve-Glob $ProjectDir $m.Groups[1].Value | ForEach-Object { [void]$files.Add($_) }
    }
    foreach ($m in [regex]::Matches($text, '<EmbeddedResource\s+Remove="([^"]+)"')) {
        Resolve-Glob $ProjectDir $m.Groups[1].Value | ForEach-Object { [void]$files.Remove($_) }
    }
    return $files
}

$upstreamCsproj = Join-Path $root 'src\officecli\officecli.csproj'
$ourCsproj = Join-Path $root 'OfficeCliEngine.csproj'
if (Test-Path $upstreamCsproj) {
    $upFiles = Get-EmbeddedFiles $upstreamCsproj (Join-Path $root 'src\officecli')
    $ourFiles = Get-EmbeddedFiles $ourCsproj $root
    $missing = @($upFiles | Where-Object { -not $ourFiles.Contains($_) })
    if ($missing.Count -eq 0) {
        Write-Report ("OK — " + $upFiles.Count + " embedded resources match the upstream csproj.")
    }
    else {
        Write-Report ("WARNING — embedded resource(s) upstream embeds but this csproj does NOT: ")
        $missing | ForEach-Object { Write-Report ("           " + $_) }
        Write-Report "         Add them to OfficeCliEngine.csproj or the engine breaks at runtime."
    }
}
else {
    Write-Report "(vendored officecli.csproj not present — resource parity check skipped)"
}

# --- write report ------------------------------------------------------------
[IO.File]::WriteAllText($reportPath, $report.ToString(), (New-Object System.Text.UTF8Encoding($false)))
Write-Report ""
Write-Report ("Report written to: $reportPath")

# --- 3. build engine ---------------------------------------------------------
if (-not $SkipBuild) {
    Write-Report ""
    Write-Report "=== Build OfficeCliEngine (Release) ==="
    dotnet build (Join-Path $root 'OfficeCliEngine.csproj') -c Release -v m -nologo
    if ($LASTEXITCODE -ne 0) { throw "Engine build failed (exit $LASTEXITCODE). Fix engine API/resource drift first." }
    Write-Report "Engine build OK."

    $aiOrch = Join-Path (Split-Path $root -Parent) 'AIOrchestrator\AIOrchestrator.csproj'
    if (Test-Path $aiOrch) {
        Write-Report ""
        Write-Report "=== Build AIOrchestrator (surfaces OfficeTool breakage) ==="
        dotnet build $aiOrch -c Debug -v m -nologo
        if ($LASTEXITCODE -ne 0) { throw "AIOrchestrator build failed (exit $LASTEXITCODE). OfficeTool.cs needs updating against the new engine." }
        Write-Report "AIOrchestrator build OK."
    }
    else { Write-Report "(AIOrchestrator not found next to this project — skipped)" }
}
else { Write-Report "(build skipped)" }

# --- 4. harness --------------------------------------------------------------
if (-not $SkipTests) {
    $harness = Join-Path (Split-Path $root -Parent) 'AIOrchestrator\OfficeTool.Tests\OfficeTool.Tests.csproj'
    if (Test-Path $harness) {
        Write-Report ""
        Write-Report "=== OfficeTool.Tests (deterministic harness) ==="
        dotnet run --project $harness -c Debug
        if ($LASTEXITCODE -ne 0) { throw "OfficeTool.Tests failed (exit $LASTEXITCODE)." }
        Write-Report "Harness OK."
    }
    else { Write-Report "(harness not found — OfficeTool.Tests skipped)" }
}
else { Write-Report "(tests skipped)" }

# --- 5. pack ----------------------------------------------------------------
if (-not $SkipPack) {
    Write-Report ""
    Write-Report "=== Pack (verifies the NuGet pipeline, no push) ==="
    dotnet pack (Join-Path $root 'OfficeCliEngine.csproj') -c Release -o (Join-Path $root 'out') -p:SkipNuGetPush=true -v m -nologo
    if ($LASTEXITCODE -ne 0) { throw "Pack failed (exit $LASTEXITCODE)." }
    Write-Report "Pack OK."
}
else { Write-Report "(pack skipped)" }

# --- final checklist ---------------------------------------------------------
Write-Report ""
Write-Report "=== Next steps (manual) ==="
Write-Report "  1. Review sync-gap-report.md and the git diff (sync-from-upstream.ps1 already verified byte-identical parity)."
Write-Report "  2. If the gap report lists new commands: add OfficeTool methods (+ harness tests)."
Write-Report "  3. Update NOTICE.md (version) and OfficePorting.md status, if versions changed."
Write-Report "  4. Commit and push to Graphene-Lab/OfficeCliEngine — the CI workflow publishes the NuGet package."
Write-Report "     (AIOrchestrator will pick it up via PackageReference 1.* where the project is absent.)"
Write-Report "Done: " + (Get-Date -Format 'yyyy-MM-dd HH:mm:ss')
