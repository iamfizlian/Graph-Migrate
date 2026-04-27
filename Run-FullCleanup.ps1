<#
.SYNOPSIS
    End-to-end post-import cleanup for one or many mailboxes.

.DESCRIPTION
    After a fresh Graph-based PST import every mailbox in mapping.csv ends
    up in roughly the same wrong shape:

        1. Mail is buried under an extra "Imported PST/Inbox/...",
           "Imported PST/Sent Items/...", etc. wrapper.
        2. The "Sent Items" you actually want is sometimes living as
           a child of the imported wrapper instead of the top-level
           well-known Sent Items folder.
        3. Every imported message has the MAPI MSGFLAG_UNSENT bit set,
           so OWA badges them all as "Draft" even though they aren't.

    This wrapper drives the three existing fix scripts in the order they
    must run, in a single command, for any subset of mailboxes:

        Step 1 - _flatten_imported.py    (Graph)
                 Merge "Imported PST/<Folder>" into <Folder> at the
                 mailbox root.  Preserves entire subtrees.

        Step 2 - _consolidate_sent.py    (Graph)
                 Re-home stragglers from "Imported PST/Sent Items"
                 (and any other Sent variant) into the well-known
                 Sent Items folder.

        Step 3 - Fix-DraftsViaOutlook.ps1   (Outlook desktop / MAPI HTTP)
                 Clear PR_MESSAGE_FLAGS bit 0x08 on every imported
                 message so OWA stops showing the "Draft" badge.

    Each step is invoked with the same mailbox selection, so the run
    is consistent end to end.  Steps 1+2 are Graph-based and run in
    minutes; Step 3 is Outlook COM and runs at ~25 items/sec, so a
    full 13-mailbox cleanup typically finishes in 1-2 hours of wall
    time, most of which is Step 3.

.PARAMETER Mailbox
    UPN of a target mailbox.  Repeatable.  Mutually exclusive with
    -MappingFile.

        .\Run-FullCleanup.ps1 -Mailbox debbiep@jteatono365.onmicrosoft.com `
                              -Config .\config.toml

.PARAMETER MappingFile
    Path to mapping.csv.  Every row's TargetMailbox is processed,
    in CSV order.

        .\Run-FullCleanup.ps1 -MappingFile .\mapping.csv -Config .\config.toml

.PARAMETER ExcludeMailbox
    UPN of a mailbox to exclude from the run.  Repeatable.  Useful
    for resuming after a partial failure or running only the mailboxes
    that haven't been cleaned up yet, without editing mapping.csv:

        # Allyson is already done -- run the other 12.
        .\Run-FullCleanup.ps1 `
            -MappingFile .\mapping.csv `
            -Config .\config.toml `
            -ExcludeMailbox allysonp@jteatono365.onmicrosoft.com

    Comparison is case-insensitive on full UPN.  Excluding a mailbox
    that isn't in the selection is a no-op (logged but not fatal).

.PARAMETER Config
    Path to config.toml (consumed by the Python steps).

.PARAMETER DryRun
    Forward --dry-run / -DryRun to all three sub-scripts.  Walks and
    reports without writing.

.PARAMETER SkipFlatten
    Skip Step 1 (_flatten_imported.py).  Use this if you've already
    flattened the mailbox(es) on a previous run -- the script is
    idempotent so re-running is safe, but skipping saves a few minutes.

.PARAMETER SkipConsolidate
    Skip Step 2 (_consolidate_sent.py).

.PARAMETER SkipDrafts
    Skip Step 3 (Fix-DraftsViaOutlook.ps1).  Useful if you only want
    the folder reshape and plan to run the draft fix later.

.PARAMETER ContinueOnError
    If a step exits non-zero, log the failure and proceed to the next
    step instead of aborting.  Off by default: a flatten failure that
    leaves mail in the wrong place will produce a confusing draft fix
    pass, so we'd rather stop and let you investigate.

.PARAMETER LogDir
    Directory for per-run transcripts.  Defaults to .\logs\cleanup .
    A timestamped subfolder is created per invocation.

.PARAMETER PythonExe
    Path to the Python interpreter that has the project's deps
    installed.  Defaults to .\.venv\Scripts\python.exe (the venv we
    create in the README's "Install" steps).

.EXAMPLE
    # Verify on one mailbox -- dry run, no writes anywhere.
    .\Run-FullCleanup.ps1 `
        -Mailbox debbiep@jteatono365.onmicrosoft.com `
        -Config  .\config.toml `
        -DryRun

.EXAMPLE
    # Real run for one mailbox.
    .\Run-FullCleanup.ps1 `
        -Mailbox debbiep@jteatono365.onmicrosoft.com `
        -Config  .\config.toml

.EXAMPLE
    # Fan out to every mailbox in mapping.csv.  This is the bulk path.
    .\Run-FullCleanup.ps1 -MappingFile .\mapping.csv -Config .\config.toml

.EXAMPLE
    # Allyson is already fully cleaned up -- cherry-pick the remaining
    # mailboxes by passing -Mailbox repeatedly.
    .\Run-FullCleanup.ps1 -Config .\config.toml `
        -Mailbox debbiep@jteatono365.onmicrosoft.com `
        -Mailbox emmak@jteatono365.onmicrosoft.com `
        -Mailbox isabellel@jteatono365.onmicrosoft.com

.EXAMPLE
    # Skip the Graph reshape on a mailbox you've already flattened by
    # hand and only run the draft-bit fix.
    .\Run-FullCleanup.ps1 `
        -Mailbox allysonp@jteatono365.onmicrosoft.com `
        -Config .\config.toml `
        -SkipFlatten -SkipConsolidate

.PREREQUISITES
    1. .\.venv populated (see README "Install" -> "Windows -- native").
    2. config.toml with valid Entra app credentials (Graph + Mail.ReadWrite).
    3. Outlook desktop installed, signed in to the admin profile, with
       FullAccess granted on every target mailbox using AutoMapping=$false.
       See Fix-DraftsViaOutlook.ps1 .PREREQUISITES for the full list and
       the registry override that bypasses Outlook's 10-minute trust prompt.

.OUTPUTS
    Writes per-run logs into .\logs\cleanup\<timestamp>\:
        step1-flatten.log
        step2-consolidate.log
        step3-drafts.log
        summary.txt
    The summary.txt records exit codes and elapsed time per step.
#>

[CmdletBinding(DefaultParameterSetName = 'ByMapping')]
param(
    [Parameter(ParameterSetName = 'ByMailbox', Mandatory = $true)]
    [string[]]$Mailbox,

    [Parameter(ParameterSetName = 'ByMapping', Mandatory = $true)]
    [string]$MappingFile,

    [Parameter(Mandatory = $true)]
    [string]$Config,

    [string[]]$ExcludeMailbox,

    [switch]$DryRun,
    [switch]$SkipFlatten,
    [switch]$SkipConsolidate,
    [switch]$SkipDrafts,
    [switch]$ContinueOnError,

    [string]$LogDir,
    [string]$PythonExe
)

$ErrorActionPreference = 'Stop'
Set-StrictMode -Version Latest

# ---------------------------------------------------------------------------
# Setup
# ---------------------------------------------------------------------------

$ScriptRoot = Split-Path -Parent $PSCommandPath

# Resolve Python interpreter.
if (-not $PythonExe) {
    $PythonExe = Join-Path $ScriptRoot '.venv\Scripts\python.exe'
}
if (-not (Test-Path -LiteralPath $PythonExe)) {
    throw "Python interpreter not found: $PythonExe`nPass -PythonExe to override, or run the README 'Install' steps to create .venv."
}

# Resolve config + mapping to absolute paths so each child step gets a
# stable path regardless of cwd.
$ConfigPath = (Resolve-Path -LiteralPath $Config).Path
$SelectionMode = $PSCmdlet.ParameterSetName   # 'ByMailbox' or 'ByMapping'

# Resolve the selected list of UPNs.  Whether the user passed -Mailbox
# or -MappingFile, downstream we always pass an explicit list of
# --mailbox / -Mailbox args so the wrapper has full control over what
# each step sees (including -ExcludeMailbox filtering).
if ($SelectionMode -eq 'ByMailbox') {
    $SelectedMailboxes = @($Mailbox)
    $MappingPath = $null
} else {
    if (-not (Test-Path -LiteralPath $MappingFile)) {
        throw "Mapping file not found: $MappingFile"
    }
    $MappingPath = (Resolve-Path -LiteralPath $MappingFile).Path

    # Pull TargetMailbox column from CSV preserving file order.  We do
    # NOT pass mapping.csv directly to child scripts -- the wrapper's
    # selection (with -ExcludeMailbox applied) is the source of truth.
    $rows = Import-Csv -LiteralPath $MappingPath
    $missingCol = $rows | Where-Object { -not $_.PSObject.Properties['TargetMailbox'] } | Select-Object -First 1
    if ($missingCol) {
        throw "Mapping file '$MappingPath' is missing the TargetMailbox column."
    }
    $SelectedMailboxes = @($rows | ForEach-Object { $_.TargetMailbox } | Where-Object { $_ })
}

# Apply -ExcludeMailbox.  Comparison is case-insensitive UPN match.
if ($ExcludeMailbox) {
    $excludeSet = New-Object 'System.Collections.Generic.HashSet[string]' ([System.StringComparer]::OrdinalIgnoreCase)
    foreach ($x in $ExcludeMailbox) {
        if ($x) { [void]$excludeSet.Add($x.Trim()) }
    }

    $beforeCount = $SelectedMailboxes.Count
    $SelectedMailboxes = @($SelectedMailboxes | Where-Object { -not $excludeSet.Contains($_) })
    $removedCount = $beforeCount - $SelectedMailboxes.Count

    # Warn (not fatal) on any -ExcludeMailbox value that didn't match
    # anything in the source list -- almost always a typo.
    if ($SelectionMode -eq 'ByMapping') {
        $sourceUpns = @($rows | ForEach-Object { $_.TargetMailbox } | Where-Object { $_ })
    } else {
        $sourceUpns = @($Mailbox)
    }
    $sourceSet = New-Object 'System.Collections.Generic.HashSet[string]' ([System.StringComparer]::OrdinalIgnoreCase)
    foreach ($u in $sourceUpns) { [void]$sourceSet.Add($u) }
    foreach ($x in $excludeSet) {
        if (-not $sourceSet.Contains($x)) {
            Write-Warning "ExcludeMailbox '$x' did not match any selected mailbox -- typo?"
        }
    }

    Write-Host ("ExcludeMailbox: removed {0} of {1} mailbox(es) from selection." -f $removedCount, $beforeCount) -ForegroundColor Yellow
}

if (-not $SelectedMailboxes -or $SelectedMailboxes.Count -eq 0) {
    throw 'Empty mailbox selection after filtering. Nothing to do.'
}

# Per-run log directory.
if (-not $LogDir) {
    $LogDir = Join-Path $ScriptRoot 'logs\cleanup'
}
$RunStamp = Get-Date -Format 'yyyyMMdd-HHmmss'
$RunDir   = Join-Path $LogDir $RunStamp
New-Item -ItemType Directory -Force -Path $RunDir | Out-Null

$SummaryPath     = Join-Path $RunDir 'summary.txt'
$LiveStatusFile  = Join-Path $RunDir 'LIVE-STATUS.txt'
# Child processes (_flatten, _consolidate) read this path from the environment
# and rewrite the file on every milestone so the operator can open one file
# for *current* state mid-run.
$env:JTET_LIVE_STATUS_FILE = $LiveStatusFile

# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

function Write-Banner {
    param([string]$Message)
    $line = '=' * ([math]::Max(60, $Message.Length + 8))
    Write-Host ''
    Write-Host $line -ForegroundColor Cyan
    Write-Host "    $Message" -ForegroundColor Cyan
    Write-Host $line -ForegroundColor Cyan
}

function Add-Summary {
    param([string]$Line)
    Add-Content -LiteralPath $SummaryPath -Value $Line
}

function Get-MailboxArgs {
    <#
        Build the --mailbox / -Mailbox pieces for the active selection.
        Returns @{ PythonArgs = @(...); PSArgs = @{...} }.

        We always materialise an explicit --mailbox list rather than
        passing mapping.csv straight through, so that -ExcludeMailbox
        filtering applies uniformly to every step (Python and PS).
        Reads $SelectedMailboxes from script scope.
    #>
    $py = @()
    foreach ($m in $SelectedMailboxes) { $py += @('--mailbox', $m) }
    return @{
        PythonArgs = $py
        PSArgs     = @{ Mailbox = [string[]]$SelectedMailboxes }
    }
}

function Invoke-Step {
    <#
        Runs one cleanup step, tees its output to a log file, and
        returns a small status object.  Honors -ContinueOnError.
    #>
    param(
        [Parameter(Mandatory)][string]$Name,
        [Parameter(Mandatory)][string]$LogFile,
        [Parameter(Mandatory)][scriptblock]$Action
    )

    Write-Banner $Name
    $sw = [Diagnostics.Stopwatch]::StartNew()
    $logPath = Join-Path $RunDir $LogFile

    try {
        # Tee child output to console + per-step log.  We use Start-Transcript
        # on a sub-scope so that anything the action emits to the success
        # stream, warning stream, and stderr is captured.
        Start-Transcript -LiteralPath $logPath -Force | Out-Null
        try {
            & $Action
            $exit = $LASTEXITCODE
            if ($null -eq $exit) { $exit = 0 }
        } finally {
            Stop-Transcript | Out-Null
        }
    } catch {
        $exit = 1
        Write-Host "ERROR: $($_.Exception.Message)" -ForegroundColor Red
        Add-Content -LiteralPath $logPath -Value "ERROR: $($_.Exception.Message)"
    }

    $sw.Stop()
    $elapsed = '{0:c}' -f $sw.Elapsed
    if ($exit -eq 0) {
        $status = 'OK'
        $color  = 'Green'
    } else {
        $status = "FAIL ($exit)"
        $color  = 'Red'
    }
    Write-Host ("[{0}] {1} in {2}" -f $status, $Name, $elapsed) -ForegroundColor $color
    Add-Summary ("{0,-32} {1,-12} {2,8} {3}" -f $Name, $status, $elapsed, $logPath)

    if ($exit -ne 0 -and -not $ContinueOnError) {
        throw "$Name failed with exit code $exit. See $logPath. Re-run with -ContinueOnError to push past failures."
    }

    return @{ Name = $Name; Exit = $exit; Elapsed = $sw.Elapsed; Log = $logPath }
}

# ---------------------------------------------------------------------------
# Plan + run
# ---------------------------------------------------------------------------

Write-Banner 'JTET full mailbox cleanup'

$selection = Get-MailboxArgs
if ($SelectionMode -eq 'ByMapping') {
    $sourceStr = " (from $MappingPath"
    if ($ExcludeMailbox) { $sourceStr += ", minus -ExcludeMailbox" }
    $sourceStr += ")"
} else {
    $sourceStr = ' (from -Mailbox args)'
}
$selectStr = ("{0} mailbox(es){1}:" -f $SelectedMailboxes.Count, $sourceStr)

$plan = @()
if (-not $SkipFlatten)     { $plan += '1. Flatten "Imported PST" -> root  (_flatten_imported.py)' }
if (-not $SkipConsolidate) { $plan += '2. Consolidate Sent Items         (_consolidate_sent.py)' }
if (-not $SkipDrafts)      { $plan += '3. Clear MSGFLAG_UNSENT bit       (Fix-DraftsViaOutlook.ps1)' }

if ($plan.Count -eq 0) {
    throw 'All three steps were skipped via -SkipFlatten / -SkipConsolidate / -SkipDrafts. Nothing to do.'
}

# Children (_flatten, _consolidate, Fix-DraftsViaOutlook) read these for LIVE-STATUS.txt
# overall_run_approx: equal weight per enabled step + mailbox progress within the step.
$env:JTET_LIVE_PIPELINE_TOTAL_STEPS = "$($plan.Count)"
$pipelineStepNum = 0

Write-Host ''
Write-Host "Selection : $selectStr"
foreach ($m in $SelectedMailboxes) { Write-Host "              - $m" }
Write-Host "Config    : $ConfigPath"
Write-Host "Python    : $PythonExe"
Write-Host "Logs      : $RunDir"
Write-Host "DryRun    : $DryRun"
Write-Host ''
Write-Host 'Plan:'
$plan | ForEach-Object { Write-Host "    $_" }
Write-Host ''

Add-Summary "JTET full cleanup run @ $RunStamp"
Add-Summary "Selection : $selectStr"
foreach ($m in $SelectedMailboxes) { Add-Summary "              - $m" }
Add-Summary "Config    : $ConfigPath"
Add-Summary "DryRun    : $DryRun"
Add-Summary ''
Add-Summary ('{0,-32} {1,-12} {2,8} {3}' -f 'Step', 'Status', 'Elapsed', 'Log')
Add-Summary ('{0,-32} {1,-12} {2,8} {3}' -f ('-' * 32), ('-' * 12), ('-' * 8), ('-' * 30))

$results = @()

Write-Host ''
Write-Host "LIVE STATUS FILE (open in Notepad; refreshes as the run works):" -ForegroundColor Green
Write-Host "  $LiveStatusFile" -ForegroundColor Green
Write-Host '  overall_run_approx: rough % through the full run (in LIVE-STATUS.txt; equal weight per step, not wall-clock).' -ForegroundColor DarkGray
Write-Host ""

# --- Step 1: Flatten -------------------------------------------------------

if (-not $SkipFlatten) {
    $pipelineStepNum++
    $env:JTET_LIVE_PIPELINE_STEP_INDEX = "$pipelineStepNum"
    Write-Host ''
    Write-Host ">>> PIPELINE STEP $pipelineStepNum of $($plan.Count): Flatten (_flatten_imported.py) — Microsoft Graph, mailboxes run in PARALLEL." -ForegroundColor Yellow
    Write-Host ">>> Log: $RunDir\step1-flatten.log  |  Search for:  PROGRESS:  to see how many of $($SelectedMailboxes.Count) mailboxes have FINISHED (each line = one mailbox done)." -ForegroundColor Yellow
    Write-Host ">>> While a large Inbox merge runs, you also get  moved N messages  every 200 msgs. Graph 500 + backoff is throttling; usually succeeds after retries." -ForegroundColor DarkGray
    Write-Host ''
    $results += Invoke-Step `
        -Name    'Step 1: Flatten Imported PST' `
        -LogFile 'step1-flatten.log' `
        -Action  {
            $env:JTET_LIVE_STATUS_FILE = $LiveStatusFile
            $pyArgs = @('_flatten_imported.py', '-c', $ConfigPath) + $selection.PythonArgs
            if ($DryRun) { $pyArgs += '--dry-run' }
            & $PythonExe @pyArgs
        }
}

# --- Step 2: Consolidate Sent ---------------------------------------------

if (-not $SkipConsolidate) {
    $pipelineStepNum++
    $env:JTET_LIVE_PIPELINE_STEP_INDEX = "$pipelineStepNum"
    Write-Host ''
    Write-Host ">>> PIPELINE STEP $pipelineStepNum of $($plan.Count): Consolidate Sent (_consolidate_sent.py) — parallel mailboxes, same PROGRESS: k/N idea in step2-consolidate.log." -ForegroundColor Yellow
    Write-Host ''
    $results += Invoke-Step `
        -Name    'Step 2: Consolidate Sent Items' `
        -LogFile 'step2-consolidate.log' `
        -Action  {
            $env:JTET_LIVE_STATUS_FILE = $LiveStatusFile
            $pyArgs = @('_consolidate_sent.py', '-c', $ConfigPath) + $selection.PythonArgs
            if ($DryRun) { $pyArgs += '--dry-run' }
            & $PythonExe @pyArgs
        }
}

# --- Step 3: Fix drafts via Outlook ---------------------------------------

if (-not $SkipDrafts) {
    $pipelineStepNum++
    $env:JTET_LIVE_PIPELINE_STEP_INDEX = "$pipelineStepNum"
    $results += Invoke-Step `
        -Name    'Step 3: Clear MSGFLAG_UNSENT' `
        -LogFile 'step3-drafts.log' `
        -Action  {
            $env:JTET_LIVE_STATUS_FILE = $LiveStatusFile
            $pIdx  = $env:JTET_LIVE_PIPELINE_STEP_INDEX
            $pTot  = $env:JTET_LIVE_PIPELINE_TOTAL_STEPS
            $env:JTET_LIVE_PIPELINE_LABEL = "Step $pIdx of $pTot (Outlook MAPI — usually slowest)"
            Write-Host ""
            Write-Host ("========== RUN-FULL-CLEANUP: pipeline step {0} of {1} (Outlook / MAPI) ==========" -f $pIdx, $pTot) -ForegroundColor Cyan
            Write-Host "LIVE-STATUS: $LiveStatusFile (updated per mailbox in this step)`n" -ForegroundColor Cyan
            $childScript = Join-Path $ScriptRoot 'Fix-DraftsViaOutlook.ps1'
            $params = @{} + $selection.PSArgs
            if ($DryRun) { $params['DryRun'] = $true }
            & $childScript @params
        }
}

# ---------------------------------------------------------------------------
# Summary
# ---------------------------------------------------------------------------

Write-Banner 'Cleanup summary'

foreach ($r in $results) {
    if ($r.Exit -eq 0) {
        $rowStatus = 'OK'
        $rowColor  = 'Green'
    } else {
        $rowStatus = "FAIL ($($r.Exit))"
        $rowColor  = 'Red'
    }
    Write-Host ('  {0,-32} {1,-10} {2}' -f $r.Name, $rowStatus, ('{0:c}' -f $r.Elapsed)) `
        -ForegroundColor $rowColor
}

Write-Host ''
Write-Host "Per-step logs : $RunDir"
Write-Host "Summary file  : $SummaryPath"

$failed = $results | Where-Object { $_.Exit -ne 0 }
if ($failed) {
    Write-Host ''
    Write-Host ("{0} step(s) failed." -f $failed.Count) -ForegroundColor Red
    exit 1
}

Write-Host ''
Write-Host 'All steps completed successfully.' -ForegroundColor Green
Add-Summary ''
Add-Summary ("Run complete. Live status was: $LiveStatusFile")
exit 0
