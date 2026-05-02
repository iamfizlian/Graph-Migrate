<#
.SYNOPSIS
  Live colorized monitor for an in-progress pstmigrate run.

.DESCRIPTION
  Tails the most recently-modified import_*.jsonl in logs/, color-codes
  every line by level + context, and prints a running stats banner every
  ~30 seconds (uploads / folders / throttling backoffs / errors / per-app /
  per-mailbox progress from state.sqlite).

  Uses a polling tail loop so the banner and heartbeat fire on time even
  when the log is silent (e.g. during a long PST extract or backoff).

.PARAMETER LogPath
  Optional path to a specific JSONL file. If omitted, auto-picks the newest
  import_*.jsonl in .\logs\.

.PARAMETER Tail
  How many recent log lines to replay on startup so you don't stare at a
  blank screen. Default: 50. Use 0 for "only new lines from now".

.PARAMETER ShowDebug
  Include DEBUG-level lines (token refreshes etc.). Off by default.

.PARAMETER NoSql
  Skip the SQLite per-mailbox progress block at each banner tick.

.PARAMETER StateDb
  Path to state.sqlite (default: .pstmigrate-state\state.sqlite).

.PARAMETER BannerSeconds
  Seconds between stats banners. Default: 30.

.PARAMETER PollMs
  How often to poll the log file for new content. Default: 500.

.EXAMPLE
  .\Watch-Migration.ps1
  .\Watch-Migration.ps1 -ShowDebug -Tail 200
  .\Watch-Migration.ps1 -LogPath logs\import_20260425_014700.jsonl
#>
[CmdletBinding()]
param(
    [string]$LogPath,
    [int]$Tail = 50,
    [switch]$ShowDebug,
    [switch]$NoSql,
    [string]$StateDb = ".pstmigrate-state\state.sqlite",
    [int]$BannerSeconds = 30,
    [int]$PollMs = 500
)

$ErrorActionPreference = "Stop"

# ---- Pick a log file -------------------------------------------------------
if (-not $LogPath) {
    $latest = Get-ChildItem -Path "logs" -Filter "import_*.jsonl" -ErrorAction SilentlyContinue |
              Sort-Object LastWriteTime -Descending |
              Select-Object -First 1
    if (-not $latest) {
        Write-Host "No import_*.jsonl found in .\logs\. Is a migration running?" -ForegroundColor Red
        exit 1
    }
    $LogPath = $latest.FullName
}

if (-not (Test-Path $LogPath)) {
    Write-Host "Log file not found: $LogPath" -ForegroundColor Red
    exit 1
}

# ---- Colors ---------------------------------------------------------------
$LevelColor = @{
    "DEBUG"   = "DarkGray"
    "INFO"    = "White"
    "WARNING" = "Yellow"
    "ERROR"   = "Red"
}

function Get-CtxColor([string]$ctx) {
    if ($null -eq $ctx)        { return "Gray" }
    if ($ctx -like "pst*")     { return "Cyan" }
    if ($ctx -like "folders*") { return "Magenta" }
    if ($ctx -like "graph*")   { return "Blue" }
    if ($ctx -like "job*")     { return "Green" }
    if ($ctx -like "upload*")  { return "DarkGreen" }
    if ($ctx -like "auth*")    { return "DarkGray" }
    return "Gray"
}

# ---- Counters --------------------------------------------------------------
$script:Counters = @{
    Folders     = 0
    Throttles   = 0
    Errors      = 0
    Extracts    = 0
    Reuses      = 0
    Uploads     = 0
    Started     = Get-Date
    LastBanner  = [DateTime]::MinValue
    LastHeartbeat = Get-Date
    LinesSeen   = 0
    PerApp      = @{}
}

$BannerInterval = [TimeSpan]::FromSeconds($BannerSeconds)
$HeartbeatInterval = [TimeSpan]::FromSeconds(15)

# ---- Helpers --------------------------------------------------------------
function Format-HumanDuration([TimeSpan]$ts) {
    if ($ts.TotalHours -ge 1) {
        return "{0:N0}h {1:N0}m" -f [Math]::Floor($ts.TotalHours), $ts.Minutes
    }
    if ($ts.TotalMinutes -ge 1) {
        return "{0:N0}m {1:N0}s" -f [Math]::Floor($ts.TotalMinutes), $ts.Seconds
    }
    return "{0:N0}s" -f $ts.TotalSeconds
}

function Get-MailboxProgress {
    if ($NoSql) { return $null }
    if (-not (Test-Path $StateDb)) { return $null }

    $sqlScript = @"
import sqlite3, sys, json, os
db = r"$StateDb"
try:
    c = sqlite3.connect(db)
    c.execute("PRAGMA busy_timeout = 2000")
    totals = {}
    for mbx, status, n in c.execute(
        "SELECT target_mailbox, status, COUNT(1) FROM messages GROUP BY target_mailbox, status"
    ):
        totals.setdefault(mbx, {})[status] = n
    pst_status = {}
    for mbx, status, items in c.execute(
        "SELECT target_mailbox, status, items_total FROM pst_runs"
    ):
        pst_status.setdefault(mbx, []).append({"s": status, "t": items or 0})
    out = []
    for mbx in sorted(set(list(totals.keys()) + list(pst_status.keys()))):
        s = totals.get(mbx, {})
        out.append({
            "m": mbx,
            "d": s.get("done", 0),
            "f": s.get("failed", 0),
            "s": s.get("skipped", 0),
            "p": pst_status.get(mbx, []),
        })
    print(json.dumps(out))
except Exception as e:
    print(json.dumps({"error": str(e)}))
"@

    $tmp = New-TemporaryFile
    try {
        Set-Content $tmp.FullName $sqlScript -Encoding UTF8
        $py = ".\.venv\Scripts\python.exe"
        if (-not (Test-Path $py)) { return $null }
        $json = & $py $tmp.FullName 2>$null
        if (-not $json) { return $null }
        return ($json | ConvertFrom-Json)
    } catch {
        return $null
    } finally {
        Remove-Item $tmp.FullName -Force -ErrorAction SilentlyContinue
    }
}

function Write-Banner {
    $now = Get-Date
    $elapsed = $now - $script:Counters.Started
    $rule = "-" * 78
    Write-Host ""
    Write-Host $rule -ForegroundColor DarkGray
    $hdr = " STATS @ {0:HH:mm:ss}   elapsed {1}   log lines seen: {2}" -f `
            $now, (Format-HumanDuration $elapsed), $script:Counters.LinesSeen
    Write-Host $hdr -ForegroundColor White

    $line = "  Uploads done : {0,-6}   Folders : {1,-4}   PST extracts : {2,-3}   reused : {3,-3}" -f `
            $script:Counters.Uploads, $script:Counters.Folders, $script:Counters.Extracts, $script:Counters.Reuses
    Write-Host $line -ForegroundColor Gray

    $line = "  Throttle backoffs: {0,-6}   Errors : {1,-4}" -f `
            $script:Counters.Throttles, $script:Counters.Errors
    $color = if ($script:Counters.Throttles -gt 0) { "Yellow" } elseif ($script:Counters.Errors -gt 0) { "Red" } else { "Gray" }
    Write-Host $line -ForegroundColor $color

    if ($script:Counters.PerApp.Count -gt 0) {
        Write-Host "  Per-app throttling:" -ForegroundColor Gray
        foreach ($name in ($script:Counters.PerApp.Keys | Sort-Object)) {
            $st = $script:Counters.PerApp[$name]
            $line = "    {0,-15}  throttles {1,-4}  total backoff {2,7:N1}s" -f `
                    $name, $st.Throttles, $st.BackoffSec
            $c = if ($st.Throttles -gt 5) { "Yellow" } else { "DarkGray" }
            Write-Host $line -ForegroundColor $c
        }
    }

    $rows = Get-MailboxProgress
    if ($rows -and -not ($rows.PSObject.Properties.Name -contains "error")) {
        if ($rows.Count -gt 0) {
            Write-Host "  Per-mailbox progress (state.sqlite):" -ForegroundColor Gray
            foreach ($r in $rows) {
                $shortMbx = ($r.m -split "@")[0]
                $line = "    {0,-15}  done {1,-6}  failed {2,-3}  skipped {3,-3}" -f `
                        $shortMbx, $r.d, $r.f, $r.s
                $c = if ($r.f -gt 0) { "Yellow" } elseif ($r.d -gt 0) { "Green" } else { "DarkGray" }
                Write-Host $line -ForegroundColor $c
                if ($r.p) {
                    foreach ($p in $r.p) {
                        $pcol = switch ($p.s) {
                            "done"      { "Green" }
                            "uploading" { "Cyan" }
                            "extracting"{ "Yellow" }
                            "error"     { "Red" }
                            default     { "DarkGray" }
                        }
                        Write-Host ("        pst -> {0,-11}  total={1}" -f $p.s, $p.t) -ForegroundColor $pcol
                    }
                }
            }
        }
    } elseif ($rows -and ($rows.PSObject.Properties.Name -contains "error")) {
        Write-Host ("  SQLite read failed: {0}" -f $rows.error) -ForegroundColor DarkYellow
    }

    Write-Host $rule -ForegroundColor DarkGray
    Write-Host ""
    $script:Counters.LastBanner = $now
}

function Update-Counters([object]$record) {
    if (-not $record) { return }
    $msg = [string]$record.message
    $ctx = [string]$record.extra.ctx
    $level = [string]$record.level.name

    $script:Counters.LinesSeen++

    if ($level -eq "ERROR") { $script:Counters.Errors++ }

    if ($ctx -like "graph*") {
        if ($msg -match "Backoff\s+([\d.]+)s.*throttled") {
            $script:Counters.Throttles++
            $sec = [double]$Matches[1]
            $appName = ($ctx -replace '^graph\[(.+)\]$', '$1')
            if (-not $script:Counters.PerApp.ContainsKey($appName)) {
                $script:Counters.PerApp[$appName] = @{ Throttles = 0; BackoffSec = 0.0 }
            }
            $script:Counters.PerApp[$appName].Throttles++
            $script:Counters.PerApp[$appName].BackoffSec += $sec
        }
    }
    elseif ($ctx -like "folders*") {
        if ($msg -like "Created folder*") { $script:Counters.Folders++ }
    }
    elseif ($ctx -eq "pst") {
        if ($msg -match "^Extracting .+\.pst .* -> ") { $script:Counters.Extracts++ }
        if ($msg -like "Reusing previous extraction*") { $script:Counters.Reuses++ }
    }
    elseif ($ctx -like "upload*") {
        if ($msg -like "Uploaded *" -or $msg -like "*upload ok*") { $script:Counters.Uploads++ }
    }
}

function Format-Timestamp([string]$repr) {
    if ($repr -match "(\d{2}:\d{2}:\d{2})") { return $Matches[1] }
    return "??:??:??"
}

function Write-LogLine([object]$record) {
    if (-not $record) { return }
    $level = [string]$record.level.name
    if (-not $ShowDebug -and $level -eq "DEBUG") { return }
    $ctx = [string]$record.extra.ctx
    $msg = [string]$record.message
    $time = Format-Timestamp $record.time.repr

    if ([string]::IsNullOrEmpty($ctx)) { $ctx = "-" }
    $ctxShort = if ($ctx.Length -gt 30) { $ctx.Substring(0, 27) + "..." } else { $ctx }

    Write-Host -NoNewline ("{0}  " -f $time) -ForegroundColor DarkGray
    $lvlColor = $LevelColor[$level]; if (-not $lvlColor) { $lvlColor = "Gray" }
    Write-Host -NoNewline ("{0,-7} " -f $level) -ForegroundColor $lvlColor
    Write-Host -NoNewline ("{0,-30} " -f $ctxShort) -ForegroundColor (Get-CtxColor $ctx)

    $color = "White"
    if     ($msg -like "*throttled*")                   { $color = "Yellow" }
    elseif ($msg -like "*Created folder*")              { $color = "Magenta" }
    elseif ($msg -like "*Reusing previous extraction*") { $color = "Cyan" }
    elseif ($msg -like "*Extracting*")                  { $color = "Cyan" }
    elseif ($msg -like "*Extracted*")                   { $color = "Green" }
    elseif ($msg -like "*messages to consider*")        { $color = "Green" }
    elseif ($msg -like "*Uploaded *" -or $msg -like "*upload ok*") { $color = "DarkGreen" }
    elseif ($msg -like "*Failed*" -or $msg -like "*FAIL*")         { $color = "Red" }
    Write-Host $msg -ForegroundColor $color
}

function Process-RawLine([string]$line) {
    if ([string]::IsNullOrWhiteSpace($line)) { return }
    try {
        $parsed = $line | ConvertFrom-Json -ErrorAction Stop
    } catch {
        return
    }
    $rec = $parsed.record
    if (-not $rec) { return }
    Update-Counters $rec
    Write-LogLine $rec
}

# ---- Header --------------------------------------------------------------
Clear-Host
Write-Host ""
Write-Host "  pstmigrate live monitor" -ForegroundColor White
Write-Host ("  Tailing : {0}" -f $LogPath) -ForegroundColor DarkGray
Write-Host ("  Tail    : {0} startup lines | banner every {1}s | poll {2}ms | DEBUG {3}" -f `
            $Tail, $BannerSeconds, $PollMs, $(if ($ShowDebug) {"on"} else {"off"})) -ForegroundColor DarkGray
Write-Host "  Ctrl-C to stop watching (the migration keeps running in its own window)" -ForegroundColor DarkGray
Write-Host ""

# ---- Replay last N lines so the screen isn't blank -----------------------
if ($Tail -gt 0) {
    Write-Host ("--- replaying last {0} log lines ---" -f $Tail) -ForegroundColor DarkGray
    Get-Content -LiteralPath $LogPath -Tail $Tail -ErrorAction SilentlyContinue |
        ForEach-Object { Process-RawLine $_ }
    Write-Host "--- live tail begins ---" -ForegroundColor DarkGray
}

# Force an initial banner immediately so the user sees stats right away
Write-Banner

# ---- Polling tail loop ---------------------------------------------------
$lastSize = (Get-Item -LiteralPath $LogPath).Length
$leftover = ""

while ($true) {
    Start-Sleep -Milliseconds $PollMs

    try {
        $info = Get-Item -LiteralPath $LogPath -ErrorAction Stop
    } catch {
        Write-Host ("  [watch] log disappeared: {0}" -f $LogPath) -ForegroundColor Red
        Start-Sleep -Seconds 2
        continue
    }
    $size = $info.Length

    if ($size -lt $lastSize) {
        # File rotated or truncated - start over
        Write-Host "  [watch] log truncated/rotated; resetting offset" -ForegroundColor DarkYellow
        $lastSize = 0
        $leftover = ""
    }

    if ($size -gt $lastSize) {
        try {
            $fs = [System.IO.FileStream]::new(
                $LogPath,
                [System.IO.FileMode]::Open,
                [System.IO.FileAccess]::Read,
                [System.IO.FileShare]::ReadWrite -bor [System.IO.FileShare]::Delete
            )
            $null = $fs.Seek($lastSize, [System.IO.SeekOrigin]::Begin)
            $reader = [System.IO.StreamReader]::new($fs, [System.Text.Encoding]::UTF8)
            $chunk = $reader.ReadToEnd()
            $reader.Dispose()
            $fs.Dispose()
        } catch {
            Write-Host ("  [watch] read error: {0}" -f $_.Exception.Message) -ForegroundColor Red
            continue
        }

        $combined = $leftover + $chunk
        $lines = $combined -split "`r?`n"
        # If chunk did not end on a newline, the last element is a partial line
        $endsWithNewline = $combined.EndsWith("`n") -or $combined.EndsWith("`r")
        if (-not $endsWithNewline -and $lines.Length -gt 0) {
            $leftover = $lines[-1]
            if ($lines.Length -gt 1) {
                $lines = $lines[0..($lines.Length - 2)]
            } else {
                $lines = @()
            }
        } else {
            $leftover = ""
        }

        foreach ($ln in $lines) {
            Process-RawLine $ln
        }
        $lastSize = $size
        $script:Counters.LastHeartbeat = Get-Date
    }

    $now = Get-Date
    if ($now - $script:Counters.LastBanner -ge $BannerInterval) {
        Write-Banner
    }
    elseif ($now - $script:Counters.LastHeartbeat -ge $HeartbeatInterval) {
        Write-Host ("  [{0:HH:mm:ss}] watching... (no new log activity for {1})" -f `
                    $now, (Format-HumanDuration ($now - $script:Counters.LastHeartbeat))) -ForegroundColor DarkGray
        $script:Counters.LastHeartbeat = $now
    }
}
