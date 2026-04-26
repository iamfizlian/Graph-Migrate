<#
.SYNOPSIS
    Clear the MSGFLAG_UNSENT bit on imported messages via Outlook desktop
    (MAPI/HTTP).  This is the only API path on Exchange Online that can
    actually mutate PR_MESSAGE_FLAGS post-import; both Microsoft Graph
    and EWS silently no-op the write while reporting Success.

.DESCRIPTION
    PSTs imported via Graph POST /messages land with PR_MESSAGE_FLAGS
    bit 0x08 (MSGFLAG_UNSENT) set, causing every imported message to
    appear as a draft in Outlook on the Web.  We've empirically verified
    in this repo that the cloud's MAPI store rejects writes to this
    property when issued from any of:

        - Graph PATCH {"isDraft": false}
        - Graph PATCH singleValueExtendedProperties Integer 0x0E07
        - Graph POST /messages/{id}/copy
        - EWS UpdateItem SetItemField  (with several values)
        - EWS UpdateItem DeleteItemField
        - EWS UpdateItem multi-property SetItemField
        - EWS CreateItem MimeContent + SavedItemFolderId=sentitems

    Outlook desktop talks to Exchange Online over MAPI/HTTP, a separate
    protocol that DOES allow store-level property writes (it's the same
    one MFCMAPI / OutlookSpy use).  This script attaches to (or launches)
    Outlook on the local machine, opens each target mailbox via delegate
    permission, and clears the bit one item at a time.

    Performance: ~25 items/sec on a typical desktop; 485 drafts in one
    mailbox finishes in under 30 seconds.

.PREREQUISITES
    1. Windows machine with Outlook desktop installed (2016 / 2019 /
       365 / LTSC).
    2. **Open Outlook BEFORE running this script** and let it finish
       syncing.  The script can cold-launch Outlook for you, but
       Windows then shows a "Programmatic access" / "Allow access"
       trust dialog, and the delegated mailboxes are still being
       attached in the background while we try to enumerate folders.
       Both can lead to "Inbox.Parent was null" errors.  Easiest is
       to start Outlook, click through any prompts, wait for the send/
       receive indicator to settle, and only then run this script.
    3. Outlook signed in to a profile.  The account on that profile
       must have FullAccess on every target mailbox, granted with
       -AutoMapping $false so the mailboxes are NOT auto-attached:

           Connect-ExchangeOnline
           $admin = "your-admin@jteatono365.onmicrosoft.com"
           @( "allysonp@jteatono365.onmicrosoft.com",
              "debbiep@jteatono365.onmicrosoft.com",
              # ...etc
           ) | ForEach-Object {
               Add-MailboxPermission -Identity $_ -User $admin `
                   -AccessRights FullAccess -InheritanceType All `
                   -AutoMapping $false
           }

       AutoMapping=false matters: with it $true, every imported mailbox
       gets auto-attached to the admin's Outlook profile, slowing
       startup and cluttering the folder pane.  We open them
       programmatically via Namespace.GetSharedDefaultFolder instead.

.PARAMETER Mailbox
    UPN of a target mailbox.  Repeatable.

.PARAMETER MappingFile
    Path to a CSV with a TargetMailbox column (the project's mapping.csv
    works as-is).  Mutually exclusive with -Mailbox.

.PARAMETER DryRun
    Walk, count, log; issue zero PropertyAccessor.SetProperty calls.

.PARAMETER Folder
    Display-name path under the mailbox root to scope the walk to,
    e.g. "Sent Items" or "Inbox/Imported PST".  Path components are
    case-insensitive and separated by "/".  If omitted, every IPF.Note
    folder in the mailbox is walked except the well-known Drafts,
    Deleted Items, Outbox, and Junk Email folders.

.EXAMPLE
    Try one mailbox in dry-run first to confirm counts:

        .\Fix-DraftsViaOutlook.ps1 `
            -Mailbox allysonp@jteatono365.onmicrosoft.com -DryRun

.EXAMPLE
    Run the fix on one mailbox, then verify with the existing Python
    inspector:

        .\Fix-DraftsViaOutlook.ps1 -Mailbox allysonp@jteatono365.onmicrosoft.com
        .\.venv\Scripts\python.exe _inspect_dates.py -c config.toml `
            --mailbox allysonp@jteatono365.onmicrosoft.com --folder sentitems --top 5

.EXAMPLE
    Once you're happy, fix every mailbox in mapping.csv:

        .\Fix-DraftsViaOutlook.ps1 -MappingFile .\mapping.csv
#>
[CmdletBinding(DefaultParameterSetName = 'ByMailbox')]
param(
    [Parameter(ParameterSetName = 'ByMailbox', Mandatory = $true)]
    [string[]]$Mailbox,

    [Parameter(ParameterSetName = 'ByMapping', Mandatory = $true)]
    [string]$MappingFile,

    [Parameter()]
    [switch]$DryRun,

    [Parameter()]
    [string]$Folder
)

# ---------------------------------------------------------------------------
# Constants
# ---------------------------------------------------------------------------

# MAPI property tag, expressed as the PropertyAccessor schema URI Outlook
# wants. 0x0E07 is PR_MESSAGE_FLAGS; 0003 is the PT_LONG type code.
$script:PR_MESSAGE_FLAGS_TAG = "http://schemas.microsoft.com/mapi/proptag/0x0E070003"
$script:MSGFLAG_UNSENT       = 0x08

# OlDefaultFolders constants (we don't load the interop assembly so we use
# integer literals -- these values are stable across all Outlook versions).
$script:olFolderDeletedItems = 3
$script:olFolderOutbox       = 4
$script:olFolderSentMail     = 5
$script:olFolderInbox        = 6
$script:olFolderDrafts       = 16
$script:olFolderJunk         = 23

# DASL filter we hand to Items.Restrict() so Outlook only iterates items
# whose UNSENT bit is set rather than every message in the folder.
#
# Outlook's DASL parser is strict:
#   - no SQL LIKE; use CI_PHRASEMATCH / CI_STARTSWITH for substring
#   - no surrounding parentheses around the top-level NOT
#   - bitwise AND result is truthy when non-zero (no need for "= 0")
# So the simplest form that works is just:
#   @SQL="<proptag>" & 8
# We do the IPM.Note message-class check in the PowerShell loop as a
# belt-and-braces guard.
$script:DRAFT_RESTRICT = '@SQL="http://schemas.microsoft.com/mapi/proptag/0x0E070003" & 8'

# System / hidden folder names we never modify, regardless of whether
# they have DefaultItemType=0.  Case-insensitive whole-name match.
# Includes the well-known "real Drafts / Outbox / Junk / Trash" set
# (covered for delegated mailboxes where GetSharedDefaultFolder may
# not return a usable EntryID), plus Teams / Yammer / RSS / sync
# scratchpads that shouldn't have UNSENT messed with.
$script:SKIP_FOLDER_NAMES = @(
    'Drafts',
    'Outbox',
    'Deleted Items',
    'Junk Email', 'Junk', 'Junk E-mail',
    'Recoverable Items',
    'Yammer Root',
    'Conversation History',
    'Conversation Action Settings',
    'Sync Issues',
    'Conflicts',
    'Local Failures',
    'Server Failures',
    'RSS Feeds', 'RSS Subscriptions',
    'Quick Step Settings',
    'Files',                       # Teams scratch
    'ExternalContacts',
    'PersonMetadata'
)

# ---------------------------------------------------------------------------
# Outlook attach / launch
# ---------------------------------------------------------------------------

function Get-OutlookApplication {
    <#
    .SYNOPSIS  Attach to a running Outlook session if there is one,
    otherwise cold-launch one.  Returns a hashtable so the caller can
    tell whether MAPI needs a moment to settle before issuing
    GetSharedDefaultFolder calls.
    #>
    [OutputType([hashtable])]
    param()
    try {
        $existing = [System.Runtime.InteropServices.Marshal]::GetActiveObject('Outlook.Application')
        Write-Verbose "Attached to a running Outlook session."
        return @{ App = $existing; FreshLaunch = $false }
    } catch {
        Write-Host "Starting Outlook..." -ForegroundColor DarkGray
        $new = New-Object -ComObject Outlook.Application
        return @{ App = $new; FreshLaunch = $true }
    }
}

function Wait-NamespaceReady {
    <#
    .SYNOPSIS  Block until the MAPI namespace has at least one store
    loaded.  After a cold launch GetSharedDefaultFolder() will silently
    return $null while Outlook is still spinning up its session, which
    later surfaces as 'Cannot bind argument to parameter Root because
    it is null' deep inside the walker.
    #>
    param(
        [Parameter(Mandatory)] $Namespace,
        [int]$TimeoutSeconds = 30
    )
    $deadline = (Get-Date).AddSeconds($TimeoutSeconds)
    while ((Get-Date) -lt $deadline) {
        try {
            if ($Namespace.Stores.Count -gt 0 -and $Namespace.DefaultStore) {
                return $true
            }
        } catch {
            # Stores collection not available yet; keep waiting.
        }
        Start-Sleep -Milliseconds 500
    }
    return $false
}

# ---------------------------------------------------------------------------
# Mailbox helpers
# ---------------------------------------------------------------------------

function Get-SkipFolderIds {
    <#
    .SYNOPSIS  EntryIDs of the folders we never modify (real Drafts, Trash,
    Outbox, Junk).  We don't want to "fix" the actual Drafts folder.
    #>
    [OutputType([System.Collections.Generic.HashSet[string]])]
    param(
        [Parameter(Mandatory)] $Namespace,
        [Parameter(Mandatory)] $Recipient
    )
    $skip = [System.Collections.Generic.HashSet[string]]::new()
    foreach ($id in @(
        $script:olFolderDrafts,
        $script:olFolderDeletedItems,
        $script:olFolderOutbox,
        $script:olFolderJunk
    )) {
        try {
            $f = $Namespace.GetSharedDefaultFolder($Recipient, $id)
            if ($f) { [void]$skip.Add($f.EntryID) }
        } catch {
            # Some mailboxes don't have an explicit Outbox; ignore.
        }
    }
    return $skip
}

function Get-FolderByPath {
    <#
    .SYNOPSIS  Resolve a "Sent Items" / "Inbox/Imported PST" path under
    the given root folder.  Case-insensitive, "/" or "\" separated.
    #>
    [OutputType([__ComObject])]
    param(
        [Parameter(Mandatory)] $Root,
        [Parameter(Mandatory)][string]$Path
    )
    $parts = $Path -split '[\\/]+' | Where-Object { $_ -ne '' }
    $current = $Root
    foreach ($name in $parts) {
        $next = $null
        foreach ($child in $current.Folders) {
            if ($child.Name -ieq $name) { $next = $child; break }
        }
        if (-not $next) {
            throw "Folder path '$Path' not found under '$($Root.Name)' (stopped at '$name')."
        }
        $current = $next
    }
    return $current
}

# ---------------------------------------------------------------------------
# Recursive walk
# ---------------------------------------------------------------------------

function Get-MailFoldersRecursive {
    <#
    .SYNOPSIS  Yield every IPF.Note folder under $Root, skipping any
    folder whose EntryID is in $SkipIds or whose name is in
    $script:SKIP_FOLDER_NAMES.

    When called against a mailbox root (no -Folder filter), pass
    -IncludeRoot:$false so we don't try to Items.Restrict() the
    mailbox root itself -- it has no real mail and Outlook's
    DASL parser can choke on its hidden contents.
    #>
    [OutputType([__ComObject[]])]
    param(
        [Parameter(Mandatory)] $Root,
        [Parameter(Mandatory)] [System.Collections.Generic.HashSet[string]]$SkipIds,
        [bool]$IncludeRoot = $true
    )
    $out = New-Object 'System.Collections.Generic.List[object]'
    if ($IncludeRoot) {
        Walk-FolderInto $Root $SkipIds $out
    } else {
        foreach ($child in $Root.Folders) {
            Walk-FolderInto $child $SkipIds $out
        }
    }
    return $out.ToArray()
}

function Walk-FolderInto {
    param(
        $Folder,
        [System.Collections.Generic.HashSet[string]]$SkipIds,
        [System.Collections.Generic.List[object]]$Out
    )
    # Whole-subtree skip: any folder whose name matches a system folder
    # we never want to touch.  We also skip its descendants because
    # things like "Yammer Root\Inbound" or "Conversation History\Team Chat"
    # are scratch areas that contain non-IPM.Note items we shouldn't
    # rewrite.
    foreach ($skipName in $script:SKIP_FOLDER_NAMES) {
        if ($Folder.Name -ieq $skipName) { return }
    }
    if ($SkipIds.Contains($Folder.EntryID)) { return }

    # DefaultItemType 0 = olMailItem; folders that hold contacts/calendars
    # have a different value and we leave them alone.
    if ($Folder.DefaultItemType -eq 0) {
        $Out.Add($Folder)
    }

    foreach ($child in $Folder.Folders) {
        Walk-FolderInto $child $SkipIds $Out
    }
}

# ---------------------------------------------------------------------------
# Per-folder fix
# ---------------------------------------------------------------------------

function Fix-FolderItems {
    <#
    .SYNOPSIS  Clear MSGFLAG_UNSENT on every IPM.Note item in $Folder.
    Returns a hashtable with Found/Fixed/Failed counts.
    #>
    [OutputType([hashtable])]
    param(
        [Parameter(Mandatory)] $Folder,
        [Parameter(Mandatory)][bool]$IsDryRun
    )
    $found  = 0
    $fixed  = 0
    $failed = 0

    try {
        $drafts = $Folder.Items.Restrict($script:DRAFT_RESTRICT)
    } catch {
        Write-Warning "  $($Folder.FolderPath): Items.Restrict failed -- $($_.Exception.Message)"
        return @{ Found = 0; Fixed = 0; Failed = 0 }
    }

    # Iterate by index to avoid the "collection modified during enumeration"
    # behaviour you can hit when SetProperty causes Outlook to re-evaluate
    # the restriction mid-loop.  We walk a snapshot of EntryIDs first and
    # then re-fetch each item from $Folder by EntryID before mutating it.
    $entryIds = New-Object 'System.Collections.Generic.List[string]'
    $count = $drafts.Count
    for ($i = 1; $i -le $count; $i++) {
        try {
            $itm = $drafts.Item($i)
            if ($itm -and $itm.EntryID) { $entryIds.Add($itm.EntryID) }
        } catch {
            # Item couldn't be loaded (corrupted, in flight, etc.) -- skip.
        }
    }
    $found = $entryIds.Count

    if ($found -eq 0) {
        return @{ Found = 0; Fixed = 0; Failed = 0 }
    }

    if ($IsDryRun) {
        return @{ Found = $found; Fixed = 0; Failed = 0 }
    }

    $store = $Folder.Store
    foreach ($eid in $entryIds) {
        try {
            $itm = $store.GetItemFromID($eid)
            if (-not $itm) { $failed++; continue }
            if ($itm.MessageClass -and -not $itm.MessageClass.StartsWith('IPM.Note')) {
                # Restriction should already exclude these, but belt-and-braces.
                continue
            }
            $pa = $itm.PropertyAccessor
            $current = [int]$pa.GetProperty($script:PR_MESSAGE_FLAGS_TAG)
            if (($current -band $script:MSGFLAG_UNSENT) -eq 0) {
                # Already clean -- somebody else's process moved faster than us.
                continue
            }
            $newVal = $current -band (-bnot $script:MSGFLAG_UNSENT)
            $pa.SetProperty($script:PR_MESSAGE_FLAGS_TAG, $newVal)
            $itm.Save()
            $fixed++
        } catch {
            $failed++
            Write-Verbose "  item $eid failed: $($_.Exception.Message)"
        }
    }

    return @{ Found = $found; Fixed = $fixed; Failed = $failed }
}

# ---------------------------------------------------------------------------
# Per-mailbox driver
# ---------------------------------------------------------------------------

function Fix-Mailbox {
    [OutputType([pscustomobject])]
    param(
        [Parameter(Mandatory)] $Namespace,
        [Parameter(Mandatory)][string]$Upn,
        [Parameter(Mandatory)][bool]$IsDryRun,
        [Parameter()][string]$FolderFilter
    )

    $stats = [ordered]@{
        Mailbox  = $Upn
        Folders  = 0
        Drafts   = 0
        Fixed    = 0
        Failed   = 0
        Status   = 'ok'
    }

    try {
        $rcpt = $Namespace.CreateRecipient($Upn)
        $null = $rcpt.Resolve()
        if (-not $rcpt.Resolved) {
            $stats.Status = "ERROR: could not resolve recipient (FullAccess granted?)"
            return [pscustomobject]$stats
        }

        # Retry a few times on cold-launched Outlook -- the first
        # GetSharedDefaultFolder call can return $null while MAPI
        # is still attaching the delegated mailbox.
        $inbox = $null
        for ($attempt = 1; $attempt -le 6; $attempt++) {
            try {
                $inbox = $Namespace.GetSharedDefaultFolder($rcpt, $script:olFolderInbox)
            } catch {
                Write-Verbose ("GetSharedDefaultFolder attempt {0} threw: {1}" -f $attempt, $_.Exception.Message)
            }
            if ($inbox) { break }
            Start-Sleep -Seconds 2
        }
        if (-not $inbox) {
            $stats.Status = "ERROR: GetSharedDefaultFolder returned null for $Upn (FullAccess granted with -AutoMapping `$false?  Or is Outlook still starting?)"
            return [pscustomobject]$stats
        }

        # Resolve the mailbox root.  Two paths:
        #   1. Inbox.Store.GetRootFolder()  -- works as soon as the store
        #      object is bound, even if the folder hierarchy isn't fully
        #      populated yet, and is the cleanest way to get the root.
        #   2. Inbox.Parent  -- fallback for store providers that don't
        #      expose GetRootFolder.  Can transiently be $null while a
        #      cold-launched Outlook attaches the delegated mailbox, so
        #      we retry it.
        $root = $null
        try {
            $store = $inbox.Store
            if ($store) { $root = $store.GetRootFolder() }
        } catch {
            Write-Verbose "Store.GetRootFolder() threw: $($_.Exception.Message)"
        }
        if (-not $root) {
            for ($attempt = 1; $attempt -le 6; $attempt++) {
                try { $root = $inbox.Parent } catch { }
                if ($root) { break }
                Start-Sleep -Seconds 2
            }
        }
        if (-not $root) {
            $stats.Status = "ERROR: could not resolve mailbox root for $Upn (Store.GetRootFolder() and Inbox.Parent both returned null -- is Outlook fully started and signed in?)"
            return [pscustomobject]$stats
        }
        $skip = Get-SkipFolderIds -Namespace $Namespace -Recipient $rcpt

        if ($FolderFilter) {
            $scoped = Get-FolderByPath -Root $root -Path $FolderFilter
            $candidates = Get-MailFoldersRecursive -Root $scoped -SkipIds $skip -IncludeRoot $true
        } else {
            # Mailbox root itself is a container, not a mail folder; don't
            # add it to candidates or Items.Restrict() will trip on hidden
            # store-level items.
            $candidates = Get-MailFoldersRecursive -Root $root -SkipIds $skip -IncludeRoot $false
        }

        Write-Host ("[{0}] {1} mail folder(s) to scan{2}" -f
            $Upn, $candidates.Count,
            $(if ($IsDryRun) { ' (DRY RUN)' } else { '' })) -ForegroundColor Cyan

        foreach ($folder in $candidates) {
            $r = Fix-FolderItems -Folder $folder -IsDryRun $IsDryRun
            if ($r.Found -gt 0) {
                $msg = "  '{0}': {1} draft(s)" -f $folder.FolderPath, $r.Found
                if (-not $IsDryRun) {
                    $msg += " -> fixed {0}, failed {1}" -f $r.Fixed, $r.Failed
                }
                Write-Host $msg
            }
            $stats.Folders += 1
            $stats.Drafts  += $r.Found
            $stats.Fixed   += $r.Fixed
            $stats.Failed  += $r.Failed
        }

        if ($stats.Failed -gt 0) {
            $stats.Status = ("{0} failed" -f $stats.Failed)
        }
    } catch {
        $stats.Status = "ERROR: $($_.Exception.Message)"
    }

    return [pscustomobject]$stats
}

# ---------------------------------------------------------------------------
# Entry point
# ---------------------------------------------------------------------------

function Main {
    if ($PSCmdlet.ParameterSetName -eq 'ByMapping') {
        if (-not (Test-Path -LiteralPath $MappingFile)) {
            throw "Mapping file not found: $MappingFile"
        }
        $rows = Import-Csv -LiteralPath $MappingFile
        $col = $rows[0].PSObject.Properties.Name | Where-Object { $_ -ieq 'TargetMailbox' } | Select-Object -First 1
        if (-not $col) {
            throw "Mapping file '$MappingFile' has no TargetMailbox column."
        }
        $script:Mailbox = @($rows.$col | Where-Object { $_ -and $_.Trim() } | Sort-Object -Unique)
    }

    if (-not $script:Mailbox -or $script:Mailbox.Count -eq 0) {
        throw "No mailboxes specified."
    }

    $launch = Get-OutlookApplication
    $outlook = $launch.App
    $ns = $outlook.GetNamespace('MAPI')
    try {
        # No-op if Outlook is already logged on; otherwise uses the default profile.
        $null = $ns.Logon($null, $null, $false, $false)
    } catch {
        # Already logged on -> Logon throws "namespace already logged on".  Ignore.
    }

    # Detect "we launched a fresh Outlook even though one was already
    # running" -- the classic symptom of a PowerShell/Outlook integrity-
    # level mismatch (e.g. PowerShell elevated, Outlook not).  COM's ROT
    # is per integrity level, so GetActiveObject() returns "Operation
    # unavailable" and we fall through to New-Object, which produces a
    # second, profile-less Outlook proxy.
    $admin = [Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()
    $isElevated = $admin.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)

    if ($launch.FreshLaunch) {
        $existingOutlook = @(Get-Process -Name OUTLOOK -ErrorAction SilentlyContinue)
        if ($existingOutlook.Count -gt 0) {
            Write-Warning ""
            Write-Warning "OUTLOOK.EXE is already running, but this PowerShell session"
            Write-Warning "could NOT attach to it via COM.  This almost always means"
            Write-Warning "the integrity level / user differs between the two:"
            Write-Warning ""
            if ($isElevated) {
                Write-Warning "  - You are in an *elevated* (Run as administrator) PowerShell."
                Write-Warning "  - Outlook is most likely running non-elevated."
                Write-Warning "  - Fix: close this PowerShell, open a normal (non-admin)"
                Write-Warning "    PowerShell window, and re-run the script."
            } else {
                Write-Warning "  - This PowerShell is non-elevated."
                Write-Warning "  - Outlook may be running under a different Windows user,"
                Write-Warning "    or was started 'as administrator'."
                Write-Warning "  - Fix: make sure both Outlook and PowerShell run as the"
                Write-Warning "    same user, at the same integrity level."
            }
            Write-Warning ""
            Write-Warning "Aborting before we drive a profile-less second Outlook instance."
            try { $outlook.Quit() } catch { }
            return
        }
        Write-Warning ""
        Write-Warning "Outlook was not running; the script started it for you."
        Write-Warning "If a 'Programmatic access' or 'Allow access' dialog appeared,"
        Write-Warning "click Allow / Yes -- the script is paused waiting for MAPI."
        Write-Warning ""
        Write-Warning "For best results, open Outlook BEFORE running this script,"
        Write-Warning "let it finish syncing, then re-run. Cold-launch can leave"
        Write-Warning "delegated mailboxes only partially attached."
        Write-Warning ""
        # Wait for the namespace + give the user time to dismiss the trust prompt.
        if (-not (Wait-NamespaceReady -Namespace $ns -TimeoutSeconds 60)) {
            Write-Warning "Outlook namespace did not become ready within 60s; proceeding anyway."
        }
    }

    # Sanity-check: if the namespace reports zero stores, no MAPI session
    # is attached and every shared-folder call will return null.  Bail
    # with a useful message rather than letting that surface as a confusing
    # 'Inbox.Parent was null' deep in the per-mailbox loop.
    try {
        if ($ns.Stores.Count -eq 0) {
            Write-Error ""
            Write-Error "Outlook is reporting zero MAPI stores in this session."
            Write-Error "That means the COM proxy we attached to has no profile loaded."
            Write-Error ""
            if ($isElevated) {
                Write-Error "You are running PowerShell as Administrator.  Try a normal"
                Write-Error "(non-admin) PowerShell window with Outlook already open."
            } else {
                Write-Error "Make sure Outlook is open, signed in to your admin profile,"
                Write-Error "and finished syncing, then re-run."
            }
            return
        }
    } catch {
        Write-Warning "Could not enumerate Namespace.Stores: $($_.Exception.Message)"
    }

    $results = New-Object 'System.Collections.Generic.List[pscustomobject]'
    foreach ($upn in $script:Mailbox) {
        $r = Fix-Mailbox -Namespace $ns -Upn $upn -IsDryRun:$DryRun.IsPresent -FolderFilter $Folder
        $results.Add($r)
    }

    Write-Host ""
    $results |
        Sort-Object Mailbox |
        Format-Table -AutoSize Mailbox, Folders, Drafts, Fixed, Failed, Status

    # Exit non-zero if any mailbox had failures.
    $bad = ($results | Where-Object { $_.Status -ne 'ok' }).Count
    if ($bad -gt 0) { exit 1 }
}

Main
