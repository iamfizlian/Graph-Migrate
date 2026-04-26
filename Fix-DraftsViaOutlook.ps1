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

# NOTE on filtering: we used to drive Items.Restrict() with a DASL
# filter "@SQL=<proptag>0x0E070003 & 8" to make Outlook return only
# the UNSENT-bit-set messages in each folder, but some Outlook desktop
# builds reject @SQL bitwise expressions on extended-property URIs
# ("Cannot parse condition" / "Error at @SQL=..." in Items.Restrict).
# Iterating Items in PowerShell and reading PR_MESSAGE_FLAGS via the
# PropertyAccessor is slower but works on every Outlook build we've
# tested on Exchange Online.

# System / hidden folder names we never modify, regardless of whether
# they have DefaultItemType=0.  Case-insensitive whole-name match.
# Includes the well-known "real Drafts / Outbox / Junk / Trash" set
# (covered for delegated mailboxes where GetSharedDefaultFolder may
# not return a usable EntryID), plus Teams / Yammer / RSS / sync
# scratchpads that shouldn't have UNSENT messed with.
$script:SKIP_FOLDER_NAMES = @(
    'Drafts',
    'Outbox',
    'Deleted Items', 'Trash',
    'Junk Email', 'Junk Mail', 'Junk', 'Junk E-mail',
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
    otherwise cold-launch one.

    Strategy:
      - Detect 'Outlook already running' via Get-Process, NOT via
        GetActiveObject.  GetActiveObject relies on COM's Running
        Object Table, which is unreliable across session boundaries
        (RDP, service-launched Outlook, etc.) and can fail even when
        the user and integrity level match.
      - Always attach via New-Object -ComObject Outlook.Application.
        Outlook is registered as a single-instance LocalServer32 COM
        component, so CoCreateInstance routes the call to the running
        instance instead of starting a new one.
      - Fall back to GetActiveObject only as a fast-path optimisation;
        if it works we save a bit of process startup, but we don't
        depend on it.
    #>
    [OutputType([hashtable])]
    param()
    $running = @(Get-Process -Name OUTLOOK -ErrorAction SilentlyContinue).Count -gt 0

    try {
        $existing = [System.Runtime.InteropServices.Marshal]::GetActiveObject('Outlook.Application')
        Write-Verbose "Attached to a running Outlook session via GetActiveObject."
        return @{ App = $existing; FreshLaunch = $false }
    } catch {
        # GetActiveObject failed; that's OK -- New-Object below will
        # attach to the running Outlook (single-instance COM server)
        # if there is one, or cold-launch otherwise.
    }

    if ($running) {
        Write-Verbose "OUTLOOK.EXE is running; attaching via CoCreateInstance."
    } else {
        Write-Host "Starting Outlook..." -ForegroundColor DarkGray
    }
    $new = New-Object -ComObject Outlook.Application
    return @{ App = $new; FreshLaunch = (-not $running) }
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

function Get-FolderDisplayPath {
    <#
    .SYNOPSIS  Return a human-readable folder path.  When the parent
    chain isn't fully bound yet, FolderPath comes back as raw EntryID
    hex; in that case we just use the folder's Name.
    #>
    param($Folder)
    try {
        $p = $Folder.FolderPath
        if ($p -and $p -notmatch '[0-9A-Fa-f]{30,}') {
            return $p
        }
    } catch { }
    try { return $Folder.Name } catch { return '<unnamed>' }
}

function Fix-FolderItems {
    <#
    .SYNOPSIS  Clear MSGFLAG_UNSENT on every IPM.Note item in $Folder
    whose UNSENT bit is currently set.  Returns Found/Fixed/Failed.

    We iterate Items directly rather than using Items.Restrict() with a
    DASL filter -- some Outlook desktop builds reject @SQL bitwise
    expressions on extended-property URIs ('Cannot parse condition').
    Iterating + filtering in PowerShell is slower but works everywhere.
    #>
    [OutputType([hashtable])]
    param(
        [Parameter(Mandatory)] $Folder,
        [Parameter(Mandatory)][bool]$IsDryRun
    )
    $found  = 0
    $fixed  = 0
    $failed = 0
    $display = Get-FolderDisplayPath $Folder

    # Snapshot EntryIDs of draft-flagged items first.  We use
    # GetFirst()/GetNext() rather than indexed Item() access -- per
    # Outlook docs the cursor pattern is much faster for sequential
    # iteration because it avoids re-resolving the index each call.
    $entryIds = New-Object 'System.Collections.Generic.List[string]'
    $itemCount = 0
    try {
        $itemCount = $Folder.Items.Count
    } catch {
        Write-Warning "  ${display}: Items.Count failed -- $($_.Exception.Message)"
        return @{ Found = 0; Fixed = 0; Failed = 0 }
    }
    if ($itemCount -eq 0) {
        return @{ Found = 0; Fixed = 0; Failed = 0 }
    }
    $items = $Folder.Items
    $scanProgressEvery = 100
    $scanned = 0
    try { $itm = $items.GetFirst() } catch { $itm = $null }
    while ($itm) {
        $scanned++
        try {
            $cls = $null
            try { $cls = $itm.MessageClass } catch { }
            if ($cls -and $cls.StartsWith('IPM.Note')) {
                $flags = 0
                try { $flags = [int]$itm.PropertyAccessor.GetProperty($script:PR_MESSAGE_FLAGS_TAG) } catch { }
                if (($flags -band $script:MSGFLAG_UNSENT) -ne 0) {
                    if ($itm.EntryID) { $entryIds.Add($itm.EntryID) }
                }
            }
        } catch {
            # Item couldn't be loaded -- skip.
        }
        if (($scanned % $scanProgressEvery) -eq 0) {
            Write-Host ("    scanned {0}/{1} ({2} drafts so far)" -f $scanned, $itemCount, $entryIds.Count) -ForegroundColor DarkGray
        }
        try { $itm = $items.GetNext() } catch { $itm = $null }
    }
    $found = $entryIds.Count

    if ($found -eq 0) {
        return @{ Found = 0; Fixed = 0; Failed = 0 }
    }

    if ($IsDryRun) {
        return @{ Found = $found; Fixed = 0; Failed = 0 }
    }

    # Re-fetch each item by EntryID via the store and clear the bit.
    $store = $null
    try { $store = $Folder.Store } catch { }
    $fixProgressEvery = 50
    $fixedSoFar = 0
    foreach ($eid in $entryIds) {
        $fixedSoFar++
        try {
            $itm = $null
            if ($store) {
                try { $itm = $store.GetItemFromID($eid) } catch { }
            }
            if (-not $itm) {
                # Fallback: ask the namespace directly.
                try { $itm = $Folder.Application.Session.GetItemFromID($eid) } catch { }
            }
            if (-not $itm) { $failed++; continue }
            if ($itm.MessageClass -and -not $itm.MessageClass.StartsWith('IPM.Note')) {
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
        if (($fixedSoFar % $fixProgressEvery) -eq 0) {
            Write-Host ("    fixed {0}/{1} ({2} failed)" -f $fixedSoFar, $found, $failed) -ForegroundColor DarkGray
        }
    }

    return @{ Found = $found; Fixed = $fixed; Failed = $failed }
}

# ---------------------------------------------------------------------------
# MAPI binding helpers
# ---------------------------------------------------------------------------

function Force-FolderBind {
    <#
    .SYNOPSIS  Force MAPI to actually bind a Folder COM proxy by reading
    a property that requires server-side resolution.  Returns $true on
    success.  Cold-launched Outlook can hand back unbound proxies whose
    Items/Store/Parent all read $null; this routine drives a few
    retries with backoff to give MAPI time to attach.
    #>
    param($Folder, [int]$MaxAttempts = 6)
    for ($attempt = 1; $attempt -le $MaxAttempts; $attempt++) {
        try {
            $null = $Folder.Items.Count
            return $true
        } catch {
            Write-Verbose ("Force-FolderBind attempt {0}: {1}" -f $attempt, $_.Exception.Message)
            Start-Sleep -Seconds 2
        }
    }
    return $false
}

function Get-WellKnownFolderBound {
    <#
    .SYNOPSIS  GetSharedDefaultFolder() with retries + a Force-FolderBind
    once we have a proxy back, so the caller gets a folder that actually
    responds to .Items / .Folders.  Returns $null on persistent failure.
    #>
    param($Namespace, $Recipient, [int]$FolderType, [int]$MaxAttempts = 6)
    for ($attempt = 1; $attempt -le $MaxAttempts; $attempt++) {
        $f = $null
        try {
            $f = $Namespace.GetSharedDefaultFolder($Recipient, $FolderType)
        } catch {
            Write-Verbose ("GetSharedDefaultFolder({0}) attempt {1}: {2}" -f $FolderType, $attempt, $_.Exception.Message)
        }
        if ($f) {
            if (Force-FolderBind $f -MaxAttempts 3) { return $f }
        }
        Start-Sleep -Seconds 2
    }
    return $null
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

        # Verify what the recipient actually resolved to.  If FullAccess
        # is missing and the address falls back to the local user, we'd
        # be about to walk the WRONG mailbox -- log enough info to spot
        # that.
        $resolvedSmtp = $null
        try {
            $exu = $rcpt.AddressEntry.GetExchangeUser()
            if ($exu) { $resolvedSmtp = $exu.PrimarySmtpAddress }
        } catch { }
        if ($resolvedSmtp) {
            Write-Host ("[{0}] recipient resolved to {1}" -f $Upn, $resolvedSmtp) -ForegroundColor DarkGray
            if ($resolvedSmtp -inotlike $Upn) {
                Write-Warning ("[{0}] resolved SMTP {1} != requested UPN -- check FullAccess permissions." -f $Upn, $resolvedSmtp)
            }
        }

        # Fetch Inbox with retries + force MAPI to bind it.  Cold-launched
        # Outlook can otherwise hand back an unbound Inbox proxy whose
        # Items / Store / Parent all read as $null.
        $inbox = Get-WellKnownFolderBound -Namespace $Namespace -Recipient $rcpt -FolderType $script:olFolderInbox
        if (-not $inbox) {
            $stats.Status = "ERROR: GetSharedDefaultFolder returned null for $Upn (FullAccess granted with -AutoMapping `$false?  Or is Outlook still starting?)"
            return [pscustomobject]$stats
        }

        # Diagnostic: confirm we're looking at the right mailbox's Inbox.
        try {
            $inboxItems = $inbox.Items.Count
            $inboxStore = $null
            try { $inboxStore = $inbox.Store.DisplayName } catch { }
            Write-Host ("[{0}] Inbox bound: name='{1}', items={2}, store='{3}'" -f
                $Upn, $inbox.Name, $inboxItems, ($inboxStore -as [string])) -ForegroundColor DarkGray
        } catch {
            Write-Verbose ("Could not read Inbox diagnostics: {0}" -f $_.Exception.Message)
        }

        # Resolve the mailbox root.  Three paths, tried in order; each is
        # robust against a different MAPI binding hiccup.
        #   1. Inbox.Store.GetRootFolder()  -- works as soon as the store
        #      COM object is bound.
        #   2. Inbox.Parent  -- needs the folder hierarchy to be loaded;
        #      transiently $null on cold-launched Outlook.
        #   3. None of the above -- we still walk Inbox + Sent Items
        #      directly via GetSharedDefaultFolder, which is enough to
        #      cover the imported-PST scenario for this project.
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
        $skip = Get-SkipFolderIds -Namespace $Namespace -Recipient $rcpt

        if ($FolderFilter) {
            if (-not $root) {
                $stats.Status = "ERROR: could not resolve mailbox root for $Upn; -Folder filter requires the root.  Try restarting Outlook first."
                return [pscustomobject]$stats
            }
            $scoped = Get-FolderByPath -Root $root -Path $FolderFilter
            $candidates = Get-MailFoldersRecursive -Root $scoped -SkipIds $skip -IncludeRoot $true
        } elseif ($root) {
            # Best path: walk every IPF.Note folder under the mailbox root.
            # The mailbox root itself is a container, not a mail folder; don't
            # add it to candidates or Items.Restrict() will trip on hidden
            # store-level items.
            $candidates = Get-MailFoldersRecursive -Root $root -SkipIds $skip -IncludeRoot $false
        } else {
            # Fallback: cold-launched Outlook hasn't fully bound the
            # delegated store, so Inbox.Parent and Inbox.Store are both
            # null.  Walk the folders we can reach directly via
            # GetSharedDefaultFolder.  Misses user-created top-level
            # folders like 'Archive' but covers Inbox + Sent Items, which
            # is where imported-PST messages always end up.
            Write-Warning ("[{0}] mailbox root unresolved; falling back to Inbox + Sent Items only." -f $Upn)
            Write-Warning ("[{0}] To cover other top-level folders, restart Outlook in this same login session and re-run." -f $Upn)
            $candidates = New-Object 'System.Collections.Generic.List[object]'
            foreach ($wellKnown in @($script:olFolderInbox, $script:olFolderSentMail)) {
                $f = Get-WellKnownFolderBound -Namespace $Namespace -Recipient $rcpt -FolderType $wellKnown
                if ($f) { Walk-FolderInto $f $skip $candidates }
            }
            $candidates = $candidates.ToArray()
        }

        Write-Host ("[{0}] {1} mail folder(s) to scan{2}" -f
            $Upn, $candidates.Count,
            $(if ($IsDryRun) { ' (DRY RUN)' } else { '' })) -ForegroundColor Cyan

        foreach ($folder in $candidates) {
            $display = Get-FolderDisplayPath $folder
            $itemCount = -1
            try { $itemCount = $folder.Items.Count } catch { }
            Write-Host ("  scanning '{0}' ({1} item(s))..." -f $display, $itemCount) -ForegroundColor DarkGray
            $r = Fix-FolderItems -Folder $folder -IsDryRun $IsDryRun
            $msg = "  '{0}': {1} draft(s)" -f $display, $r.Found
            if (-not $IsDryRun -and $r.Found -gt 0) {
                $msg += " -> fixed {0}, failed {1}" -f $r.Fixed, $r.Failed
            }
            Write-Host $msg
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

    if ($launch.FreshLaunch) {
        Write-Warning ""
        Write-Warning "Outlook was not already running; the script started it."
        Write-Warning "If a 'Programmatic access' / 'Allow access' dialog appears,"
        Write-Warning "click Allow / Yes -- the script is waiting for MAPI."
        Write-Warning ""
        # Wait for the namespace + give the user time to dismiss the trust prompt.
        if (-not (Wait-NamespaceReady -Namespace $ns -TimeoutSeconds 90)) {
            Write-Warning "Outlook namespace did not become ready within 90s; proceeding anyway."
        }
    }

    # Sanity-check: if the namespace reports zero stores, no MAPI session
    # is attached and every shared-folder call will return null.  Bail
    # with a useful message rather than letting that surface as a confusing
    # 'Inbox.Parent was null' deep in the per-mailbox loop.
    $storeCount = -1
    try { $storeCount = $ns.Stores.Count } catch { }
    if ($storeCount -le 0) {
        Write-Error ""
        Write-Error "Outlook is reporting $storeCount MAPI stores in this session."
        Write-Error "That means no Outlook profile is attached to the COM session"
        Write-Error "we're driving.  Common causes:"
        Write-Error ""
        Write-Error "  - Outlook is open under a different Windows user (e.g. a"
        Write-Error "    different RDP session or service account)."
        Write-Error "  - Outlook hasn't finished launching/signing in yet."
        Write-Error "  - The Outlook profile is corrupted or asks for credentials."
        Write-Error ""
        Write-Error "Open Outlook in this same login session, sign in to the admin"
        Write-Error "profile, wait for the 'Connected' status, then re-run."
        return
    }
    Write-Verbose ("Namespace ready: {0} store(s) attached." -f $storeCount)

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
