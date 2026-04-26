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
    2. Outlook signed in to a profile.  The account on that profile
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

# DASL filter we hand to Items.Restrict() so Outlook only iterates draft
# items in each folder rather than every message.  Equivalent SQL:
#   PR_MESSAGE_FLAGS & 0x08 != 0  AND  MessageClass LIKE 'IPM.Note%'
$script:DRAFT_RESTRICT = (
    '@SQL=' +
    '(NOT("http://schemas.microsoft.com/mapi/proptag/0x0E070003" & 8 = 0))' +
    ' AND ("urn:schemas:httpmail:messageclass" LIKE ''IPM.Note%'')'
)

# ---------------------------------------------------------------------------
# Outlook attach / launch
# ---------------------------------------------------------------------------

function Get-OutlookApplication {
    [OutputType([__ComObject])]
    param()
    try {
        $existing = [System.Runtime.InteropServices.Marshal]::GetActiveObject('Outlook.Application')
        Write-Verbose "Attached to a running Outlook session."
        return $existing
    } catch {
        Write-Host "Starting Outlook..." -ForegroundColor DarkGray
        $new = New-Object -ComObject Outlook.Application
        return $new
    }
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
    folder whose EntryID is in $SkipIds.  Returns folders bottom-up so
    that callers iterate leaves first; the order doesn't actually matter
    here, it's just a clean DFS.
    #>
    [OutputType([__ComObject[]])]
    param(
        [Parameter(Mandatory)] $Root,
        [Parameter(Mandatory)] [System.Collections.Generic.HashSet[string]]$SkipIds
    )
    $out = New-Object 'System.Collections.Generic.List[object]'
    Walk-FolderInto $Root $SkipIds $out
    return $out.ToArray()
}

function Walk-FolderInto {
    param(
        $Folder,
        [System.Collections.Generic.HashSet[string]]$SkipIds,
        [System.Collections.Generic.List[object]]$Out
    )
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

        $inbox = $Namespace.GetSharedDefaultFolder($rcpt, $script:olFolderInbox)
        $root  = $inbox.Parent
        $skip  = Get-SkipFolderIds -Namespace $Namespace -Recipient $rcpt

        if ($FolderFilter) {
            $scoped = Get-FolderByPath -Root $root -Path $FolderFilter
            $candidates = Get-MailFoldersRecursive -Root $scoped -SkipIds $skip
        } else {
            $candidates = Get-MailFoldersRecursive -Root $root -SkipIds $skip
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

    $outlook = Get-OutlookApplication
    $ns = $outlook.GetNamespace('MAPI')
    try {
        # No-op if Outlook is already logged on; otherwise uses the default profile.
        $null = $ns.Logon($null, $null, $false, $false)
    } catch {
        # Already logged on -> Logon throws "namespace already logged on".  Ignore.
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
