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
    [string]$Folder,

    # Outlook's Object Model silently no-ops PropertyAccessor.SetProperty()
    # on PR_MESSAGE_FLAGS (it's documented as a "computed" property).  The
    # Redemption COM library (https://www.dimastr.com/redemption/) bypasses
    # this restriction by writing through Extended MAPI directly.  By
    # default we try Redemption if it's registered on the box and only fall
    # back to the OOM path if it's not.  Pass -NoRedemption to force the
    # OOM path (useful only as an A/B control test -- it will not actually
    # clear the bit).
    [Parameter()]
    [switch]$NoRedemption,

    # Cap how many items we attempt per folder.  Used for debug-style
    # runs where we want to inspect 1-3 items with full diagnostics
    # rather than process hundreds.  0 / unset = no cap.
    [Parameter()]
    [int]$MaxItems = 0
)

# ---------------------------------------------------------------------------
# Constants
# ---------------------------------------------------------------------------

# MAPI property tag, expressed as the PropertyAccessor schema URI Outlook
# wants. 0x0E07 is PR_MESSAGE_FLAGS; 0003 is the PT_LONG type code.
$script:PR_MESSAGE_FLAGS_TAG       = "http://schemas.microsoft.com/mapi/proptag/0x0E070003"
$script:PR_INTERNET_MESSAGE_ID_TAG = "http://schemas.microsoft.com/mapi/proptag/0x1035001E"
$script:MSGFLAG_UNSENT             = 0x08

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
# Redemption (Extended MAPI) helpers
# ---------------------------------------------------------------------------

function Get-RedemptionSession {
    <#
    .SYNOPSIS  Build an RDOSession that shares the running Outlook
    MAPI session.  Returns $null if Redemption isn't registered.

    Why we need this: Outlook's Object Model silently rejects writes
    to PR_MESSAGE_FLAGS via PropertyAccessor.SetProperty -- the call
    returns Success but the bit doesn't move.  Redemption's RDOMail
    talks to Extended MAPI directly (the same layer MFCMAPI uses),
    so it can write computed-but-not-actually-computed properties
    that OOM blocks.  We share Outlook's MAPIOBJECT so Redemption
    inherits the same authenticated session and FullAccess delegation
    permissions -- no separate logon, no second profile.
    #>
    [OutputType([object])]
    param(
        [Parameter(Mandatory)] $Namespace
    )
    $rdo = $null
    try {
        $rdo = New-Object -ComObject Redemption.RDOSession
    } catch {
        Write-Verbose ("Redemption.RDOSession not available: {0}" -f $_.Exception.Message)
        return $null
    }
    try {
        # Hand Redemption Outlook's existing MAPI session.  Without this
        # RDOSession would attempt its own profile logon, which fails on
        # locked-down servers and would not see Outlook's delegate
        # FullAccess grants either way.
        $rdo.MAPIOBJECT = $Namespace.MAPIOBJECT
    } catch {
        Write-Warning ("Redemption was loaded but could not attach to Outlook's MAPI session: {0}" -f $_.Exception.Message)
        Write-Warning '  Falling back to the OOM path -- expect 0 fixes / many "stuck" items.'
        return $null
    }
    return $rdo
}

# Emits a verbose snapshot of PR_MESSAGE_FLAGS at every interesting
# step on the first item we process, then sets $script:_diagShown so
# subsequent items run quietly.  Diagnostic output looks like:
#
#   DIAG src.Fields[0x0E070003]   = 0x0001 (UNSENT=False, READ=True, ...)
#   DIAG dst after Items.Add      = 0x0009 (UNSENT=True, ...)
#   DIAG dst after CopyTo         = 0x0001 (UNSENT=False, ...)
#   DIAG dst after Sent=true      = 0x0001
#   DIAG dst after Fields write   = 0x0001
#   DIAG dst after Save()         = 0x0009 (UNSENT=True)   <- server overrode!
#
# That last line is the smoking gun if Exchange Online is force-setting
# UNSENT during the new-message save.  If the line shows UNSENT=False
# then the COPY path is working and the verify pass is wrong.
$script:_diagShown = $false

# Redemption registers RDOMail.Fields as a *parameterized property*
# at the IDispatch level: callable as `mail.Fields(propTag)` rather
# than as a property that returns a collection.  PowerShell exposes
# such properties as System.Management.Automation.PSParameterizedProperty.
# That means:
#   - `$mail.Fields.Item($tag)` fails -- PSParameterizedProperty has
#                                        no Item method.
#   - `$mail.Fields.GetType().InvokeMember('Item', ...)` also fails
#                                        for the same reason.
# The two ways that *do* work are:
#   1. `$mail.Fields.Invoke($tag)` -- calls the get accessor through
#                                     the PSParameterizedProperty
#                                     wrapper.
#   2. `$mail.GetType().InvokeMember('Fields', GetProperty, ..., @($tag))`
#                                  -- treats Fields as the indexed
#                                     property of the parent RDOMail
#                                     and dispatches through its
#                                     IDispatch.  This is the only
#                                     option for the SET path
#                                     (PSParameterizedProperty has
#                                     no settable Invoke).
#
# Both helpers accept either an int prop tag (0x0E070003) OR a DASL
# string ("http://schemas.microsoft.com/mapi/proptag/0x0E070003")
# -- Redemption supports both index types.
function Get-RdoMailField {
    [OutputType([object])]
    param(
        [Parameter(Mandatory)]$Mail,
        [Parameter(Mandatory)]$Key
    )
    return $Mail.Fields.Invoke([object]$Key)
}

function Set-RdoMailField {
    param(
        [Parameter(Mandatory)]$Mail,
        [Parameter(Mandatory)]$Key,
        [Parameter(Mandatory)]$Value
    )
    [void]$Mail.GetType().InvokeMember(
        'Fields',
        [System.Reflection.BindingFlags]::SetProperty,
        $null,
        $Mail,
        @([object]$Key, [object]$Value)
    )
}

function Format-MessageFlags {
    param([int]$Flags)
    $bits = New-Object 'System.Collections.Generic.List[string]'
    if ($Flags -band 0x01) { [void]$bits.Add('READ') }
    if ($Flags -band 0x02) { [void]$bits.Add('UNMOD') }
    if ($Flags -band 0x04) { [void]$bits.Add('SUBMIT') }
    if ($Flags -band 0x08) { [void]$bits.Add('UNSENT') }
    if ($Flags -band 0x10) { [void]$bits.Add('HASATTACH') }
    if ($Flags -band 0x20) { [void]$bits.Add('FROMME') }
    if ($Flags -band 0x40) { [void]$bits.Add('ASSOC') }
    if ($Flags -band 0x80) { [void]$bits.Add('RESEND') }
    if ($bits.Count -eq 0) { 'none' } else { $bits -join '|' }
}

function Clear-UnsentBitViaRedemptionCopy {
    <#
    .SYNOPSIS  Clear MSGFLAG_UNSENT by recreating the message.

    Why a copy and not an in-place update: per MAPI spec
    (MS-OXCMSG, IMessage::SetProps),
    PR_MESSAGE_FLAGS becomes effectively immutable after the first
    IMessage::SaveChanges -- the cloud store silently rejects further
    writes to that property regardless of which client (Graph, EWS,
    OOM, or even Extended MAPI via Redemption) issues them.  So
    even Redemption's $rdoMail.Sent = $true ; Save() is a no-op on
    an existing message.

    BUT -- on a *brand new* message, the client owns PR_MESSAGE_FLAGS
    until the first save.  So we:
        1. Open the source message via Redemption.
        2. Add a fresh blank message to the same folder
           (RDOFolder.Items.Add returns an unsaved IMessage).
        3. RDOMail.CopyTo(blank) copies every property except
           EntryID/StoreEntryID/RecordKey/InstanceKey/SearchKey/
           ParentEntryID/etc, which Redemption excludes by default.
        4. Set blank.Sent = $true *before* Save() so the first save
           writes PR_MESSAGE_FLAGS without the UNSENT bit.
        5. Save the new message -- it lands without the Draft badge.
        6. Delete the original.

    Caveats:
        * The new item gets a new EntryID and a new InternetMessageId
          (Redemption preserves the source MIME ID by default; we
          could explicitly preserve it if needed).
        * Conversation threading is preserved because
          PR_CONVERSATION_INDEX / TOPIC are copied.
        * DateTimeReceived, DateTimeSent, Sender* are all copied.

    Returns one of:
        'fixed'   - new message created with UNSENT clear, original deleted
        'noop'    - new message landed but its UNSENT bit is also set
                    (cloud store applied UNSENT during save anyway --
                    extremely unusual, would mean the server is
                    auto-flagging *all* PSTs-imported-style writes)
        'failed'  - exception during any step (source NOT deleted)

    Sets the [ref] $NewEntryId to the new message's EntryID on
    success so the caller can record the mapping if it wants.
    #>
    [OutputType([string])]
    param(
        [Parameter(Mandatory)] $RdoSession,
        [Parameter(Mandatory)][string]$EntryId,
        [Parameter(Mandatory)][string]$StoreId,
        [Parameter()] [ref]$NewEntryId
    )

    $tagInt = 0x0E070003
    $diag   = -not $script:_diagShown

    $src    = $null
    $folder = $null
    $dst    = $null
    try {
        $src = $RdoSession.GetMessageFromID($EntryId, $StoreId)
        if (-not $src) { return 'failed' }

        # Containing folder for the new copy.  We deliberately keep the
        # new item in the SAME folder as the original so the user-visible
        # location does not move.
        $folder = $src.Parent
        if (-not $folder) { return 'failed' }

        $msgClass = 'IPM.Note'
        try { if ($src.MessageClass) { $msgClass = $src.MessageClass } } catch { }

        if ($diag) {
            Write-Host '' -ForegroundColor Yellow
            Write-Host '==== DIAG: first copy-and-replace; logging every flag write ====' -ForegroundColor Yellow
            Write-Host ("DIAG src.MessageClass = '{0}'" -f $msgClass) -ForegroundColor Yellow
            try {
                $f0 = [int](Get-RdoMailField -Mail $src -Key $tagInt)
                Write-Host ("DIAG src.Fields[0x0E070003]   = 0x{0:X4}  ({1})" -f $f0, (Format-MessageFlags $f0)) -ForegroundColor Yellow
            } catch {
                Write-Host "DIAG src.Fields read failed: $($_.Exception.Message)" -ForegroundColor Red
            }
        }

        # Items.Add returns an RDOMail that has NOT been saved yet.  Until
        # the first Save(), we own all of its properties including
        # PR_MESSAGE_FLAGS.
        $dst = $folder.Items.Add($msgClass)
        if (-not $dst) { return 'failed' }
        if ($diag) {
            try {
                $f1 = [int](Get-RdoMailField -Mail $dst -Key $tagInt)
                Write-Host ("DIAG dst after Items.Add      = 0x{0:X4}  ({1})" -f $f1, (Format-MessageFlags $f1)) -ForegroundColor Yellow
            } catch {
                Write-Host "DIAG dst.Fields read after Items.Add failed: $($_.Exception.Message)" -ForegroundColor Red
            }
        }

        # CopyTo against an IMessage destination copies properties without
        # saving.  Redemption excludes the identity / instance properties
        # automatically, so we don't have to enumerate an exclusion list.
        $src.CopyTo($dst)
        if ($diag) {
            try {
                $f2 = [int](Get-RdoMailField -Mail $dst -Key $tagInt)
                Write-Host ("DIAG dst after CopyTo         = 0x{0:X4}  ({1})" -f $f2, (Format-MessageFlags $f2)) -ForegroundColor Yellow
            } catch { }
        }

        # Belt + suspenders: try BOTH ways to clear UNSENT before the
        # first save.
        #
        #   (a) RDOMail.Sent = $true           -- Redemption's high-level
        #                                         wrapper for clearing
        #                                         MSGFLAG_UNSENT.
        #   (b) Fields[0x0E070003] = explicit  -- direct Extended MAPI
        #                                         property write, in
        #                                         case (a) is a no-op
        #                                         on this Redemption
        #                                         build.
        try { $dst.Sent = $true } catch {
            if ($diag) { Write-Host "DIAG dst.Sent = `$true threw: $($_.Exception.Message)" -ForegroundColor Red }
        }
        if ($diag) {
            try {
                $f3 = [int](Get-RdoMailField -Mail $dst -Key $tagInt)
                Write-Host ("DIAG dst after Sent=`$true     = 0x{0:X4}  ({1})" -f $f3, (Format-MessageFlags $f3)) -ForegroundColor Yellow
            } catch { }
        }

        try {
            $current = [int](Get-RdoMailField -Mail $dst -Key $tagInt)
            $newVal  = $current -band (-bnot $script:MSGFLAG_UNSENT)
            Set-RdoMailField -Mail $dst -Key $tagInt -Value $newVal
        } catch {
            if ($diag) { Write-Host "DIAG dst.Fields write threw: $($_.Exception.Message)" -ForegroundColor Red }
        }
        if ($diag) {
            try {
                $f4 = [int](Get-RdoMailField -Mail $dst -Key $tagInt)
                Write-Host ("DIAG dst after Fields write   = 0x{0:X4}  ({1})" -f $f4, (Format-MessageFlags $f4)) -ForegroundColor Yellow
            } catch { }
        }

        $dst.Save()
        if ($diag) {
            try {
                $f5 = [int](Get-RdoMailField -Mail $dst -Key $tagInt)
                Write-Host ("DIAG dst after Save() inproc  = 0x{0:X4}  ({1})" -f $f5, (Format-MessageFlags $f5)) -ForegroundColor Yellow
            } catch { }
            # Re-open via EntryID so we read the SERVER's state, not the
            # in-process RDOMail's cached property bag.  This is the
            # ground truth.
            try {
                $newEid = $dst.EntryID
                $reopen = $RdoSession.GetMessageFromID($newEid, $StoreId)
                if ($reopen) {
                    $f6 = [int](Get-RdoMailField -Mail $reopen -Key $tagInt)
                    $color = if ($f6 -band $script:MSGFLAG_UNSENT) { 'Red' } else { 'Green' }
                    Write-Host ("DIAG dst RE-FETCH from store = 0x{0:X4}  ({1})  <-- ground truth" -f $f6, (Format-MessageFlags $f6)) -ForegroundColor $color
                } else {
                    Write-Host 'DIAG could not re-fetch new message via GetMessageFromID' -ForegroundColor Red
                }
            } catch {
                Write-Host "DIAG re-fetch failed: $($_.Exception.Message)" -ForegroundColor Red
            }
            Write-Host '==== /DIAG ====' -ForegroundColor Yellow
            Write-Host '' -ForegroundColor Yellow
            $script:_diagShown = $true
        }
    } catch {
        Write-Verbose ("RDO copy-replace failed for {0}: {1}" -f $EntryId, $_.Exception.Message)
        return 'failed'
    }

    # Verify by RE-FETCHING the new message via its EntryID (the
    # in-process $dst's property bag can lag behind the server).
    $verified = $false
    $newEid   = $null
    try {
        $newEid = $dst.EntryID
        $reopen = $RdoSession.GetMessageFromID($newEid, $StoreId)
        if ($reopen) {
            $after = [int](Get-RdoMailField -Mail $reopen -Key $tagInt)
            if (($after -band $script:MSGFLAG_UNSENT) -eq 0) { $verified = $true }
        }
    } catch {
        Write-Verbose ("RDO verify of new copy failed for {0}: {1}" -f $EntryId, $_.Exception.Message)
    }

    if (-not $verified) {
        # Server set UNSENT during Save() despite our pre-save clear.
        # Don't delete the source -- caller will count this as 'noop' and
        # we'll still have the original to investigate / re-import.
        return 'noop'
    }

    if ($PSBoundParameters.ContainsKey('NewEntryId') -and $newEid) {
        try { $NewEntryId.Value = $newEid } catch { }
    }

    # Source recreation succeeded; delete the original so the user sees
    # exactly one copy in the folder.  If Delete throws, we still return
    # 'fixed' but the user will see a duplicate -- worth surfacing.
    try {
        $src.Delete()
    } catch {
        Write-Warning ("  copy succeeded but original delete failed (duplicate left in folder): {0}" -f $_.Exception.Message)
    }

    return 'fixed'
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
        [Parameter(Mandatory)][bool]$IsDryRun,
        [Parameter()] $RdoSession   # may be $null -> falls back to OOM
    )
    $found  = 0
    $fixed  = 0
    $failed = 0
    $display = Get-FolderDisplayPath $Folder

    # Single pass over the folder to:
    #   1. Snapshot EntryIDs of draft-flagged items (the work list).
    #   2. Build a clean-twin index { lower(IMID) -> EntryID } for items
    #      that already have UNSENT clear -- these are likely
    #      already-clean ghost copies left by previous failed-verify
    #      runs, and we want to short-circuit the copy-and-replace path
    #      by deleting the dirty original instead of creating yet
    #      another duplicate.
    #
    # We use GetFirst()/GetNext() rather than indexed Item() access --
    # per Outlook docs the cursor pattern is much faster for sequential
    # iteration because it avoids re-resolving the index each call.
    $entryIds     = New-Object 'System.Collections.Generic.List[string]'
    $cleanTwinMap = @{}
    $itemCount = 0
    try {
        $itemCount = $Folder.Items.Count
    } catch {
        Write-Warning "  ${display}: Items.Count failed -- $($_.Exception.Message)"
        return @{ Found = 0; Fixed = 0; Failed = 0; SilentNoop = 0; ResolvedByTwin = 0 }
    }
    if ($itemCount -eq 0) {
        return @{ Found = 0; Fixed = 0; Failed = 0; SilentNoop = 0; ResolvedByTwin = 0 }
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
                } else {
                    # Already-clean candidate.  Index it by IMID so a
                    # dirty draft with the same IMID can defer to it
                    # instead of creating a third instance.
                    $imid = $null
                    try { $imid = [string]$itm.PropertyAccessor.GetProperty($script:PR_INTERNET_MESSAGE_ID_TAG) } catch { }
                    if ($imid) {
                        $key = $imid.ToLowerInvariant()
                        if (-not $cleanTwinMap.ContainsKey($key)) {
                            $cleanTwinMap[$key] = $itm.EntryID
                        }
                    }
                }
            }
        } catch {
            # Item couldn't be loaded -- skip.
        }
        if (($scanned % $scanProgressEvery) -eq 0) {
            Write-Host ("    scanned {0}/{1} ({2} drafts so far, {3} clean twin(s) indexed)" `
                -f $scanned, $itemCount, $entryIds.Count, $cleanTwinMap.Count) -ForegroundColor DarkGray
        }
        try { $itm = $items.GetNext() } catch { $itm = $null }
    }
    $found = $entryIds.Count

    if ($found -eq 0) {
        return @{ Found = 0; Fixed = 0; Failed = 0; SilentNoop = 0; ResolvedByTwin = 0 }
    }

    if ($IsDryRun) {
        return @{ Found = $found; Fixed = 0; Failed = 0; SilentNoop = 0; ResolvedByTwin = 0 }
    }

    # Re-fetch each item by EntryID and clear the bit.  Two write paths:
    #
    #   * Redemption (preferred when $RdoSession is non-null).  Goes
    #     through Extended MAPI (IMessage::SetProps) directly, which
    #     bypasses Outlook's OOM "computed property" block on
    #     PR_MESSAGE_FLAGS.  The Sent boolean is Redemption's
    #     high-level wrapper around clearing MSGFLAG_UNSENT.
    #
    #   * OOM PropertyAccessor (fallback when Redemption isn't
    #     available).  Documented to silently no-op on
    #     PR_MESSAGE_FLAGS -- left in only as a control test, will
    #     produce mostly $silentNoop counts.
    #
    # In either case we re-read the property AFTER the save and only
    # count an item as Fixed when the MSGFLAG_UNSENT bit actually
    # cleared on the server.
    $store = $null
    $storeId = $null
    try { $store = $Folder.Store }   catch { }
    try { $storeId = $Folder.StoreID } catch { }
    $useRdo = ($null -ne $RdoSession -and $null -ne $storeId)

    $fixProgressEvery = 50
    $fixedSoFar = 0
    $silentNoop = 0
    $resolvedByTwin = 0
    $cap = $script:MaxItemsCap
    foreach ($eid in $entryIds) {
        if ($cap -gt 0 -and $fixedSoFar -ge $cap) {
            Write-Host ("    -MaxItems cap ({0}) reached; stopping early." -f $cap) -ForegroundColor Yellow
            break
        }
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

            # Twin-resolution: if we already saw a clean message in this
            # folder with the same InternetMessageId, the dirty draft is
            # a duplicate from a previous run's botched verify.  Just
            # delete it -- no need to create a third copy.
            $srcImid = $null
            try { $srcImid = [string]$pa.GetProperty($script:PR_INTERNET_MESSAGE_ID_TAG) } catch { }
            if ($srcImid -and $cleanTwinMap.ContainsKey($srcImid.ToLowerInvariant())) {
                try {
                    $itm.Delete()
                    $fixed++
                    $resolvedByTwin++
                } catch {
                    $failed++
                    Write-Verbose "  twin-delete failed for ${eid}: $($_.Exception.Message)"
                }
                if (($fixedSoFar % $fixProgressEvery) -eq 0) {
                    Write-Host ("    processed {0}/{1} (verified-fixed {2}, twin-resolved {3}, silent-noop {4}, errored {5})" `
                        -f $fixedSoFar, $found, $fixed, $resolvedByTwin, $silentNoop, $failed) -ForegroundColor DarkGray
                }
                continue
            }

            if ($useRdo) {
                # Extended-MAPI copy-and-replace path via Redemption.
                # Creates a new message with UNSENT clear, deletes the
                # original.  The OOM $itm we have above will become a
                # zombie handle after this call -- which is fine because
                # we don't reference it again in this iteration.
                $newEid = $null
                $outcome = Clear-UnsentBitViaRedemptionCopy `
                    -RdoSession $RdoSession `
                    -EntryId    $eid `
                    -StoreId    $storeId `
                    -NewEntryId ([ref]$newEid)
                switch ($outcome) {
                    'fixed'  {
                        $fixed++
                        # Register the new copy as a twin so subsequent
                        # siblings in the same run pick it up via the
                        # cheap delete path instead of a second
                        # copy-and-replace.
                        if ($srcImid -and $newEid) {
                            $cleanTwinMap[$srcImid.ToLowerInvariant()] = $newEid
                        }
                    }
                    'noop'   { $silentNoop++ }
                    default  { $failed++ }
                }
            } else {
                # OOM PropertyAccessor path -- known broken, kept as control.
                $newVal = $current -band (-bnot $script:MSGFLAG_UNSENT)
                $pa.SetProperty($script:PR_MESSAGE_FLAGS_TAG, $newVal)
                $itm.Save()

                $verifyItm = $null
                try {
                    if ($store) { $verifyItm = $store.GetItemFromID($eid) }
                    if (-not $verifyItm) {
                        $verifyItm = $Folder.Application.Session.GetItemFromID($eid)
                    }
                } catch { $verifyItm = $null }
                if ($verifyItm) {
                    $after = $null
                    try { $after = [int]$verifyItm.PropertyAccessor.GetProperty($script:PR_MESSAGE_FLAGS_TAG) } catch { }
                    if ($null -ne $after -and ($after -band $script:MSGFLAG_UNSENT) -eq 0) {
                        $fixed++
                    } else {
                        $silentNoop++
                    }
                } else {
                    $silentNoop++
                }
            }
        } catch {
            $failed++
            Write-Verbose "  item $eid failed: $($_.Exception.Message)"
        }
        if (($fixedSoFar % $fixProgressEvery) -eq 0) {
            Write-Host ("    processed {0}/{1} (verified-fixed {2}, twin-resolved {3}, silent-noop {4}, errored {5})" `
                -f $fixedSoFar, $found, $fixed, $resolvedByTwin, $silentNoop, $failed) -ForegroundColor DarkGray
        }
    }

    if ($resolvedByTwin -gt 0) {
        Write-Host ("    twin-resolved {0}/{1} item(s) by deleting dirty originals against pre-existing clean copies." `
            -f $resolvedByTwin, $found) -ForegroundColor DarkGreen
    }

    if ($silentNoop -gt 0) {
        Write-Warning ("  {0}: {1}/{2} write(s) reported success but the UNSENT bit did NOT clear on the server." -f $display, $silentNoop, $found)
        if ($useRdo) {
            Write-Warning '  The Redemption COPY-and-REPLACE path landed a new message but'
            Write-Warning '  the cloud store applied UNSENT during Save() on those particular'
            Write-Warning '  items.  Re-running the script usually picks them up on the next'
            Write-Warning '  pass; if they remain stuck after 2-3 runs, the original message'
            Write-Warning '  shape may be unusual (e.g. embedded forwarded item) -- inspect'
            Write-Warning '  by EntryID in MFCMAPI before treating as data-loss-risk.'
        } else {
            Write-Warning '  This is the documented Outlook OOM block on PR_MESSAGE_FLAGS (computed property).'
            Write-Warning '  Install Outlook Redemption (https://www.dimastr.com/redemption/) and re-run --'
            Write-Warning '  the script will detect it and use the Extended MAPI copy-and-replace path automatically.'
        }
    }

    return @{
        Found          = $found
        Fixed          = $fixed
        Failed         = $failed
        SilentNoop     = $silentNoop
        ResolvedByTwin = $resolvedByTwin
    }
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
        [Parameter()][string]$FolderFilter,
        [Parameter()] $RdoSession   # may be $null -> OOM fallback
    )

    $stats = [ordered]@{
        Mailbox  = $Upn
        Folders  = 0
        Drafts   = 0
        Fixed    = 0
        Stuck    = 0
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
            $r = Fix-FolderItems -Folder $folder -IsDryRun $IsDryRun -RdoSession $RdoSession
            $msg = "  '{0}': {1} draft(s)" -f $display, $r.Found
            if (-not $IsDryRun -and $r.Found -gt 0) {
                $stuckPart = ''
                if ($r.SilentNoop -gt 0) {
                    $stuckPart = ", stuck {0}" -f $r.SilentNoop
                }
                $twinPart = ''
                if ($r.ResolvedByTwin -gt 0) {
                    $twinPart = ", twin-resolved {0}" -f $r.ResolvedByTwin
                }
                $msg += " -> fixed {0}{1}{2}, failed {3}" -f $r.Fixed, $twinPart, $stuckPart, $r.Failed
            }
            Write-Host $msg
            $stats.Folders += 1
            $stats.Drafts  += $r.Found
            $stats.Fixed   += $r.Fixed
            $stats.Stuck   += [int]$r.SilentNoop
            $stats.Failed  += $r.Failed
        }

        if ($stats.Failed -gt 0) {
            $stats.Status = ("{0} failed" -f $stats.Failed)
        } elseif ($stats.Stuck -gt 0) {
            $stats.Status = ("{0} stuck (OOM no-op)" -f $stats.Stuck)
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

    # Stand up a Redemption session that shares Outlook's MAPI session.
    # If Redemption isn't installed (or -NoRedemption is set) we fall
    # back to the OOM PropertyAccessor path, which is documented to
    # silently no-op on PR_MESSAGE_FLAGS -- Stuck count will reflect
    # that and the per-folder warning will tell the user to install it.
    $rdo = $null
    if (-not $NoRedemption) {
        $rdo = Get-RedemptionSession -Namespace $ns
        if ($rdo) {
            Write-Host "Redemption (Extended MAPI) session attached -- using it for PR_MESSAGE_FLAGS writes." -ForegroundColor Green
        } else {
            Write-Warning ""
            Write-Warning "Outlook Redemption is not registered on this machine."
            Write-Warning "PR_MESSAGE_FLAGS writes will go through the OOM PropertyAccessor"
            Write-Warning "and Microsoft documents that path as a silent no-op for this"
            Write-Warning "property -- expect Stuck = Drafts and Fixed = 0."
            Write-Warning ""
            Write-Warning "  Install Redemption (~30 seconds, free for in-house use):"
            Write-Warning "    1. Download from https://www.dimastr.com/redemption/"
            Write-Warning "    2. Unzip Redemption64.dll into C:\Program Files\Redemption\"
            Write-Warning "    3. From an elevated cmd:  regsvr32 ""C:\Program Files\Redemption\Redemption64.dll"""
            Write-Warning "    4. Re-run this script -- it will pick Redemption up automatically."
            Write-Warning ""
        }
    } else {
        Write-Warning "-NoRedemption specified; using the broken OOM path as a control test."
    }

    $script:MaxItemsCap = [int]$MaxItems

    $results = New-Object 'System.Collections.Generic.List[pscustomobject]'
    foreach ($upn in $script:Mailbox) {
        $r = Fix-Mailbox -Namespace $ns -Upn $upn -IsDryRun:$DryRun.IsPresent `
            -FolderFilter $Folder -RdoSession $rdo
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
