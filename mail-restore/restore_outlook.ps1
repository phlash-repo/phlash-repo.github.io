<#
.SYNOPSIS
    Restore mail from a local Outlook data store (cached OST or a PST) into a
    newly configured Carbonio IMAP account, by copying items store-to-store
    inside Outlook and letting Outlook sync them to the server.

.DESCRIPTION
    Scenario: the mail server died, and the only copy of the mail is the
    Outlook offline cache (.ost) on each user's machine.

    IMPORTANT - BEFORE RUNNING ANYTHING:
      * Do NOT remove the old account from Outlook and do NOT delete the old
        Outlook profile. An .ost file is tied to its account/profile - once
        the account is removed, Outlook cannot reopen the orphaned .ost.
        (If that has already happened, convert the .ost with `readpst` and
        use imap_restore.py instead - see README.md.)
      * Make a file-level backup copy of the .ost/.pst first
        (%LOCALAPPDATA%\Microsoft\Outlook\*.ost). If Outlook refuses to
        start because the old server is unreachable, start it in offline
        mode: File > Work Offline, or `outlook.exe /safe`.

    Steps:
      1. In the EXISTING profile (or a new one that still has the old data
         file), add the new Carbonio account:
         File > Add Account > IMAP, server = your Carbonio host, port 993/SSL.
      2. Let the new account's folder list appear in the folder pane.
      3. Run this script. It walks every mail folder of the source store,
         creates the same folder tree in the Carbonio store, and copies the
         messages across. Outlook then uploads them to Carbonio in the
         background.
      4. LEAVE OUTLOOK OPEN until the sync has finished (send/receive
         progress is visible under Send/Receive > Show Progress; folder
         'sync issues' should stay empty).

    Only mail items are copied. Calendars/contacts/tasks cannot be restored
    over IMAP - export those separately if needed.

.PARAMETER SourceStore
    Display name of the source store as shown in Outlook's folder pane
    (e.g. "alice@oldserver.com" or "Outlook Data File"). If omitted, the
    script lists all stores and prompts.

.PARAMETER TargetStore
    Display name of the new Carbonio account's store. If omitted, prompts.

.PARAMETER PstPath
    Optional path to a .pst file to attach first (for archives exported
    earlier). Do not point this at an .ost - Outlook cannot open those.

.PARAMETER DryRun
    Walk the folders and report counts without copying anything.

.PARAMETER Dedupe
    Before copying each folder, read the Message-IDs already present in the
    target folder and skip messages that are already there. Makes re-runs
    safe after an interruption (slower on big folders).

.PARAMETER DedupeScope
    'Folder' (default): a message is skipped only if it already exists in
    the SAME folder on the target. 'Account' (implies -Dedupe): the whole
    target store is indexed first and a message that exists ANYWHERE in it
    is never copied - use this when combining several sources (this script,
    a Thunderbird restore via imap_restore.py, and mail already received by
    the new server) that may have filed the same message in different
    folders. Let Outlook FULLY SYNC the Carbonio account before running,
    since the index is built from Outlook's local cache of the target.

.PARAMETER ThrottleMs
    Pause this many milliseconds between item copies (default 0). Use a
    small value (e.g. 50) if Outlook or the server struggles with the load.

.EXAMPLE
    powershell -ExecutionPolicy Bypass -File .\restore_outlook.ps1 -DryRun

.EXAMPLE
    powershell -ExecutionPolicy Bypass -File .\restore_outlook.ps1 `
        -SourceStore "alice@oldserver.com" -TargetStore "alice@newserver.com" -Dedupe

.EXAMPLE
    # Combining with other sources / mail already on the server: never copy a
    # message that already exists anywhere in the Carbonio account.
    powershell -ExecutionPolicy Bypass -File .\restore_outlook.ps1 -DedupeScope Account
#>
[CmdletBinding()]
param(
    [string]$SourceStore,
    [string]$TargetStore,
    [string]$PstPath,
    [switch]$DryRun,
    [switch]$Dedupe,
    [ValidateSet('Folder', 'Account')]
    [string]$DedupeScope = 'Folder',
    [int]$ThrottleMs = 0,
    [string[]]$SkipFolders = @('Outbox', 'Sync Issues', 'Conflicts',
        'Local Failures', 'Server Failures', 'RSS Feeds', 'RSS Subscriptions',
        'Conversation History', 'Quick Step Settings', 'Yammer Root',
        'ExternalContacts', 'Files', 'Journal', 'Notes', 'Tasks',
        'Calendar', 'Contacts', 'Suggested Contacts')
)

$ErrorActionPreference = 'Stop'
$olMailItem = 0        # OlItemType.olMailItem / DefaultItemType
$olMailClass = 43      # OlObjectClass.olMail
$PR_INTERNET_MESSAGE_ID = 'http://schemas.microsoft.com/mapi/proptag/0x1035001F'

# Source folder name -> preferred Carbonio folder name.
$FolderNameMap = @{
    'Sent Items'       = 'Sent'
    'Sent'             = 'Sent'
    'Deleted Items'    = 'Trash'
    'Trash'            = 'Trash'
    'Junk E-mail'      = 'Junk'
    'Junk Email'       = 'Junk'
    'Spam'             = 'Junk'
    'Inbox'            = 'Inbox'
    'Drafts'           = 'Drafts'
}

$script:Copied = 0
$script:Skipped = 0
$script:Failed = 0

function Get-Store {
    param($Namespace, [string]$Name, [string]$Role)
    $stores = @($Namespace.Stores)
    if ($Name) {
        $match = $stores | Where-Object { $_.DisplayName -eq $Name }
        if (-not $match) {
            Write-Host "No store called '$Name'. Available stores:" -ForegroundColor Yellow
            $stores | ForEach-Object { Write-Host "  - $($_.DisplayName)" }
            throw "Store '$Name' not found."
        }
        return $match
    }
    Write-Host "`nAvailable Outlook stores:" -ForegroundColor Cyan
    for ($i = 0; $i -lt $stores.Count; $i++) {
        Write-Host ("  [{0}] {1}  ({2})" -f $i, $stores[$i].DisplayName, $stores[$i].FilePath)
    }
    $sel = Read-Host "Select the $Role store by number"
    return $stores[[int]$sel]
}

function Get-TargetChildFolder {
    # Find (case-insensitively) or create a child folder in the target store.
    param($Parent, [string]$Name)
    foreach ($f in @($Parent.Folders)) {
        if ($f.Name -ieq $Name) { return $f }
    }
    if ($DryRun) {
        Write-Host "    [dry-run] would create folder '$Name'"
        return $null
    }
    return $Parent.Folders.Add($Name)
}

function Add-FolderMessageIds {
    param($Folder, $Ids)
    foreach ($item in @($Folder.Items)) {
        try {
            if ($item.Class -eq $olMailClass) {
                $id = $item.PropertyAccessor.GetProperty($PR_INTERNET_MESSAGE_ID)
                if ($id) { [void]$Ids.Add($id) }
            }
        } catch { }
    }
}

function Get-TargetMessageIds {
    # The leading comma stops PowerShell unrolling the HashSet on return.
    param($Folder)
    $ids = New-Object 'System.Collections.Generic.HashSet[string]'
    if ($null -ne $Folder) { Add-FolderMessageIds -Folder $Folder -Ids $ids }
    return ,$ids
}

function Get-AllStoreMessageIds {
    param($Root)
    $ids = New-Object 'System.Collections.Generic.HashSet[string]'
    $stack = New-Object System.Collections.Stack
    foreach ($f in @($Root.Folders)) { $stack.Push($f) }
    while ($stack.Count -gt 0) {
        $f = $stack.Pop()
        $isMail = $true
        try { $isMail = ($f.DefaultItemType -eq $olMailItem) } catch { }
        if ($isMail) { Add-FolderMessageIds -Folder $f -Ids $ids }
        foreach ($sub in @($f.Folders)) { $stack.Push($sub) }
    }
    return ,$ids
}

function Copy-MailFolder {
    param($SourceFolder, $TargetParent, [string]$Path)

    if ($SkipFolders -contains $SourceFolder.Name) {
        Write-Host "  Skipping '$Path' (in skip list)" -ForegroundColor DarkGray
        return
    }

    $isMailFolder = $true
    try { $isMailFolder = ($SourceFolder.DefaultItemType -eq $olMailItem) } catch { }

    $targetFolder = $null
    if ($isMailFolder) {
        $targetName = $SourceFolder.Name
        if ($FolderNameMap.ContainsKey($SourceFolder.Name)) {
            $targetName = $FolderNameMap[$SourceFolder.Name]
        }
        $targetFolder = Get-TargetChildFolder -Parent $TargetParent -Name $targetName

        $total = $SourceFolder.Items.Count
        Write-Host "  $Path  ($total items)"
        if ($total -gt 0 -and -not $DryRun) {
            $existing = $null
            if ($null -ne $script:GlobalIds) {
                $existing = $script:GlobalIds
            } elseif ($Dedupe) {
                $existing = Get-TargetMessageIds -Folder $targetFolder
                if ($existing.Count -gt 0) {
                    Write-Host "    target already holds $($existing.Count) message(s)"
                }
            }
            $n = 0
            # Snapshot item EntryIDs first: the live Items collection reorders
            # itself while items are added elsewhere in the same session.
            $entryIds = @()
            foreach ($item in @($SourceFolder.Items)) {
                try { $entryIds += $item.EntryID } catch { $script:Failed++ }
            }
            $ns = $SourceFolder.Session
            foreach ($eid in $entryIds) {
                $n++
                if ($n % 50 -eq 0) {
                    Write-Progress -Activity "Copying $Path" -Status "$n / $total" `
                        -PercentComplete ([int](100 * $n / [math]::Max($total,1)))
                }
                try {
                    $item = $ns.GetItemFromID($eid)
                    if ($item.Class -ne $olMailClass) { $script:Skipped++; continue }
                    $mid = $null
                    if ($null -ne $existing) {
                        try { $mid = $item.PropertyAccessor.GetProperty($PR_INTERNET_MESSAGE_ID) } catch { }
                        if ($mid -and $existing.Contains($mid)) { $script:Skipped++; continue }
                    }
                    $copy = $item.Copy()
                    [void]$copy.Move($targetFolder)
                    if ($mid) { [void]$existing.Add($mid) }
                    $script:Copied++
                    if ($ThrottleMs -gt 0) { Start-Sleep -Milliseconds $ThrottleMs }
                } catch {
                    $script:Failed++
                    Write-Host "    FAILED: $($_.Exception.Message)" -ForegroundColor Red
                }
            }
            Write-Progress -Activity "Copying $Path" -Completed
        } elseif ($total -gt 0) {
            Write-Host "    [dry-run] would copy up to $total item(s)"
        }
    } else {
        Write-Host "  Skipping '$Path' (not a mail folder)" -ForegroundColor DarkGray
    }

    # Recurse into subfolders; non-mail parents still get walked in case a
    # mail folder is nested underneath.
    foreach ($sub in @($SourceFolder.Folders)) {
        $nextParent = if ($targetFolder) { $targetFolder } else { $TargetParent }
        Copy-MailFolder -SourceFolder $sub -TargetParent $nextParent -Path "$Path\$($sub.Name)"
    }
}

# --------------------------------------------------------------------------

Write-Host 'Attaching to Outlook...' -ForegroundColor Cyan
$outlook = New-Object -ComObject Outlook.Application
$namespace = $outlook.GetNamespace('MAPI')

if ($PstPath) {
    if ($PstPath -match '\.ost$') {
        throw 'Outlook cannot attach an .ost file. Use the existing profile that owns it, or convert it with readpst (see README.md).'
    }
    Write-Host "Attaching PST: $PstPath"
    $namespace.AddStore($PstPath)
}

$src = Get-Store -Namespace $namespace -Name $SourceStore -Role 'SOURCE (old mail)'
$dst = Get-Store -Namespace $namespace -Name $TargetStore -Role 'TARGET (new Carbonio account)'

if ($src.StoreID -eq $dst.StoreID) {
    throw 'Source and target are the same store - aborting.'
}

Write-Host "`nSource : $($src.DisplayName)  ($($src.FilePath))"
Write-Host "Target : $($dst.DisplayName)"
if ($DryRun) { Write-Host '*** DRY RUN - nothing will be copied ***' -ForegroundColor Yellow }
Write-Host ''

$srcRoot = $src.GetRootFolder()
$dstRoot = $dst.GetRootFolder()

$script:GlobalIds = $null
if ($DedupeScope -eq 'Account') {
    $Dedupe = $true
    Write-Host 'Indexing existing messages across the whole target store (this reads' -ForegroundColor Cyan
    Write-Host "Outlook's local cache - make sure the Carbonio account has fully synced)..." -ForegroundColor Cyan
    $script:GlobalIds = Get-AllStoreMessageIds -Root $dstRoot
    Write-Host "  $($script:GlobalIds.Count) existing message(s) indexed"
}

foreach ($folder in @($srcRoot.Folders)) {
    Copy-MailFolder -SourceFolder $folder -TargetParent $dstRoot -Path $folder.Name
}

Write-Host ''
Write-Host ("Done. copied={0}  skipped={1}  failed={2}" -f $script:Copied, $script:Skipped, $script:Failed) -ForegroundColor Green
if (-not $DryRun) {
    Write-Host @'

NEXT STEPS
  * Leave Outlook OPEN. The messages are now in the Carbonio account's local
    cache and Outlook will upload them over IMAP in the background.
  * Watch Send/Receive > Show Progress until it goes quiet, and check the
    folders on the server (e.g. via Carbonio webmail) before closing.
  * If some folders did not sync, right-click the Carbonio account >
    'Update Folder List', then Send/Receive All Folders (F9).
'@
}
