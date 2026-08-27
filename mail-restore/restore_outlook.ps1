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

.PARAMETER ThrottleMs
    Pause this many milliseconds between item copies (default 0). Use a
    small value (e.g. 50) if Outlook or the server struggles with the load.

.EXAMPLE
    powershell -ExecutionPolicy Bypass -File .\restore_outlook.ps1 -DryRun

.EXAMPLE
    powershell -ExecutionPolicy Bypass -File .\restore_outlook.ps1 `
        -SourceStore "alice@oldserver.com" -TargetStore "alice@newserver.com" -Dedupe
#>
[CmdletBinding()]
param(
    [string]$SourceStore,
    [string]$TargetStore,
    [string]$PstPath,
    [switch]$DryRun,
    [switch]$Dedupe,
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

function Get-TargetMessageIds {
    param($Folder)
    $ids = New-Object 'System.Collections.Generic.HashSet[string]'
    if ($null -eq $Folder) { return $ids }
    foreach ($item in @($Folder.Items)) {
        try {
            if ($item.Class -eq $olMailClass) {
                $id = $item.PropertyAccessor.GetProperty($PR_INTERNET_MESSAGE_ID)
                if ($id) { [void]$ids.Add($id) }
            }
        } catch { }
    }
    return $ids
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
            if ($Dedupe) {
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
                    if ($Dedupe) {
                        $mid = $null
                        try { $mid = $item.PropertyAccessor.GetProperty($PR_INTERNET_MESSAGE_ID) } catch { }
                        if ($mid -and $existing.Contains($mid)) { $script:Skipped++; continue }
                    }
                    $copy = $item.Copy()
                    [void]$copy.Move($targetFolder)
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
