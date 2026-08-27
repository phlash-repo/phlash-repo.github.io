<#
.SYNOPSIS
    Windows helper that restores a Thunderbird mail cache to Carbonio by
    driving imap_restore.py: it finds Python, locates the Thunderbird
    profile, always DRY-RUNS FIRST so you can see what would be uploaded,
    and only uploads after you confirm.

.DESCRIPTION
    Run this on the user's laptop from the mail-restore folder. Every folder
    in the Thunderbird archive is recreated on the Carbonio server and each
    message is restored into the folder it was filed in. Messages already
    present in that folder on the server (e.g. mail received since the
    migration, or a previous partial run) are skipped, so re-running is safe.

    BEFORE RUNNING: put Thunderbird in offline mode (File > Offline > Work
    Offline) or close it, and back up the profile directory
    (%APPDATA%\Thunderbird\Profiles) - see README.md section 0.

    Needs Python 3 on the laptop. If missing, install from
    https://www.python.org/downloads/ (tick "Add python.exe to PATH") or:
        winget install Python.Python.3.12

.PARAMETER Server
    Carbonio hostname (e.g. mail.example.com). Prompted for if omitted.

.PARAMETER User
    The account to restore into (e.g. alice@example.com). Prompted for if
    omitted.

.PARAMETER ProfileDir
    A specific Thunderbird profile directory. Default: auto-detect every
    profile on this machine (--find-thunderbird).

.PARAMETER DedupeScope
    'Folder' (default): per-folder replication - a message is skipped only
    if it is already in the same folder on the server. 'Account': never
    upload a message that exists anywhere in the mailbox (for combining
    several sources - see README.md).

.PARAMETER Insecure
    Pass when the Carbonio server still has a self-signed certificate.

.PARAMETER ExtraArgs
    Any additional imap_restore.py options, e.g.
    -ExtraArgs @('--prefix','Restored','--throttle','0.1')

.EXAMPLE
    powershell -ExecutionPolicy Bypass -File .\restore_thunderbird.ps1

.EXAMPLE
    powershell -ExecutionPolicy Bypass -File .\restore_thunderbird.ps1 `
        -Server mail.example.com -User alice@example.com
#>
[CmdletBinding()]
param(
    [string]$Server,
    [string]$User,
    [string]$ProfileDir,
    [ValidateSet('Folder', 'Account')]
    [string]$DedupeScope = 'Folder',
    [switch]$Insecure,
    [string[]]$ExtraArgs = @()
)

$ErrorActionPreference = 'Stop'

function Find-Python {
    # Returns @(executable, leading-args...) or $null.
    $candidates = @()
    $candidates += , @('py', '-3')
    $candidates += , @('python')
    $candidates += , @('python3')
    foreach ($cand in $candidates) {
        $cmd = Get-Command $cand[0] -ErrorAction SilentlyContinue
        if (-not $cmd) { continue }
        $lead = @($cand | Select-Object -Skip 1)
        try {
            $ver = (& $cand[0] @lead --version 2>&1) -join ' '
        } catch { continue }
        # Skip the Microsoft Store placeholder, which prints an install hint.
        if ($ver -match 'Python 3\.') { return ,$cand }
    }
    return $null
}

$python = Find-Python
if (-not $python) {
    Write-Host 'Python 3 was not found on this machine.' -ForegroundColor Red
    Write-Host 'Install it from https://www.python.org/downloads/ (tick "Add python.exe to PATH")'
    Write-Host 'or run:  winget install Python.Python.3.12'
    Write-Host 'then run this script again.'
    exit 1
}
$pyExe = $python[0]
$pyLead = @($python | Select-Object -Skip 1)
Write-Host ("Using Python: {0}" -f ((& $pyExe @pyLead --version 2>&1) -join ' '))

$script = Join-Path $PSScriptRoot 'imap_restore.py'
if (-not (Test-Path $script)) {
    throw "imap_restore.py not found next to this script ($script)."
}

if (-not $Server) { $Server = Read-Host 'Carbonio server hostname (e.g. mail.example.com)' }
if (-not $User) { $User = Read-Host "Account to restore into (e.g. alice@example.com)" }

$sec = Read-Host "IMAP password for $User" -AsSecureString
$bstr = [Runtime.InteropServices.Marshal]::SecureStringToBSTR($sec)
try {
    $env:IMAP_PASSWORD = [Runtime.InteropServices.Marshal]::PtrToStringBSTR($bstr)
} finally {
    [Runtime.InteropServices.Marshal]::ZeroFreeBSTR($bstr)
}

$pyArgs = @($script, '--server', $Server, '--user', $User)
if ($ProfileDir) { $pyArgs += @('--thunderbird', $ProfileDir) }
else { $pyArgs += '--find-thunderbird' }
if ($DedupeScope -eq 'Account') { $pyArgs += @('--dedupe-scope', 'account') }
if ($Insecure) { $pyArgs += '--insecure' }
$pyArgs += $ExtraArgs

try {
    Write-Host ''
    Write-Host '=== Pass 1: DRY RUN (nothing is uploaded) ===' -ForegroundColor Cyan
    & $pyExe @pyLead @pyArgs --dry-run
    if ($LASTEXITCODE -ne 0) {
        Write-Host 'The dry run reported failures - check imap_restore_failures.log before continuing.' -ForegroundColor Yellow
    }

    Write-Host ''
    $go = Read-Host 'Proceed with the real upload? [y/N]'
    if ($go -notmatch '^[Yy]') {
        Write-Host 'Stopped. Nothing was uploaded.'
        exit 0
    }

    Write-Host ''
    Write-Host '=== Pass 2: uploading to the server ===' -ForegroundColor Cyan
    & $pyExe @pyLead @pyArgs
    if ($LASTEXITCODE -eq 0) {
        Write-Host ''
        Write-Host 'Finished. Check the folders in Carbonio webmail, then compare' -ForegroundColor Green
        Write-Host 'per-folder message counts against Thunderbird.' -ForegroundColor Green
    } else {
        Write-Host ''
        Write-Host 'Finished with failures - see imap_restore_failures.log.' -ForegroundColor Yellow
        Write-Host 'It is safe to re-run this script; uploaded mail will not be duplicated.'
    }
} finally {
    Remove-Item Env:IMAP_PASSWORD -ErrorAction SilentlyContinue
}
