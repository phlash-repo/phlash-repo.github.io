# Restoring lost mail to Carbonio from local Outlook / Thunderbird caches

Runbook and scripts for the scenario where the old Zimbra server (and its
archives) are gone, a fresh Carbonio server has replaced it, and the only
remaining copies of the mail are the local caches on users' machines
(Outlook OST files and Thunderbird profiles, both from IMAP accounts).

The tools provided:

| Script | Use for | Runs on |
|---|---|---|
| [`restore_thunderbird.ps1`](restore_thunderbird.ps1) | Thunderbird laptops — guided wrapper around `imap_restore.py` (dry-runs first, asks before uploading) | Windows |
| [`restore_outlook.ps1`](restore_outlook.ps1) | Outlook machines where the old profile still opens | Windows with Outlook installed |
| [`imap_restore.py`](imap_restore.py) | The engine behind the Thunderbird wrapper; also handles mbox files, Maildir, `.eml` trees (incl. `readpst` output from Outlook OST/PST) directly | Windows / macOS / Linux, Python 3.7+, no extra packages |

**Default behavior — full folder replication.** Every folder in the local
archive is recreated on the Carbonio server (including nested subfolders)
and each email is restored into the folder it was filed in. Duplicate
checking is per folder: a message is uploaded unless that same message is
already in that same folder on the server. All tools are safe to re-run —
an interrupted restore just continues where it left off.

## Quick start on a Windows laptop

1. Do **section 0 below first** (offline mode + backup) — it protects the
   only remaining copy of the mail.
2. Copy this `mail-restore` folder onto the laptop.
3. **Thunderbird laptop:** open PowerShell in the folder and run

   ```powershell
   powershell -ExecutionPolicy Bypass -File .\restore_thunderbird.ps1
   ```

   It finds Python (tells you how to install it if missing), auto-detects
   the Thunderbird profile, prompts for server/account/password, shows a
   dry run of everything it would upload, and only uploads after you
   confirm. Add `-Insecure` if the Carbonio certificate is self-signed.
4. **Outlook laptop:** follow section 3 (add the Carbonio IMAP account in
   the existing Outlook profile, then run `restore_outlook.ps1`).
5. Verify in Carbonio webmail (section 5).

---

## 0. First — protect the evidence

The local caches are now the *only* copy of the mail. Before touching
anything on a user's machine:

1. **Do not delete or "fix" the old account** in Outlook or Thunderbird, and
   do not let the client "repair" the account against the new server —
   an IMAP client that reconnects and sees empty server folders may
   resync and **delete its local copies**.
   - Thunderbird: File → Offline → Work Offline (or just disconnect from
     the network before opening it).
   - Outlook: Send/Receive tab → **Work Offline**, or start with
     `outlook.exe /safe`.
2. **Copy the raw data files to a backup drive** before doing anything else:
   - Outlook OST: `%LOCALAPPDATA%\Microsoft\Outlook\*.ost`
     (older setups: `Documents\Outlook Files\*.pst`)
   - Thunderbird profile:
     - Windows: `%APPDATA%\Thunderbird\Profiles\<xxxx>.default*`
     - macOS: `~/Library/Thunderbird/Profiles/<xxxx>.default*`
     - Linux: `~/.thunderbird/<xxxx>.default*`
3. **Check what is actually there.** Thunderbird only keeps full local
   copies of IMAP folders if *"Synchronization & Storage → Keep messages
   for this account on this computer"* was enabled (it is the default for
   recent versions). Outlook cached mode keeps a slice controlled by the
   "Mail to keep offline" slider (often 12 months, sometimes "All"). The
   dry-run modes below tell you exactly how much is recoverable per machine.

## 1. Prepare the Carbonio side

On the Carbonio server (as the `zextras` user):

```sh
# create/verify each account
carbonio prov ca alice@example.com 'TempPassw0rd!'

# make sure IMAP is available (defaults: IMAPS on 993)
carbonio prov gs $(zmhostname) zimbraImapSSLServerEnabled zimbraImapServerEnabled

# messages larger than this are rejected on upload — raise it if users have
# big attachments (value in bytes; then restart services)
carbonio prov gacf zimbraMtaMaxMessageSize
```

Also check mailbox quotas (`zimbraMailQuota`) are large enough to hold the
restored archive. Restore **one test mailbox first**, verify it in Carbonio
webmail, then roll out to the rest.

## 2. Thunderbird machines → `restore_thunderbird.ps1`

**On Windows, just run the wrapper** — it drives `imap_restore.py` for you
(finds Python, auto-detects the profile, dry-runs first, asks before
uploading):

```powershell
powershell -ExecutionPolicy Bypass -File .\restore_thunderbird.ps1
# or non-interactively:
powershell -ExecutionPolicy Bypass -File .\restore_thunderbird.ps1 `
    -Server mail.example.com -User alice@example.com
```

If Python 3 isn't installed yet: `winget install Python.Python.3.12`, or
https://www.python.org/downloads/ with "Add python.exe to PATH" ticked.

Running the engine directly works too (on Windows use `py` instead of
`python3`), on the user's machine or against a copied profile on an admin
machine. Only Python 3 is needed — no third-party packages:

```sh
# 1. see what would be restored (auto-detects profiles on this machine)
python3 imap_restore.py --server mail.example.com --user alice@example.com \
    --find-thunderbird --dry-run

# 2. do it for real
python3 imap_restore.py --server mail.example.com --user alice@example.com \
    --find-thunderbird
```

Or point it at a copied profile directory explicitly with
`--thunderbird /path/to/xxxx.default-release`.

What it does:

- scans `ImapMail/` (the offline IMAP cache) and `Mail/` (Local Folders) for
  mbox files, reconstructing subfolders from `.sbd` directories;
- skips messages Thunderbird had already deleted but never compacted away
  (`X-Mozilla-Status` expunged flag), and skips Outbox/Unsent;
- maps common folder names onto Carbonio's defaults
  (`Sent Items`/`Sent Mail` → `Sent`, `Deleted Items` → `Trash`,
  `Spam` → `Junk`, …) — add your own with `--map "Old Name=New/Name"`;
- uploads via IMAP APPEND preserving each message's original date and
  read/unread state;
- skips anything whose Message-ID is already in the target folder, so
  re-runs are safe;
- writes failures to `imap_restore_failures.log` and carries on.

Useful options: `--prefix Restored` (restore under a `Restored/` folder for
review instead of straight into the live folders), `--insecure` (self-signed
certificate), `--throttle 0.1` (be gentle on the server), `--max-size`,
`--no-dedupe`. The password can be supplied via the `IMAP_PASSWORD`
environment variable instead of the interactive prompt.

For many users, loop it from a shell with a CSV of `user,password,profile-path`.

## 3. Outlook machines → `restore_outlook.ps1`

**Path A — the old profile still opens (preferred).** OST files only open
inside the profile that created them, so if the old account is still
configured, work inside it:

1. Start Outlook **offline** (Send/Receive → Work Offline) so nothing tries
   to sync against a dead server. Confirm the old mail is visible.
2. File → Add Account → add the new Carbonio account as **IMAP**
   (server: your Carbonio host, port 993, SSL). Let its folders appear.
3. Run:

   ```powershell
   powershell -ExecutionPolicy Bypass -File .\restore_outlook.ps1 -DryRun
   powershell -ExecutionPolicy Bypass -File .\restore_outlook.ps1 -Dedupe
   ```

   The script prompts you to pick the source store (the old account /
   "Outlook Data File") and the target store (the Carbonio account), then
   recreates the folder tree and copies every mail item across.
4. Take Outlook back **online** and leave it running — it uploads the copied
   messages to Carbonio over IMAP in the background. Watch
   Send/Receive → Show Progress, and verify in Carbonio webmail.

Notes: only mail items are copied (calendar/contacts can't travel over
IMAP — export those as .ics/.csv separately). `-Dedupe` makes re-runs
safe; `-ThrottleMs 50` slows the copy if Outlook chokes; `-PstPath` attaches
an old archive `.pst` first.

**Path B — the profile is gone / Outlook won't open the OST.** An orphaned
OST cannot be re-attached to Outlook. Convert it with `readpst` (part of
libpst, which reads OST as well as PST — on Linux/WSL:
`apt install pst-utils`, macOS: `brew install libpst`):

```sh
readpst -r -e -o out/ alice.ost      # folder tree of .eml files
python3 imap_restore.py --server mail.example.com --user alice@example.com \
    --eml out/ --dry-run             # then again without --dry-run
```

The same recipe works for any old `.pst` archive files you find lying
around.

## 4. Combining sources — Outlook + Thunderbird + mail already on the server

All of the above can safely target the same mailbox. Both tools check the
target for duplicates before uploading (Message-ID, with a Date/From/Subject
fingerprint as fallback), so:

- **Mail the new Carbonio server has already received** stays untouched, and
  a local cached copy of the same message is skipped, not duplicated.
- **The same message present in both Outlook and Thunderbird** is only
  restored once — whichever source runs first wins.
- **Interrupted runs** can simply be re-run.

By default the duplicate check is *per folder*: a message is skipped only if
it is already in the same folder on the server. That is right for a plain
restore, but when combining sources the same message often lives in
*different* folders in different places — e.g. still in `INBOX` on the
server (redelivered after the migration) but filed under a project folder
in someone's Thunderbird, or filed differently by the Outlook and
Thunderbird users. For that, switch to **account-wide dedupe**, which
indexes every folder first and never uploads a message that exists
anywhere in the mailbox:

```sh
python3 imap_restore.py --server mail.example.com --user alice@example.com \
    --thunderbird ... --dedupe-scope account
```

```powershell
powershell -ExecutionPolicy Bypass -File .\restore_outlook.ps1 -DedupeScope Account
```

Recommended order per mailbox:

1. Leave whatever the new server has already received exactly as it is —
   the restore works around it.
2. Restore the **richest/most complete archive first** (usually the one
   with the longest history); with account-wide dedupe its filing wins.
3. Run the remaining sources afterwards — they only fill in the gaps.

Two caveats:

- With `-DedupeScope Account` the Outlook script builds its index from
  Outlook's local cache of the Carbonio account, so let that account
  **fully sync** (all folders, Send/Receive quiet) before running.
- Account-wide dedupe keeps *one* copy per mailbox. If you genuinely want
  the same message in two folders (rare), stay with the per-folder default
  for that source.

## 5. Verify

- Compare per-folder message counts between the client and Carbonio webmail.
- Check `imap_restore_failures.log` / the PowerShell failure count; the
  usual causes are oversized messages (raise `zimbraMtaMaxMessageSize`) or
  corrupt messages in the cache (usually unrecoverable).
- Once a mailbox is verified, take a fresh server-side backup — and this
  time, keep one off the mail server.

## Limitations

- Only what the client actually cached locally can be restored — headers-only
  folders (Outlook slider, Thunderbird sync disabled) yield nothing.
- Calendar, contacts and tasks are out of scope for IMAP; recover them from
  the clients' native export functions if needed.
- Folder ACLs/shares, tags and filters from the old Zimbra are gone; only the
  mail itself comes back.
