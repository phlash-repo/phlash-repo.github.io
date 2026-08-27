#!/usr/bin/env python3
"""
imap_restore.py - restore locally cached mail back to a Carbonio (or any IMAP) server.

Designed for the "server died, only copies are the local caches" scenario:

  * Thunderbird profiles  : offline IMAP cache + Local Folders (mbox format)
  * Generic mbox files    : single file or a directory tree of mbox files
  * Maildir               : standard Maildir tree
  * .eml directory trees  : e.g. the output of `readpst -r -e` run against an
                            Outlook OST/PST file

Messages are uploaded with IMAP APPEND, preserving:
  * the original message date (INTERNALDATE is taken from the Date: header)
  * read/unread state where the source records it (Thunderbird X-Mozilla-Status,
    Maildir info flags); .eml files are marked read by default
  * folder hierarchy (Thunderbird `.sbd` directories become subfolders)

It de-duplicates on Message-ID against what is already in the target folder,
so it is safe to re-run after a partial/interrupted restore.

Examples
--------
Dry run against an auto-detected Thunderbird profile:

    python3 imap_restore.py --server mail.example.com --user alice@example.com \
        --find-thunderbird --dry-run

Real restore from a specific Thunderbird profile:

    python3 imap_restore.py --server mail.example.com --user alice@example.com \
        --thunderbird "~/.thunderbird/abcd1234.default-release"

Restore mail exported from an Outlook OST with readpst:

    readpst -r -e -o out/ alice.ost
    python3 imap_restore.py --server mail.example.com --user alice@example.com \
        --eml out/

Requires only the Python 3 standard library (Python 3.7+).
"""

import argparse
import base64
import getpass
import hashlib
import imaplib
import mailbox
import os
import re
import socket
import ssl
import sys
import time
from datetime import datetime, timezone
from email.parser import BytesHeaderParser
from email.utils import parsedate_to_datetime

# Large folders produce long FETCH/SEARCH responses.
imaplib._MAXLINE = 20_000_000

HEADER_PARSER = BytesHeaderParser()

# Well-known folder names (lower-cased) -> Carbonio/Zimbra default folder.
DEFAULT_FOLDER_MAP = {
    "inbox": "INBOX",
    "sent": "Sent",
    "sent items": "Sent",
    "sent mail": "Sent",
    "sent messages": "Sent",
    "sent-mail": "Sent",
    "drafts": "Drafts",
    "trash": "Trash",
    "deleted items": "Trash",
    "deleted messages": "Trash",
    "junk": "Junk",
    "junk e-mail": "Junk",
    "junk email": "Junk",
    "spam": "Junk",
    "bulk mail": "Junk",
    "archives": "Archive",
    "archive": "Archive",
}

# Folders that make no sense to restore.
SKIP_FOLDERS = {"outbox", "unsent messages", "sync issues", "conflicts",
                "local failures", "server failures", "rss feeds",
                "conversation history", "search folders"}

# Thunderbird X-Mozilla-Status bits
MOZ_READ = 0x0001
MOZ_EXPUNGED = 0x0008

# Thunderbird files that live next to mbox files and must be ignored.
TB_SKIP_EXT = {".msf", ".dat", ".sqlite", ".json", ".html", ".ini",
               ".sqlite-wal", ".sqlite-shm", ".bak", ".mozmsgs"}
TB_SKIP_NAMES = {"msgfilterrules.dat", "popstate.dat", "filterlog.html"}


def log(msg):
    print(msg, flush=True)


class FailureLog:
    def __init__(self, path):
        self.path = path
        self._fh = None
        self.count = 0

    def record(self, folder, ident, error):
        if self._fh is None:
            self._fh = open(self.path, "a", encoding="utf-8")
        self._fh.write("%s\t%s\t%s\t%s\n"
                       % (datetime.now().isoformat(timespec="seconds"),
                          folder, ident, error))
        self._fh.flush()
        self.count += 1

    def close(self):
        if self._fh:
            self._fh.close()


# --------------------------------------------------------------------------
# IMAP helpers
# --------------------------------------------------------------------------

def imap_utf7_encode(name):
    """Encode a folder name using IMAP modified UTF-7 (RFC 3501 5.1.3)."""
    out = bytearray()
    buf = ""

    def flush():
        nonlocal buf
        if buf:
            b64 = base64.b64encode(buf.encode("utf-16-be"))
            out.extend(b"&" + b64.rstrip(b"=").replace(b"/", b",") + b"-")
            buf = ""

    for ch in name:
        o = ord(ch)
        if 0x20 <= o <= 0x7E:
            flush()
            if ch == "&":
                out.extend(b"&-")
            else:
                out.append(o)
        else:
            buf += ch
    flush()
    return out.decode("ascii")


def imap_quote(name):
    return '"' + name.replace("\\", "\\\\").replace('"', '\\"') + '"'


DEDUPE_FIELDS = "MESSAGE-ID DATE FROM SUBJECT"


def header_block(raw):
    m = re.search(rb"\r?\n\r?\n", raw)
    return raw[:m.start()] if m else raw[:16384]


def dedupe_key(headers):
    """Identity of a message for duplicate detection: its Message-ID, or a
    hash of Date/From/Subject when it has none."""
    h = headers.replace(b"\r\n", b"\n")
    h = re.sub(rb"\n[ \t]+", b" ", h)  # unfold
    m = re.search(rb"^Message-ID:[ \t]*(<[^>]+>)", h, re.I | re.M)
    if m:
        return m.group(1).strip()
    parts = []
    for name in (rb"Date", rb"From", rb"Subject"):
        mm = re.search(rb"^" + name + rb":[ \t]*(.*?)[ \t]*$", h, re.I | re.M)
        parts.append(mm.group(1).strip() if mm else b"")
    if not any(parts):
        return None
    return b"F:" + hashlib.sha1(b"\x00".join(parts)).hexdigest().encode()


class Uploader:
    def __init__(self, args, failures):
        self.args = args
        self.failures = failures
        self.imap = None
        self.delimiter = "/"
        self.known_folders = set()
        self.msgid_cache = {}     # folder -> set of message-ids
        self.global_keys = None   # account-wide key set (--dedupe-scope account)
        self.stats = {"uploaded": 0, "skipped_dupe": 0, "failed": 0,
                      "skipped_other": 0}

    # -- connection -------------------------------------------------------

    def connect(self):
        a = self.args
        log("Connecting to %s:%d ..." % (a.server, a.port))
        if a.no_ssl:
            self.imap = imaplib.IMAP4(a.server, a.port)
            if a.starttls:
                self.imap.starttls(ssl.create_default_context())
        else:
            ctx = ssl.create_default_context()
            if a.insecure:
                ctx.check_hostname = False
                ctx.verify_mode = ssl.CERT_NONE
            self.imap = imaplib.IMAP4_SSL(a.server, a.port, ssl_context=ctx)
        self.imap.login(a.user, a.password)
        self._detect_delimiter()
        log("Logged in as %s (hierarchy delimiter %r)" % (a.user, self.delimiter))

    def reconnect(self):
        try:
            if self.imap is not None:
                self.imap.logout()
        except Exception:
            pass
        for attempt in range(5):
            try:
                self.connect()
                return
            except Exception as e:
                wait = 2 ** (attempt + 1)
                log("  reconnect failed (%s), retrying in %ds" % (e, wait))
                time.sleep(wait)
        raise RuntimeError("could not reconnect to server")

    def _detect_delimiter(self):
        try:
            typ, data = self.imap.list('""', '""')
            if typ == "OK" and data and data[0]:
                m = re.search(rb'\(.*?\)\s+"?([^"\s]+)"?\s', data[0])
                if m:
                    self.delimiter = m.group(1).decode()
        except Exception:
            pass

    # -- folders ----------------------------------------------------------

    def map_folder(self, parts):
        """Map a source folder path (list of components) to a target path string."""
        if not parts:
            return None
        if parts[0].lower() in SKIP_FOLDERS:
            return None
        # user-supplied explicit mappings first (match on full joined path)
        joined = "/".join(parts)
        for src, dst in self.args.map:
            if joined.lower() == src.lower():
                return dst.replace("/", self.delimiter) if dst else None
        # well-known top-level names
        head = DEFAULT_FOLDER_MAP.get(parts[0].lower(), parts[0])
        mapped = [head] + list(parts[1:])
        if self.args.prefix and mapped[0] != "INBOX":
            mapped = [self.args.prefix] + mapped
        elif self.args.prefix and mapped[0] == "INBOX":
            mapped = [self.args.prefix, "INBOX"] + list(parts[1:])
        return self.delimiter.join(mapped)

    def ensure_folder(self, folder):
        if folder in self.known_folders:
            return
        parts = folder.split(self.delimiter)
        for i in range(1, len(parts) + 1):
            partial = self.delimiter.join(parts[:i])
            if partial in self.known_folders or partial.upper() == "INBOX":
                continue
            enc = imap_quote(imap_utf7_encode(partial))
            typ, data = self.imap.select(enc, readonly=True)
            if typ == "OK":
                self.imap.close()
            else:
                if self.args.dry_run:
                    log("  [dry-run] would create folder %r" % partial)
                else:
                    typ, data = self.imap.create(enc)
                    if typ != "OK" and b"ALREADYEXISTS" not in (data[0] or b""):
                        log("  warning: CREATE %r -> %s %s" % (partial, typ, data))
                    self.imap.subscribe(enc)
            self.known_folders.add(partial)
        self.known_folders.add(folder)

    def _folder_keys(self, enc):
        """Dedupe keys of every message in one (already encoded) mailbox."""
        ids = set()
        typ, _ = self.imap.select(enc, readonly=True)
        if typ != "OK":
            return None
        typ, data = self.imap.uid("search", None, "ALL")
        uids = data[0].split() if typ == "OK" and data and data[0] else []
        for i in range(0, len(uids), 500):
            batch = b",".join(uids[i:i + 500])
            typ, resp = self.imap.uid(
                "fetch", batch,
                "(BODY.PEEK[HEADER.FIELDS (%s)])" % DEDUPE_FIELDS)
            if typ != "OK":
                continue
            for part in resp:
                if isinstance(part, tuple) and part[1]:
                    key = dedupe_key(part[1])
                    if key:
                        ids.add(key)
        try:
            self.imap.close()
        except Exception:
            pass
        return ids

    def existing_msgids(self, folder):
        if folder in self.msgid_cache:
            return self.msgid_cache[folder]
        ids = set()
        # Account-wide scope already indexed every folder up front.
        if not self.args.no_dedupe and self.args.dedupe_scope == "folder":
            enc = imap_quote(imap_utf7_encode(folder))
            got = self._folder_keys(enc)
            if got:
                ids = got
                log("  %d existing message(s) indexed in %r on server"
                    % (len(ids), folder))
        self.msgid_cache[folder] = ids
        return ids

    def scan_account_keys(self):
        """Index the dedupe keys of every message in every folder of the
        target account (--dedupe-scope account)."""
        self.global_keys = set()
        log("Indexing existing messages across the whole account ...")
        typ, data = self.imap.list('""', '"*"')
        if typ != "OK":
            log("  warning: LIST failed, falling back to per-folder dedupe")
            self.args.dedupe_scope = "folder"
            return
        folders = 0
        for line in data or []:
            if isinstance(line, tuple):      # literal-quoted mailbox name
                flags, name = line[0], line[1]
            else:
                if not line:
                    continue
                m = re.match(rb'\(([^)]*)\)\s+(?:"[^"]*"|NIL)\s+(.+)$', line)
                if not m:
                    continue
                flags, name = m.group(1), m.group(2).strip()
                if name.startswith(b'"') and name.endswith(b'"'):
                    name = name[1:-1].replace(b'\\"', b'"').replace(b"\\\\", b"\\")
            if b"\\Noselect" in flags or b"\\NoSelect" in flags:
                continue
            got = self._folder_keys(imap_quote(name.decode("ascii", "replace")))
            if got is not None:
                folders += 1
                self.global_keys |= got
        log("  %d existing message(s) indexed across %d folder(s)"
            % (len(self.global_keys), folders))

    # -- messages ---------------------------------------------------------

    def upload(self, folder, raw, flags, when, ident):
        """Append one raw RFC822 message (bytes) to `folder`."""
        a = self.args
        key = dedupe_key(header_block(raw))
        seen = self.existing_msgids(folder)
        if key and (key in seen
                    or (self.global_keys is not None
                        and key in self.global_keys)):
            self.stats["skipped_dupe"] += 1
            return
        if a.max_size and len(raw) > a.max_size:
            self.stats["skipped_other"] += 1
            self.failures.record(folder, ident,
                                 "skipped: %d bytes exceeds --max-size" % len(raw))
            return
        if a.dry_run:
            self.stats["uploaded"] += 1
            if key:
                seen.add(key)
                if self.global_keys is not None:
                    self.global_keys.add(key)
            return

        body = re.sub(rb"\r?\n", b"\r\n", raw)
        flag_str = "(" + " ".join(flags) + ")" if flags else None
        date = imaplib.Time2Internaldate(when.timestamp()) if when else None
        enc = imap_quote(imap_utf7_encode(folder))

        for attempt in range(3):
            try:
                typ, data = self.imap.append(enc, flag_str, date, body)
                if typ == "OK":
                    self.stats["uploaded"] += 1
                    if key:
                        seen.add(key)
                        if self.global_keys is not None:
                            self.global_keys.add(key)
                    if a.throttle:
                        time.sleep(a.throttle)
                    return
                raise RuntimeError("APPEND -> %s %s" % (typ, data))
            except (imaplib.IMAP4.abort, socket.error, ssl.SSLError,
                    BrokenPipeError, ConnectionError) as e:
                log("  connection problem (%s), reconnecting..." % e)
                self.reconnect()
            except Exception as e:
                self.stats["failed"] += 1
                self.failures.record(folder, ident, str(e))
                return
        self.stats["failed"] += 1
        self.failures.record(folder, ident, "gave up after reconnect attempts")


def message_date(raw, fallback=None):
    try:
        headers = HEADER_PARSER.parsebytes(raw[:16384])
        d = headers.get("Date")
        if d:
            dt = parsedate_to_datetime(d)
            if dt.tzinfo is None:
                dt = dt.replace(tzinfo=timezone.utc)
            # Guard against absurd values that some servers reject.
            if 1980 <= dt.year <= datetime.now().year + 1:
                return dt
    except Exception:
        pass
    return fallback or datetime.now(timezone.utc)


# --------------------------------------------------------------------------
# Sources
# --------------------------------------------------------------------------

def looks_like_mbox(path):
    try:
        if os.path.getsize(path) == 0:
            return False
        with open(path, "rb") as fh:
            return fh.read(5) == b"From "
    except OSError:
        return False


def iter_mbox(path):
    """Yield (raw_bytes, flags, ident) from an mbox file, honouring
    Thunderbird X-Mozilla-Status where present."""
    box = mailbox.mbox(path, create=False)
    try:
        for key in box.iterkeys():
            try:
                raw = box.get_bytes(key)
            except Exception as e:
                yield None, None, "%s[%s]: unreadable (%s)" % (path, key, e)
                continue
            flags = []
            m = re.search(rb"^X-Mozilla-Status:\s*([0-9a-fA-F]{1,8})",
                          raw[:8192], re.M)
            if m:
                status = int(m.group(1), 16)
                if status & MOZ_EXPUNGED:
                    continue  # deleted but folder never compacted
                if status & MOZ_READ:
                    flags.append("\\Seen")
            else:
                flags.append("\\Seen")
            yield raw, flags, "%s[%s]" % (os.path.basename(path), key)
    finally:
        box.close()


def thunderbird_roots(profile):
    roots = []
    for sub in ("ImapMail", "Mail"):
        base = os.path.join(profile, sub)
        if os.path.isdir(base):
            for server in sorted(os.listdir(base)):
                d = os.path.join(base, server)
                if os.path.isdir(d):
                    roots.append(d)
    return roots


def scan_thunderbird_root(root):
    """Yield (folder_parts, mbox_path) for every mbox under one account dir."""
    for dirpath, dirnames, filenames in os.walk(root):
        dirnames[:] = [d for d in dirnames if d.endswith(".sbd")]
        rel = os.path.relpath(dirpath, root)
        parts = [] if rel == "." else [p[:-4] if p.endswith(".sbd") else p
                                       for p in rel.split(os.sep)]
        for fn in sorted(filenames):
            ext = os.path.splitext(fn)[1].lower()
            if ext in TB_SKIP_EXT or fn.lower() in TB_SKIP_NAMES:
                continue
            path = os.path.join(dirpath, fn)
            if looks_like_mbox(path):
                yield parts + [fn], path


def find_thunderbird_profiles():
    candidates = []
    home = os.path.expanduser("~")
    for base in (os.path.join(os.environ.get("APPDATA", ""),
                              "Thunderbird", "Profiles"),
                 os.path.join(home, "Library", "Thunderbird", "Profiles"),
                 os.path.join(home, ".thunderbird")):
        if base and os.path.isdir(base):
            for d in sorted(os.listdir(base)):
                p = os.path.join(base, d)
                if os.path.isdir(p) and (
                        os.path.isdir(os.path.join(p, "ImapMail"))
                        or os.path.isdir(os.path.join(p, "Mail"))):
                    candidates.append(p)
    return candidates


def iter_maildir(path):
    box = mailbox.Maildir(path, create=False)
    for key in box.iterkeys():
        msg = box.get_message(key)
        raw = bytes(msg)
        flags = []
        if "S" in msg.get_flags():
            flags.append("\\Seen")
        if "F" in msg.get_flags():
            flags.append("\\Flagged")
        yield raw, flags, "%s[%s]" % (path, key)


def iter_eml_tree(root):
    """Yield (folder_parts, [eml files]) for a directory tree of .eml files."""
    for dirpath, dirnames, filenames in os.walk(root):
        dirnames.sort()
        emls = sorted(f for f in filenames if f.lower().endswith(".eml"))
        if not emls:
            continue
        rel = os.path.relpath(dirpath, root)
        parts = [] if rel == "." else rel.split(os.sep)
        yield parts, [os.path.join(dirpath, f) for f in emls]


# --------------------------------------------------------------------------
# Driver
# --------------------------------------------------------------------------

def process_mbox_file(up, parts, path):
    folder = up.map_folder(parts)
    if folder is None:
        log("Skipping folder %r (in skip list)" % "/".join(parts))
        return
    log("Folder %r  <-  %s" % (folder, path))
    up.ensure_folder(folder)
    n = 0
    for raw, flags, ident in iter_mbox(path):
        if raw is None:
            up.stats["failed"] += 1
            up.failures.record(folder, ident, "unreadable message")
            continue
        up.upload(folder, raw, flags, message_date(raw), ident)
        n += 1
        if n % 200 == 0:
            log("  ... %d messages processed" % n)
    log("  %d message(s) processed" % n)


def main(argv=None):
    ap = argparse.ArgumentParser(
        description="Restore locally cached mail to a Carbonio/IMAP server.",
        formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--server", required=True, help="IMAP server hostname")
    ap.add_argument("--port", type=int, default=None,
                    help="IMAP port (default 993, or 143 with --no-ssl)")
    ap.add_argument("--user", required=True, help="account, e.g. alice@example.com")
    ap.add_argument("--password", default=None,
                    help="password (default: $IMAP_PASSWORD or interactive prompt)")
    ap.add_argument("--no-ssl", action="store_true",
                    help="plain IMAP instead of IMAPS")
    ap.add_argument("--starttls", action="store_true",
                    help="upgrade a --no-ssl connection with STARTTLS")
    ap.add_argument("--insecure", action="store_true",
                    help="skip TLS certificate verification (self-signed certs)")

    src = ap.add_argument_group("sources (give one or more)")
    src.add_argument("--thunderbird", action="append", default=[],
                     metavar="PROFILE_DIR",
                     help="Thunderbird profile directory (repeatable)")
    src.add_argument("--find-thunderbird", action="store_true",
                     help="auto-detect Thunderbird profiles on this machine")
    src.add_argument("--mbox", action="append", default=[], metavar="PATH",
                     help="an mbox file, or a directory tree of mbox files")
    src.add_argument("--maildir", action="append", default=[], metavar="PATH",
                     help="a Maildir directory")
    src.add_argument("--eml", action="append", default=[], metavar="DIR",
                     help=".eml directory tree (e.g. readpst -r -e output)")

    ap.add_argument("--prefix", default=None,
                    help="restore everything under this folder, e.g. 'Restored'")
    ap.add_argument("--map", action="append", default=[], metavar="SRC=DST",
                    help="extra folder mapping, e.g. 'Old Stuff=Archive/Old'. "
                         "An empty DST skips the folder. Repeatable.")
    ap.add_argument("--dry-run", action="store_true",
                    help="show what would be uploaded without changing anything")
    ap.add_argument("--no-dedupe", action="store_true",
                    help="skip the Message-ID duplicate check")
    ap.add_argument("--dedupe-scope", choices=("folder", "account"),
                    default="folder",
                    help="'folder' (default): skip a message only if it is "
                         "already in the SAME folder on the server. "
                         "'account': index every folder first and never "
                         "upload a message that exists ANYWHERE in the "
                         "mailbox - use this when combining several sources "
                         "(Outlook + Thunderbird + mail already received by "
                         "the new server) that may have filed the same "
                         "message in different folders.")
    ap.add_argument("--max-size", type=int, default=0, metavar="BYTES",
                    help="skip messages larger than this (0 = no limit)")
    ap.add_argument("--throttle", type=float, default=0.0, metavar="SECONDS",
                    help="pause between uploads to be gentle on the server")
    ap.add_argument("--log", default="imap_restore_failures.log",
                    help="file to record failed/skipped messages")
    args = ap.parse_args(argv)

    if args.port is None:
        args.port = 143 if args.no_ssl else 993

    maps = []
    for m in args.map:
        if "=" not in m:
            ap.error("--map must look like SRC=DST (got %r)" % m)
        s, d = m.split("=", 1)
        maps.append((s, d))
    args.map = maps

    if args.find_thunderbird:
        found = find_thunderbird_profiles()
        if not found:
            log("No Thunderbird profiles found on this machine.")
        for p in found:
            log("Found Thunderbird profile: %s" % p)
        args.thunderbird.extend(found)

    args.thunderbird = [os.path.expanduser(p) for p in args.thunderbird]
    if not (args.thunderbird or args.mbox or args.maildir or args.eml):
        ap.error("no sources given: use --thunderbird/--find-thunderbird/"
                 "--mbox/--maildir/--eml")

    if args.password is None:
        args.password = os.environ.get("IMAP_PASSWORD") or \
            getpass.getpass("IMAP password for %s: " % args.user)

    failures = FailureLog(args.log)
    up = Uploader(args, failures)
    up.connect()
    if args.dedupe_scope == "account" and not args.no_dedupe:
        up.scan_account_keys()
    if args.dry_run:
        log("*** DRY RUN - nothing will be uploaded ***")

    try:
        for profile in args.thunderbird:
            log("== Thunderbird profile: %s" % profile)
            roots = thunderbird_roots(profile)
            if not roots:
                log("  no ImapMail/Mail directories found - is this a profile dir?")
            for root in roots:
                log("-- account store: %s" % root)
                for parts, path in scan_thunderbird_root(root):
                    process_mbox_file(up, parts, path)

        for path in args.mbox:
            path = os.path.expanduser(path)
            if os.path.isfile(path):
                name = os.path.splitext(os.path.basename(path))[0]
                process_mbox_file(up, [name], path)
            else:
                for dirpath, _dirs, files in os.walk(path):
                    for fn in sorted(files):
                        p = os.path.join(dirpath, fn)
                        if not looks_like_mbox(p):
                            continue
                        rel = os.path.relpath(p, path)
                        parts = rel.split(os.sep)
                        # readpst -r names every file "mbox"; use the dir name
                        if parts[-1].lower() == "mbox":
                            parts = parts[:-1] or ["INBOX"]
                        else:
                            parts[-1] = os.path.splitext(parts[-1])[0]
                        process_mbox_file(up, parts, p)

        for path in args.maildir:
            path = os.path.expanduser(path)
            folder = up.map_folder([os.path.basename(os.path.normpath(path))])
            if folder is None:
                continue
            log("Folder %r  <-  maildir %s" % (folder, path))
            up.ensure_folder(folder)
            for raw, flags, ident in iter_maildir(path):
                up.upload(folder, raw, flags, message_date(raw), ident)

        for root in args.eml:
            root = os.path.expanduser(root)
            log("== .eml tree: %s" % root)
            for parts, files in iter_eml_tree(root):
                folder = up.map_folder(parts or ["INBOX"])
                if folder is None:
                    log("Skipping folder %r (in skip list)" % "/".join(parts))
                    continue
                log("Folder %r  (%d messages)" % (folder, len(files)))
                up.ensure_folder(folder)
                for f in files:
                    try:
                        with open(f, "rb") as fh:
                            raw = fh.read()
                    except OSError as e:
                        up.stats["failed"] += 1
                        failures.record(folder, f, str(e))
                        continue
                    up.upload(folder, raw, ["\\Seen"], message_date(raw), f)
    except KeyboardInterrupt:
        log("\nInterrupted - it is safe to re-run; already-uploaded messages "
            "will be skipped by the duplicate check.")
    finally:
        try:
            up.imap.logout()
        except Exception:
            pass
        failures.close()

    s = up.stats
    log("\nDone. uploaded=%d  duplicates-skipped=%d  failed=%d  other-skipped=%d"
        % (s["uploaded"], s["skipped_dupe"], s["failed"], s["skipped_other"]))
    if failures.count:
        log("Details of failures/skips are in: %s" % args.log)
    return 1 if s["failed"] else 0


if __name__ == "__main__":
    sys.exit(main())
