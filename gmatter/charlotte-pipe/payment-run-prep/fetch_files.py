#!/usr/bin/env python3
"""
fetch_files.py - replaces the manual FileZilla download + `_charlotte` dedupe + folder placement.

Remote layout:   YYYY/YYYY-MM-Month/{Templates,Unspecified,Credits}
Local layout:    <LOCAL_BASE>/{templates,unspecified,credits}   (plus done/ subfolders)

Rules
  * If a file exists as both  X.xlsx  and  X_charlotte.xlsx  -> download only the _charlotte copy.
  * If only one version exists -> download that one.
  * Files already present locally (in the folder OR its done/ subfolder) are skipped unless --force.

Usage
  export CPF_FTP_PASSWORD='...'            # never put the password in code or the repo
  python fetch_files.py --check            # verify connection + that the month folder exists (no download)
  python fetch_files.py --dry-run          # show what would be downloaded / skipped
  python fetch_files.py                    # download for the current month
  python fetch_files.py --month 2026-09    # a different month
  python fetch_files.py --test             # write to test_templates/test_unspecified/test_credits (safe for trials)
  python fetch_files.py --no-test          # write to the real folders

In VS Code you can skip the flags: edit TEST_MODE / DRY_RUN / CHECK_ONLY / FORCE / MONTH in RUN SETTINGS below.
"""
import argparse
import calendar
import getpass
import os
import re
import socket
import ssl
import sys
from datetime import date
from ftplib import FTP, FTP_TLS, error_perm
from pathlib import Path

# ----------------------------- CONFIG ---------------------------------------
FTP_HOST = os.environ.get("CPF_FTP_HOST", "charlottepipe.sharefileftp.com")
FTP_PORT = int(os.environ.get("CPF_FTP_PORT", "990"))          # implicit FTPS default
FTP_USER = os.environ.get("CPF_FTP_USER", "ud807ab068")
# password: env var CPF_FTP_PASSWORD, else you are prompted

LOCAL_BASE = Path(os.environ.get(
    "CPF_LOCAL_BASE", "/Users/lorimartella/Documents/gmatter/charlotte_pipe"))

# ---- RUN SETTINGS: edit these when running from VS Code (command-line flags override them) ----
TEST_MODE = True     # True  -> save to test_templates / test_unspecified / test_credits
                     # False -> save to the real templates / unspecified / credits folders
DRY_RUN = False      # True  -> only list what would be downloaded, write nothing
CHECK_ONLY = False   # True  -> only test login + month folder, download nothing
FORCE = False        # True  -> re-download files that already exist locally
MONTH = None         # None = current month, or e.g. "2026-09"
# ------------------------------------------------------------------------------------------------

# remote subfolder name (case-insensitive) -> local folder name
FOLDER_MAP = {"templates": "templates", "unspecified": "unspecified", "credits": "credits"}
CHARLOTTE_SUFFIX = "_charlotte"
# ---------------------------------------------------------------------------


class ImplicitFTP_TLS(FTP_TLS):
    """FTP over implicit TLS (TLS from the first byte, port 990).
    Python's FTP_TLS is explicit-only (AUTH TLS), so wrap the socket ourselves."""

    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self._sock = None

    @property
    def sock(self):
        return self._sock

    @sock.setter
    def sock(self, value):
        if value is not None and not isinstance(value, ssl.SSLSocket):
            value = self.context.wrap_socket(value, server_hostname=self.host)
        self._sock = value

    def ntransfercmd(self, cmd, rest=None):
        # data channel: plain connect, then TLS (reusing the control session, which many servers require)
        conn, size = FTP.ntransfercmd(self, cmd, rest)
        if self._prot_p:
            conn = self.context.wrap_socket(
                conn, server_hostname=self.host, session=self.sock.session)
        return conn, size


def connect(host, port, user, password, tls="implicit", timeout=60):
    if tls == "implicit":
        ftp = ImplicitFTP_TLS(timeout=timeout)
        ftp.connect(host, port)
        ftp.login(user, password)
        ftp.prot_p()
    elif tls == "explicit":
        ftp = FTP_TLS(timeout=timeout)
        ftp.connect(host, port)
        ftp.login(user, password)
        ftp.prot_p()
    else:  # "none" - local testing only
        ftp = FTP(timeout=timeout)
        ftp.connect(host, port)
        ftp.login(user, password)
    ftp.set_pasv(True)
    return ftp


def list_dir(ftp, path):
    """Return [(name, is_dir)] for a remote directory. MLSD if supported, else NLST."""
    try:
        out = []
        for name, facts in ftp.mlsd(path):
            if facts.get("type") in ("cdir", "pdir"):
                continue
            out.append((name, facts.get("type") == "dir"))
        return out
    except error_perm:
        pass
    out = []
    for p in ftp.nlst(path):
        name = p.rstrip("/").split("/")[-1]
        if name in (".", ".."):
            continue
        try:  # directory test
            cur = ftp.pwd()
            ftp.cwd(f"{path}/{name}")
            ftp.cwd(cur)
            out.append((name, True))
        except error_perm:
            out.append((name, False))
    return out


def find_child(ftp, parent, wanted, want_dir=True):
    """Case-insensitive lookup of a child name under parent."""
    for name, is_dir in list_dir(ftp, parent):
        if name.lower() == wanted.lower() and is_dir == want_dir:
            return name
    return None


def month_folder_name(ym):
    y, m = ym.split("-")
    return f"{y}-{m}-{calendar.month_name[int(m)]}"          # 2026-10-October


def choose_files(names):
    """Apply the _charlotte rule. Returns (keep, skipped) as lists of (name, reason)."""
    def split(n):
        stem, dot, ext = n.rpartition(".")
        if not dot:
            stem, ext = n, ""
        return stem, ext

    groups = {}
    for n in names:
        stem, ext = split(n)
        is_c = stem.lower().endswith(CHARLOTTE_SUFFIX)
        base = stem[: -len(CHARLOTTE_SUFFIX)] if is_c else stem
        groups.setdefault((base.lower(), ext.lower()), {})["c" if is_c else "o"] = n
    keep, skipped = [], []
    for g in groups.values():
        if "c" in g:
            keep.append(g["c"])
            if "o" in g:
                skipped.append((g["o"], "original - _charlotte version exists"))
        else:
            keep.append(g["o"])
    return sorted(keep), sorted(skipped)


def already_have(local_dir, name):
    return (local_dir / name).exists() or (local_dir / "done" / name).exists()


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    B = argparse.BooleanOptionalAction   # gives --test / --no-test, --dry-run / --no-dry-run, etc.
    ap.add_argument("--month", default=MONTH, help="YYYY-MM (default: MONTH setting, else current month)")
    ap.add_argument("--dry-run", action=B, default=DRY_RUN, help="list what would be downloaded; download nothing")
    ap.add_argument("--check", action=B, default=CHECK_ONLY, help="test login + month folder only")
    ap.add_argument("--force", action=B, default=FORCE, help="re-download files already present locally")
    ap.add_argument("--test", action=B, default=TEST_MODE,
                    help="write to test_templates / test_unspecified / test_credits (use --no-test for real folders)")
    ap.add_argument("--local-base", default=str(LOCAL_BASE))
    ap.add_argument("--host", default=FTP_HOST)
    ap.add_argument("--port", type=int, default=FTP_PORT)
    ap.add_argument("--user", default=FTP_USER)
    ap.add_argument("--tls", choices=["implicit", "explicit", "none"], default="implicit",
                    help="default implicit (matches FileZilla 'Require implicit FTP over TLS')")
    a = ap.parse_args()

    ym = a.month or date.today().strftime("%Y-%m")
    if not re.fullmatch(r"\d{4}-(0[1-9]|1[0-2])", ym):
        sys.exit("--month must look like 2026-10")
    year, mfolder = ym[:4], month_folder_name(ym)
    password = os.environ.get("CPF_FTP_PASSWORD") or getpass.getpass(f"FTP password for {a.user}: ")
    base = Path(a.local_base)

    print(f"MODE: {'TEST folders (test_*)' if a.test else 'REAL folders'}"
          f"{' | DRY RUN' if a.dry_run else ''}{' | CHECK ONLY' if a.check else ''} | month {ym}")
    print(f"Connecting to {a.host}:{a.port} as {a.user} ({a.tls} TLS) ...")
    try:
        ftp = connect(a.host, a.port, a.user, password, a.tls)
    except (OSError, socket.error, ssl.SSLError, error_perm) as e:
        sys.exit(f"CONNECTION FAILED: {e}")
    print("Login OK.")

    try:
        root = ftp.pwd().rstrip("/")
        y = find_child(ftp, root, year)
        if not y:
            sys.exit(f"Remote folder '{year}' not found under {root or '/'}")
        m = find_child(ftp, f"{root}/{y}", mfolder)
        if not m:
            have = [n for n, d in list_dir(ftp, f"{root}/{y}") if d]
            sys.exit(f"Remote folder '{mfolder}' not found in {y}. Found: {sorted(have)}")
        month_path = f"{root}/{y}/{m}"
        print(f"Month folder: {month_path}")

        subs = {n.lower(): n for n, d in list_dir(ftp, month_path) if d}
        for want in FOLDER_MAP:
            if want not in subs:
                print(f"  WARNING: no '{want}' folder in {m}")
        if a.check:
            print("CHECK OK - connection and month folder verified. Nothing downloaded.")
            return

        totals = dict(download=0, have=0, orig_skipped=0, failed=0)
        for want, local_name in FOLDER_MAP.items():
            if want not in subs:
                continue
            rpath = f"{month_path}/{subs[want]}"
            ldir = base / (f"test_{local_name}" if a.test else local_name)
            files = [n for n, d in list_dir(ftp, rpath) if not d and not n.startswith(".")]
            keep, skipped = choose_files(files)
            print(f"\n[{subs[want]}] {len(files)} remote files -> {len(keep)} to use "
                  f"({len(skipped)} originals superseded by _charlotte)")
            for n, why in skipped:
                print(f"    skip   {n}   ({why})")
            totals["orig_skipped"] += len(skipped)
            if not a.dry_run:
                ldir.mkdir(parents=True, exist_ok=True)
            for n in keep:
                if not a.force and already_have(ldir, n):
                    print(f"    have   {n}")
                    totals["have"] += 1
                    continue
                if a.dry_run:
                    print(f"    WOULD DOWNLOAD  {n}")
                    totals["download"] += 1
                    continue
                tmp = ldir / (n + ".part")
                try:
                    with open(tmp, "wb") as fh:
                        ftp.retrbinary(f"RETR {rpath}/{n}", fh.write)
                    tmp.replace(ldir / n)
                    print(f"    got    {n}")
                    totals["download"] += 1
                except Exception as e:  # noqa
                    tmp.unlink(missing_ok=True)
                    print(f"    FAILED {n}: {e}")
                    totals["failed"] += 1

        verb = "would download" if a.dry_run else "downloaded"
        print(f"\nSUMMARY ({ym}): {verb} {totals['download']}, already local {totals['have']}, "
              f"originals skipped {totals['orig_skipped']}, failed {totals['failed']}")
        if totals["failed"]:
            sys.exit(1)
    finally:
        try:
            ftp.quit()
        except Exception:
            pass


if __name__ == "__main__":
    main()