"""
Detect whether a .pptx file is currently held open by another process.

Why this exists
---------------
PowerPoint, Keynote, and LibreOffice all keep an open file handle on a deck
while it's being viewed/edited. Writing the same path from a script during
that window is unsafe:

  * On Windows, the OS will refuse the write outright.
  * On macOS and Linux, the write may succeed but the application is holding
    a stale in-memory copy; the next "Save" inside that app will overwrite
    the script's output silently.

Either way the user ends up with a confusing "where did my change go?" bug.
This module surfaces the conflict before we touch the file.

Detection strategy (best-effort, cross-platform, no external deps)
------------------------------------------------------------------
1. **PowerPoint owner file** — When PowerPoint opens "deck.pptx", it creates a
   sibling lockfile named "~$deck.pptx" in the same directory. Its presence
   is a strong signal PowerPoint has the file open. (Word and Excel use the
   same convention.)

2. **`lsof` probe (macOS / Linux)** — If `lsof` is on PATH, query which
   processes have an open handle on the file. Returns rich info: PID,
   command name, user. POSIX only; quietly skipped on Windows or when lsof
   is missing.

3. **Windows file-locking probe** — On Windows we attempt to open the file
   with exclusive write access; if that fails with PermissionError /
   OSError(EACCES), something else holds it. We can't enumerate the holder
   without external deps, but we can at least block the write.

The function returns a list of "lock holder" dicts. An empty list means
"as far as we can tell, nothing else is holding the file"; it is NOT a
guarantee — file locks are inherently racy.
"""

from __future__ import annotations

import errno
import os
import platform
import shutil
import subprocess
import sys


def _check_owner_file(path: str) -> list[dict]:
    """Look for the PowerPoint/Office owner-lock sibling file."""
    directory, filename = os.path.split(os.path.abspath(path))
    owner_path = os.path.join(directory, "~$" + filename)
    if os.path.exists(owner_path):
        return [{
            "kind": "office_owner_file",
            "path": owner_path,
            "hint": ("This is the lockfile PowerPoint / Word / Excel create "
                     "while a document is open. Close the deck in that app "
                     "to release it."),
        }]
    return []


def _check_lsof(path: str) -> list[dict]:
    """Use `lsof` to find processes holding the file. POSIX only."""
    if platform.system() == "Windows":
        return []
    if shutil.which("lsof") is None:
        return []
    try:
        # -F pcuLn => null-delimited fields: pid, command, user, login, name
        # We only need pid, command, user.
        proc = subprocess.run(
            ["lsof", "-F", "pcun", "--", os.path.abspath(path)],
            capture_output=True, text=True, timeout=3,
        )
    except (subprocess.TimeoutExpired, FileNotFoundError, OSError):
        return []

    if proc.returncode not in (0, 1):
        # 1 = "no processes hold the file" on most lsofs; >1 = real error
        return []

    holders: list[dict] = []
    current: dict = {}
    for raw in proc.stdout.splitlines():
        if not raw:
            continue
        tag, val = raw[0], raw[1:]
        if tag == "p":
            if current:
                holders.append(current)
            current = {"kind": "open_handle", "pid": val}
        elif tag == "c":
            current["command"] = val
        elif tag == "u":
            current["user"] = val
        elif tag == "n":
            current["path"] = val
    if current:
        holders.append(current)

    self_pid = str(os.getpid())
    holders = [h for h in holders if h.get("pid") != self_pid]

    for h in holders:
        cmd = (h.get("command") or "").lower()
        if any(needle in cmd for needle in (
                "powerpnt", "powerpoint", "keynote", "soffice",
                "libreoffice", "openoffice", "wpsoffice")):
            h["hint"] = (
                f"Close '{h.get('command', 'the application')}' (PID "
                f"{h.get('pid')}) — it has the deck open and will "
                f"overwrite our changes when next saved.")
        else:
            h["hint"] = (
                f"Process '{h.get('command', '?')}' (PID {h.get('pid')}) has "
                f"the deck open. Close it before writing.")
    return holders


def _check_windows_write_lock(path: str) -> list[dict]:
    """On Windows, probe for an exclusive write lock by trying to open."""
    if platform.system() != "Windows":
        return []
    if not os.path.exists(path):
        return []
    try:
        with open(path, "r+b"):
            pass
    except PermissionError as e:
        return [{
            "kind": "windows_write_lock",
            "error": str(e),
            "hint": ("Another process (likely PowerPoint) holds an exclusive "
                     "lock on this file. Close that app and retry."),
        }]
    except OSError as e:
        if e.errno in (errno.EACCES, errno.EBUSY):
            return [{
                "kind": "windows_write_lock",
                "error": str(e),
                "hint": ("Another process holds the file. Close it and "
                         "retry."),
            }]
        return []
    return []


def detect_holders(path: str) -> list[dict]:
    """Return a list of plausible lock holders for *path*.

    Each holder is a dict with at minimum a 'kind' and 'hint'. May also
    include 'pid', 'command', 'user', 'path', 'error' depending on source.
    Empty list means "no conflict detected" (not a guarantee).
    """
    if not os.path.exists(path):
        return []
    holders: list[dict] = []
    holders.extend(_check_owner_file(path))
    holders.extend(_check_lsof(path))
    holders.extend(_check_windows_write_lock(path))
    return holders


def format_holders(path: str, holders: list[dict]) -> str:
    """Human-readable error block for stderr."""
    lines = [
        f"[FAIL] '{path}' appears to be open in another application.",
        "       Writing now would either fail or be silently overwritten",
        "       the next time that app saves.",
        "",
    ]
    for i, h in enumerate(holders, 1):
        kind = h.get("kind", "?")
        if kind == "office_owner_file":
            lines.append(
                f"  [{i}] PowerPoint / Office lockfile present at "
                f"{h.get('path')}")
        elif kind == "open_handle":
            lines.append(
                f"  [{i}] PID {h.get('pid')} ({h.get('command', '?')}, "
                f"user {h.get('user', '?')}) has an open handle")
        elif kind == "windows_write_lock":
            lines.append(
                f"  [{i}] Windows reports the file is exclusively held: "
                f"{h.get('error')}")
        else:
            lines.append(f"  [{i}] {kind}: {h}")
        if h.get("hint"):
            lines.append(f"      hint: {h['hint']}")
    lines.extend([
        "",
        "Action: close the deck in PowerPoint / Keynote / LibreOffice and",
        "rerun. Pass --force to bypass this check (only do this if you",
        "are absolutely certain nothing else has the file open).",
    ])
    return "\n".join(lines)


def assert_writable(path: str, *, force: bool = False,
                    label: str = "deck") -> None:
    """Raise SystemExit(2) if *path* is held open by something else.

    No-op when *force* is True or when *path* doesn't yet exist (writing a
    fresh file can't conflict with an existing handle).
    """
    if force:
        return
    holders = detect_holders(path)
    if not holders:
        return
    sys.stderr.write(format_holders(path, holders) + "\n")
    raise SystemExit(2)


__all__ = ["detect_holders", "format_holders", "assert_writable"]
