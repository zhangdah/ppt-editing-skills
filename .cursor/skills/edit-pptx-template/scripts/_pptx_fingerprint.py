"""
Compute and compare fingerprints of a .pptx for staleness detection.

Why this exists
---------------
Audit JSON is a one-shot snapshot. If the user opens the deck in PowerPoint,
edits something, and saves — the audit on disk no longer matches the file.
Patches drafted against the stale audit may target text that has moved,
been deleted, or been replaced.

We avoid that by stamping every audit with a fingerprint of the source
.pptx, and comparing it against the live file before any subsequent step.
The fingerprint is cheap to compute and cheap to compare:

    {
      "size": 31415,                  # bytes on disk
      "mtime_ns": 1714201234567890,   # filesystem mtime, ns precision
      "sha256": "abc..."              # full content hash, definitive
    }

Two-stage compare: cheap fields first (size, mtime), then sha256 only when
those agree. mtime alone is unreliable (PowerPoint may preserve mtime, or
the user may "touch" the file); sha256 alone is slow on big decks. Together
they're both fast and correct.

A fingerprint mismatch never raises by itself — it's just data. The caller
decides how to react (re-audit, prompt, abort).
"""

from __future__ import annotations

import hashlib
import os


def _sha256_of_file(path: str, chunk: int = 1 << 16) -> str:
    h = hashlib.sha256()
    with open(path, "rb") as f:
        while True:
            buf = f.read(chunk)
            if not buf:
                break
            h.update(buf)
    return h.hexdigest()


def fingerprint(path: str) -> dict:
    """Return a fingerprint dict for *path*. Raises FileNotFoundError if missing."""
    st = os.stat(path)
    return {
        "size": st.st_size,
        "mtime_ns": st.st_mtime_ns,
        "sha256": _sha256_of_file(path),
    }


def quick_match(path: str, prior: dict) -> bool | None:
    """Compare cheap fields only. Returns:

      True  — almost certainly unchanged (size + mtime match)
      False — definitely changed (size or mtime differs)
      None  — can't tell from cheap fields alone
    """
    try:
        st = os.stat(path)
    except FileNotFoundError:
        return False
    if prior.get("size") != st.st_size:
        return False
    if prior.get("mtime_ns") != st.st_mtime_ns:
        return False
    return None


def is_stale(path: str, prior: dict) -> tuple[bool, str]:
    """Decide whether *path* still matches the *prior* fingerprint.

    Returns (stale, reason). `stale=False` means content is identical to
    the snapshot; `stale=True` means it has changed (or vanished).
    """
    if not os.path.exists(path):
        return True, "deck no longer exists at the recorded path"

    cheap = quick_match(path, prior)
    if cheap is False:
        st = os.stat(path)
        if prior.get("size") != st.st_size:
            return True, (
                f"size changed: {prior.get('size')} -> {st.st_size} bytes")
        return True, (
            f"mtime changed: {prior.get('mtime_ns')} -> {st.st_mtime_ns}")

    # Cheap fields agree — confirm with sha256 in case mtime was preserved
    # across a save (some apps do this).
    live_sha = _sha256_of_file(path)
    if prior.get("sha256") != live_sha:
        return True, (
            f"content hash changed despite identical size+mtime: "
            f"{prior.get('sha256', '')[:12]}... -> {live_sha[:12]}...")
    return False, "fingerprint matches; deck unchanged since audit"


__all__ = ["fingerprint", "quick_match", "is_stale"]
