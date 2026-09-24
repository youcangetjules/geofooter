#!/usr/bin/env python3
"""Crash dumps and debug-log paths for the GeoFooter suite.

- Crash / unhandled exceptions → ``<install_root>/crashlogs/``
- Operational debug logs       → ``<install_root>/debuglog/``
- Databases                    → ``<install_root>/datastore/``

Install root is configured via ``geofooter_paths`` (GURI Database tab).
"""

from __future__ import annotations

import os
import sys
import traceback
from datetime import datetime, timezone
from pathlib import Path
from typing import Optional

try:
    from geofooter.paths import (
        crashlogs_dir,
        datastore_path,
        debug_log_path,
        debuglog_dir,
        get_install_root,
    )
except ImportError:  # pragma: no cover — bootstrap before paths module exists
    def get_install_root():
        return Path(r"C:\GeoFooter")

    def crashlogs_dir():
        p = get_install_root() / "crashlogs"
        p.mkdir(parents=True, exist_ok=True)
        return p

    def debuglog_dir():
        p = get_install_root() / "debuglog"
        p.mkdir(parents=True, exist_ok=True)
        return p

    def debug_log_path(name: str):
        return debuglog_dir() / name

    def datastore_path(name: str):
        p = get_install_root() / "datastore"
        p.mkdir(parents=True, exist_ok=True)
        return p / name


def ensure_dirs() -> Path:
    crashlogs_dir()
    debuglog_dir()
    return crashlogs_dir()


def crashlog_path(prefix: str = "crash") -> Path:
    """Return a new timestamped crash log path under crashlogs/."""
    ensure_dirs()
    stamp = datetime.now(timezone.utc).strftime("%Y%m%d_%H%M%S")
    safe = "".join(c if c.isalnum() or c in "-_" else "_" for c in (prefix or "crash"))
    return crashlogs_dir() / f"{safe}_{stamp}.log"


def write_crash(
    text: str,
    *,
    prefix: str = "crash",
    exc: Optional[BaseException] = None,
) -> Path:
    """Write *text* (and optional exception traceback) to crashlogs/. Returns path."""
    path = crashlog_path(prefix)
    parts = [text.rstrip(), ""]
    if exc is not None:
        parts.append("".join(traceback.format_exception(type(exc), exc, exc.__traceback__)))
    path.write_text("\n".join(parts).rstrip() + "\n", encoding="utf-8", errors="replace")
    return path


def install_sys_excepthook(prefix: str = "crash") -> None:
    """Install a process-wide hook that dumps unhandled exceptions to crashlogs/."""
    ensure_dirs()
    previous = sys.excepthook

    def _hook(exc_type, exc, tb) -> None:
        try:
            body = (
                f"Unhandled exception in process {os.getpid()}\n"
                f"argv: {sys.argv!r}\n"
                f"cwd: {os.getcwd()}\n"
                f"install_root: {get_install_root()}\n"
            )
            path = write_crash(body, prefix=prefix, exc=exc if isinstance(exc, BaseException) else None)
            if not isinstance(exc, BaseException):
                path.write_text(
                    path.read_text(encoding="utf-8", errors="replace")
                    + "".join(traceback.format_exception(exc_type, exc, tb)),
                    encoding="utf-8",
                    errors="replace",
                )
            try:
                sys.stderr.write(f"[GeoFooter] crash dump written: {path}\n")
            except Exception:
                pass
        except Exception:
            pass
        previous(exc_type, exc, tb)

    sys.excepthook = _hook


def operational_log_path(name: str) -> Path:
    """Path under debuglog/ for non-crash operational/debug logs."""
    return debug_log_path(name)
