#!/usr/bin/env python3
"""GeoFooter install-root and relative path helpers.

The install root is configurable (GURI → Database → Install root). A pointer
file always lives at ``%LOCALAPPDATA%\\GeoFooter\\install_root.txt`` so VBA and
Python can resolve paths without hard-coding ``C:\\GeoFooter``.

Layout under the install root (relative):
  assets/   brand, icons, pages
  datastore/   local SQLite DBs
  debuglog/    operational debug logs
  crashlogs/   unhandled exception dumps
  VBA/      Outlook modules + scanners
  scripts/  maintenance / launchers
  docs/     guides

Per-user runtime state (settings JSON, secrets, job queues) stays under
``%LOCALAPPDATA%\\GeoFooter\\`` and is independent of the install root.
"""

from __future__ import annotations

import json
import os
from pathlib import Path
from typing import Optional, Union

PathLike = Union[str, Path]

_POINTER_NAME = "install_root.txt"
_POINTER_JSON = "geofooter_paths.json"
_ENV_VAR = "GEOFOOTER_ROOT"

# Cached after first resolve (call clear_cache() after set_install_root).
_cached_root: Optional[Path] = None


def user_data_dir() -> Path:
    """Per-user GeoFooter state (%LOCALAPPDATA%\\GeoFooter)."""
    local = (os.environ.get("LOCALAPPDATA") or "").strip()
    if local:
        p = Path(local) / "GeoFooter"
    else:
        p = Path(os.environ.get("TEMP") or ".") / "GeoFooter"
    p.mkdir(parents=True, exist_ok=True)
    return p


def _pointer_file() -> Path:
    return user_data_dir() / _POINTER_NAME


def _pointer_json() -> Path:
    return user_data_dir() / _POINTER_JSON


def clear_cache() -> None:
    global _cached_root
    _cached_root = None


def _discover_from_file() -> Optional[Path]:
    """Walk parents of this module looking for a suite root marker."""
    here = Path(__file__).resolve()
    for parent in [here.parent, *here.parents]:
        if (parent / "VERSION").is_file() and (
            (parent / "guri_gui.py").is_file()
            or (parent / "VBA" / "geolocate_headers.py").is_file()
        ):
            return parent
    return None


def get_install_root(*, refresh: bool = False) -> Path:
    """Return the configured GeoFooter install root (absolute)."""
    global _cached_root
    if _cached_root is not None and not refresh:
        return _cached_root

    candidates: list[Path] = []

    env = (os.environ.get(_ENV_VAR) or "").strip().strip('"')
    if env:
        candidates.append(Path(env))

    ptr = _pointer_file()
    if ptr.is_file():
        try:
            line = ptr.read_text(encoding="utf-8", errors="replace").strip().splitlines()
            if line and line[0].strip():
                candidates.append(Path(line[0].strip().strip('"')))
        except Exception:
            pass

    jpath = _pointer_json()
    if jpath.is_file():
        try:
            data = json.loads(jpath.read_text(encoding="utf-8"))
            raw = str((data or {}).get("install_root") or "").strip()
            if raw:
                candidates.append(Path(raw))
        except Exception:
            pass

    discovered = _discover_from_file()
    if discovered is not None:
        candidates.append(discovered)

    # Legacy default if the classic folder still exists.
    legacy = Path(r"C:\GeoFooter")
    if legacy.is_dir():
        candidates.append(legacy)

    for c in candidates:
        try:
            resolved = c.expanduser().resolve()
            if resolved.is_dir():
                _cached_root = resolved
                return resolved
        except Exception:
            continue

    # Last resort: create/use discovered parent or cwd/GeoFooter
    fallback = discovered or (Path.cwd() / "GeoFooter")
    try:
        fallback.mkdir(parents=True, exist_ok=True)
    except Exception:
        fallback = Path.cwd()
    _cached_root = fallback.resolve()
    return _cached_root


def set_install_root(path: PathLike) -> Path:
    """Persist install root for Python + VBA and return the absolute path."""
    root = Path(path).expanduser().resolve()
    root.mkdir(parents=True, exist_ok=True)
    for sub in ("datastore", "debuglog", "crashlogs", "assets", "output"):
        (root / sub).mkdir(parents=True, exist_ok=True)

    ptr = _pointer_file()
    ptr.write_text(str(root) + "\n", encoding="utf-8")
    _pointer_json().write_text(
        json.dumps({"install_root": str(root)}, indent=2) + "\n",
        encoding="utf-8",
    )
    os.environ[_ENV_VAR] = str(root)
    clear_cache()
    return get_install_root(refresh=True)


def rel(*parts: str) -> Path:
    """Join *parts* under the install root."""
    return get_install_root().joinpath(*parts)


def datastore_dir() -> Path:
    p = rel("datastore")
    p.mkdir(parents=True, exist_ok=True)
    return p


def debuglog_dir() -> Path:
    p = rel("debuglog")
    p.mkdir(parents=True, exist_ok=True)
    return p


def crashlogs_dir() -> Path:
    p = rel("crashlogs")
    p.mkdir(parents=True, exist_ok=True)
    return p


def assets_dir() -> Path:
    return rel("assets")


def icons_dir() -> Path:
    return rel("assets", "icons")


def vba_dir() -> Path:
    return rel("VBA")


def scripts_dir() -> Path:
    return rel("scripts")


def datastore_path(name: str) -> Path:
    return datastore_dir() / name


def debug_log_path(name: str) -> Path:
    return debuglog_dir() / name


def brand_path(*parts: str) -> Path:
    return rel("assets", "brand", *parts)
