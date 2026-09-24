#!/usr/bin/env python3
"""Filesystem location scraper for GURI — inventory files under configured roots.

Read-only by design: GURI never writes to, renames, or stamps the source files.
A GURI for a file is an *internal database tag* that references the path; the
file on disk is left untouched.
"""

from __future__ import annotations

import json
import os
import re
from dataclasses import asdict, dataclass, field
from datetime import datetime
from pathlib import Path
from typing import Any, Callable, Dict, Iterable, List, Optional, Sequence, Set

from guri_action_detector import extract_deadlines


def _local_geofooter() -> Path:
    local = os.environ.get("LOCALAPPDATA", "")
    if local:
        return Path(local) / "GeoFooter"
    return Path(r"C:\GeoFooter")


def locations_path() -> Path:
    return _local_geofooter() / "guri_scan_locations.json"


def file_cache_path() -> Path:
    return _local_geofooter() / "guri_file_scrape.json"


# Extension → GURI document_type code
EXT_DOCUMENT_TYPES: Dict[str, str] = {
    ".doc": "0a",
    ".docx": "0a",
    ".rtf": "0a",
    ".odt": "0a",
    ".xls": "0b",
    ".xlsx": "0b",
    ".csv": "0b",
    ".ods": "0b",
    ".ppt": "0c",
    ".pptx": "0c",
    ".odp": "0c",
    ".pdf": "0d",
    ".txt": "0e",
    ".md": "0e",
    ".log": "0e",
    ".json": "0e",
    ".xml": "0e",
    ".html": "0e",
    ".htm": "0e",
    ".png": "0f",
    ".jpg": "0f",
    ".jpeg": "0f",
    ".gif": "0f",
    ".bmp": "0f",
    ".webp": "0f",
    ".tif": "0f",
    ".tiff": "0f",
    ".mp4": "10",
    ".mov": "10",
    ".avi": "10",
    ".mkv": "10",
    ".mp3": "11",
    ".wav": "11",
    ".m4a": "11",
    ".zip": "12",
    ".7z": "12",
    ".rar": "12",
    ".gz": "12",
}

# Readable as plain text for deadline cues (kept small)
TEXT_PREVIEW_EXTS: Set[str] = {
    ".txt",
    ".md",
    ".log",
    ".csv",
    ".json",
    ".xml",
    ".html",
    ".htm",
    ".rtf",
}

SKIP_DIR_NAMES: Set[str] = {
    ".git",
    ".svn",
    ".hg",
    "__pycache__",
    "node_modules",
    ".venv",
    "venv",
    ".tox",
    ".mypy_cache",
    ".pytest_cache",
    "$recycle.bin",
    "system volume information",
    "appdata",
}

DEFAULT_MAX_FILES_PER_LOCATION = 5000
DEFAULT_MAX_PREVIEW_BYTES = 80_000
DEFAULT_MAX_FILE_BYTES = 250 * 1024 * 1024  # skip listing content for huge files, still inventory


@dataclass
class ScanLocation:
    path: str
    enabled: bool = True
    recursive: bool = True
    label: str = ""

    def to_dict(self) -> Dict[str, Any]:
        return {
            "path": self.path,
            "enabled": bool(self.enabled),
            "recursive": bool(self.recursive),
            "label": self.label or "",
        }

    @classmethod
    def from_dict(cls, raw: Any) -> Optional["ScanLocation"]:
        if isinstance(raw, str):
            path = raw.strip()
            if not path:
                return None
            return cls(path=path)
        if not isinstance(raw, dict):
            return None
        path = str(raw.get("path") or "").strip()
        if not path:
            return None
        return cls(
            path=path,
            enabled=bool(raw.get("enabled", True)),
            recursive=bool(raw.get("recursive", True)),
            label=str(raw.get("label") or "").strip(),
        )


@dataclass
class ScrapedFile:
    path: str
    name: str
    extension: str
    document_type: str
    size: int
    modified: str
    location_root: str
    relative_path: str
    preview: str = ""
    deadlines: List[Dict[str, Any]] = field(default_factory=list)
    error: str = ""

    def to_dict(self) -> Dict[str, Any]:
        return asdict(self)

    def to_detection_dict(self) -> Dict[str, Any]:
        """Shape compatible with deadline / action UI rows (pseudo-mail)."""
        due_bits = [d for d in (self.deadlines or []) if isinstance(d, dict)]
        subject = self.name
        body = self.preview or self.relative_path
        mail = {
            "entry_id": f"file:{self.path}",
            "store_id": "file",
            "account_smtp": "file",
            "subject": subject,
            "sender": "File scrape",
            "to": "",
            "cc": "",
            "received": self.modified,
            "unread": False,
            "importance": 1,
            "flag_status": 0,
            "message_class": "IPM.File",
            "conversation": self.location_root,
            "body_preview": body[:2500],
            "recipient_role": "to",
            "is_meeting_request": False,
            "file_path": self.path,
            "document_type": self.document_type,
            "relative_path": self.relative_path,
        }
        actions: List[str] = []
        reasons: List[str] = [f"File under {self.location_root}"]
        if due_bits:
            actions.append("deadline")
            reasons.append("Deadline cues in file text")
        return {
            "mail": mail,
            "actions": actions,
            "reasons": reasons,
            "importance": "high" if due_bits else "normal",
            "score": 40 if due_bits else 10,
            "status": "—",
            "deadlines": due_bits,
            "source": "file",
            "file": self.to_dict(),
        }


def load_scan_locations() -> List[ScanLocation]:
    path = locations_path()
    if not path.exists():
        return []
    try:
        data = json.loads(path.read_text(encoding="utf-8"))
    except Exception:
        return []
    raw_list: Any
    if isinstance(data, list):
        raw_list = data
    elif isinstance(data, dict):
        raw_list = data.get("locations") or []
    else:
        return []
    out: List[ScanLocation] = []
    seen: Set[str] = set()
    for item in raw_list or []:
        loc = ScanLocation.from_dict(item)
        if not loc:
            continue
        key = os.path.normcase(os.path.abspath(loc.path))
        if key in seen:
            continue
        seen.add(key)
        out.append(loc)
    return out


def save_scan_locations(locations: Iterable[ScanLocation]) -> Path:
    path = locations_path()
    path.parent.mkdir(parents=True, exist_ok=True)
    cleaned: List[Dict[str, Any]] = []
    seen: Set[str] = set()
    for loc in locations or []:
        if not isinstance(loc, ScanLocation):
            loc = ScanLocation.from_dict(loc)  # type: ignore[arg-type]
            if not loc:
                continue
        key = os.path.normcase(os.path.abspath(loc.path))
        if key in seen:
            continue
        seen.add(key)
        cleaned.append(loc.to_dict())
    payload = {
        "locations": cleaned,
        "updated": datetime.now().isoformat(timespec="seconds"),
    }
    path.write_text(json.dumps(payload, indent=2), encoding="utf-8")
    return path


def load_file_scrape_cache() -> Dict[str, Any]:
    path = file_cache_path()
    if not path.exists():
        return {"items": [], "scraped_at": "", "roots": []}
    try:
        data = json.loads(path.read_text(encoding="utf-8"))
    except Exception:
        return {"items": [], "scraped_at": "", "roots": []}
    if not isinstance(data, dict):
        return {"items": [], "scraped_at": "", "roots": []}
    items = data.get("items") if isinstance(data.get("items"), list) else []
    return {
        "items": items,
        "scraped_at": str(data.get("scraped_at") or ""),
        "roots": list(data.get("roots") or []),
        "file_count": int(data.get("file_count") or len(items)),
        "deadline_count": int(data.get("deadline_count") or 0),
    }


def save_file_scrape_cache(items: Sequence[Dict[str, Any]], *, roots: Sequence[str]) -> Path:
    path = file_cache_path()
    path.parent.mkdir(parents=True, exist_ok=True)
    deadline_count = sum(
        1 for it in items if isinstance(it, dict) and (it.get("deadlines") or [])
    )
    payload = {
        "scraped_at": datetime.now().isoformat(timespec="seconds"),
        "roots": [str(r) for r in roots],
        "file_count": len(items),
        "deadline_count": deadline_count,
        "items": list(items),
    }
    path.write_text(json.dumps(payload, indent=2), encoding="utf-8")
    return path


def document_type_for_path(path: Path) -> str:
    ext = path.suffix.lower()
    return EXT_DOCUMENT_TYPES.get(ext, "99")


def _should_skip_dir(name: str) -> bool:
    return name.lower() in SKIP_DIR_NAMES or name.startswith(".")


def _read_text_preview(path: Path, *, max_bytes: int = DEFAULT_MAX_PREVIEW_BYTES) -> str:
    """Read a small preview for cue detection. Opens read-only; never writes."""
    try:
        with path.open("rb") as fh:
            raw = fh.read(max_bytes)
    except Exception:
        return ""
    for enc in ("utf-8", "utf-16", "cp1252", "latin-1"):
        try:
            text = raw.decode(enc)
            break
        except Exception:
            text = ""
    if not text:
        return ""
    # Strip crude HTML tags for cue matching
    if path.suffix.lower() in {".html", ".htm"}:
        text = re.sub(r"<[^>]+>", " ", text)
    return re.sub(r"\s+", " ", text).strip()


def scrape_file_locations(
    locations: Optional[Sequence[ScanLocation]] = None,
    *,
    max_per_location: int = DEFAULT_MAX_FILES_PER_LOCATION,
    extract_text: bool = True,
    extra_deadline_cues: Optional[Sequence[str]] = None,
    progress: Optional[Callable[[str], None]] = None,
) -> List[ScrapedFile]:
    """Walk enabled locations and return inventory (+ optional deadline cues)."""
    locs = list(locations) if locations is not None else load_scan_locations()
    results: List[ScrapedFile] = []
    seen_files: Set[str] = set()

    def _log(msg: str) -> None:
        if progress:
            try:
                progress(msg)
            except Exception:
                pass

    for loc in locs:
        if not loc.enabled:
            continue
        root = Path(loc.path)
        if not root.exists() or not root.is_dir():
            _log(f"Skip missing: {loc.path}")
            continue
        root_resolved = root.resolve()
        _log(f"Scanning {root_resolved}…")
        count_here = 0

        if loc.recursive:
            walker = os.walk(root_resolved, topdown=True, followlinks=False)
        else:
            # Non-recursive: single directory listing
            try:
                names = [p.name for p in root_resolved.iterdir()]
            except Exception as exc:
                _log(f"Cannot list {root_resolved}: {exc}")
                continue
            walker = [(str(root_resolved), [], names)]

        for dirpath, dirnames, filenames in walker:
            if loc.recursive:
                dirnames[:] = [d for d in dirnames if not _should_skip_dir(d)]
            for fname in filenames:
                if count_here >= max_per_location:
                    _log(f"Hit max {max_per_location} files under {root_resolved}")
                    break
                fpath = Path(dirpath) / fname
                try:
                    if not fpath.is_file():
                        continue
                    key = os.path.normcase(str(fpath.resolve()))
                except Exception:
                    continue
                if key in seen_files:
                    continue
                seen_files.add(key)

                err = ""
                size = 0
                modified = ""
                preview = ""
                deadlines: List[Dict[str, Any]] = []
                try:
                    st = fpath.stat()
                    size = int(st.st_size)
                    modified = datetime.fromtimestamp(st.st_mtime).isoformat(
                        timespec="seconds"
                    )
                except Exception as exc:
                    err = str(exc)

                ext = fpath.suffix.lower()
                doc_type = document_type_for_path(fpath)
                if (
                    extract_text
                    and not err
                    and ext in TEXT_PREVIEW_EXTS
                    and 0 < size <= DEFAULT_MAX_FILE_BYTES
                ):
                    preview = _read_text_preview(fpath)
                    if preview:
                        try:
                            deadlines = extract_deadlines(
                                f"{fname}\n{preview}",
                                extra_cues=extra_deadline_cues,
                            )
                        except Exception:
                            deadlines = []

                try:
                    rel = str(fpath.resolve().relative_to(root_resolved))
                except Exception:
                    rel = fname

                results.append(
                    ScrapedFile(
                        path=str(fpath.resolve()),
                        name=fname,
                        extension=ext,
                        document_type=doc_type,
                        size=size,
                        modified=modified,
                        location_root=str(root_resolved),
                        relative_path=rel.replace("\\", "/"),
                        preview=preview[:2500],
                        deadlines=deadlines,
                        error=err,
                    )
                )
                count_here += 1
            if count_here >= max_per_location:
                break

    results.sort(key=lambda f: (f.location_root.lower(), f.relative_path.lower()))
    _log(f"Done — {len(results)} file(s)")
    return results


def scraped_files_to_detection_items(
    files: Sequence[ScrapedFile],
) -> List[Dict[str, Any]]:
    return [f.to_detection_dict() for f in files]
