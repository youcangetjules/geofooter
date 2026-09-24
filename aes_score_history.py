"""Read and edit AES per-sender score history (``aes_sender_history.json``).

The scan engine (``VBA\\geolocate_headers.py``) keeps a smoothed score per
sender (``score_ema``) that it blends with each new scan. These helpers let a
user inspect, override or clear that memory - e.g. after a scoring bug or
while tuning - using the same lock-file protocol the scanner uses, so edits
never interleave with a scan's read-modify-write.
"""

from __future__ import annotations

import json
import os
import time
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
from typing import Any, Dict, Iterable, List, Optional
from urllib.parse import urlsplit

_LOCK_STALE_SEC = 30


def data_dir() -> Path:
    local = os.environ.get("LOCALAPPDATA", "")
    return Path(local) / "GeoFooter" if local else Path(r"C:\GeoFooter")


def history_path() -> Path:
    return data_dir() / "aes_sender_history.json"


def threat_intel_cache_path() -> Path:
    return data_dir() / "threat_intel_cache.json"


class HistoryLockedError(RuntimeError):
    """A scan is holding the sender history lock."""


@dataclass
class SenderScore:
    email: str
    domain: str
    count: int
    first_seen: str
    last_seen: str
    score: Optional[int]


def sender_domain(email: str) -> str:
    return email.rsplit("@", 1)[-1].strip().lower() if "@" in email else ""


def _domain_matches(host: str, domain: str) -> bool:
    host = (host or "").strip().lower().rstrip(".")
    domain = (domain or "").strip().lower().rstrip(".")
    return bool(domain) and (host == domain or host.endswith("." + domain))


def _acquire_lock(timeout: float = 5.0) -> int:
    lock_path = str(history_path()) + ".lock"
    data_dir().mkdir(parents=True, exist_ok=True)
    deadline = time.time() + timeout
    while True:
        try:
            return os.open(lock_path, os.O_CREAT | os.O_EXCL | os.O_RDWR)
        except FileExistsError:
            try:
                if time.time() - os.path.getmtime(lock_path) > _LOCK_STALE_SEC:
                    os.unlink(lock_path)
                    continue
            except OSError:
                pass
            if time.time() >= deadline:
                raise HistoryLockedError(
                    "Sender history is busy (a scan is running) - try again shortly."
                )
            time.sleep(0.05)


def _release_lock(fd: int) -> None:
    try:
        os.close(fd)
    except OSError:
        pass
    try:
        os.unlink(str(history_path()) + ".lock")
    except OSError:
        pass


def _read() -> Dict[str, Any]:
    path = history_path()
    if not path.is_file():
        return {}
    data = json.loads(path.read_text(encoding="utf-8"))
    return data if isinstance(data, dict) else {}


def _write(data: Dict[str, Any]) -> None:
    path = history_path()
    path.parent.mkdir(parents=True, exist_ok=True)
    data["updated"] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    tmp = Path(str(path) + ".tmp")
    tmp.write_text(json.dumps(data), encoding="utf-8")
    os.replace(str(tmp), str(path))


def _as_score(value: Any) -> Optional[int]:
    if value is None:
        return None
    try:
        return max(0, min(100, int(value)))
    except (TypeError, ValueError):
        return None


def load_sender_scores() -> List[SenderScore]:
    """All senders on file, highest stored score first."""
    rows: List[SenderScore] = []
    for email, entry in (_read().get("senders") or {}).items():
        if not isinstance(entry, dict):
            continue
        try:
            count = int(entry.get("count") or 0)
        except (TypeError, ValueError):
            count = 0
        rows.append(
            SenderScore(
                email=email,
                domain=sender_domain(email),
                count=count,
                first_seen=str(entry.get("first_seen") or ""),
                last_seen=str(entry.get("last_seen") or ""),
                score=_as_score(entry.get("score_ema")),
            )
        )
    rows.sort(key=lambda r: (-(r.score if r.score is not None else -1), r.email))
    return rows


def _mutate(emails: Iterable[str], fn) -> int:
    wanted = {e.strip().lower() for e in emails if e and e.strip()}
    if not wanted:
        return 0
    fd = _acquire_lock()
    try:
        data = _read()
        senders = data.setdefault("senders", {})
        changed = 0
        for email in list(senders):
            if email in wanted and isinstance(senders[email], dict):
                if fn(senders, email):
                    changed += 1
        if changed:
            _write(data)
        return changed
    finally:
        _release_lock(fd)


def set_scores(emails: Iterable[str], score: int) -> int:
    """Override the stored score; the next scan smooths from this value."""
    value = max(0, min(100, int(score)))

    def apply(senders: Dict[str, Any], email: str) -> bool:
        senders[email]["score_ema"] = value
        return True

    return _mutate(emails, apply)


def reset_scores(emails: Iterable[str]) -> int:
    """Forget the stored score; the next scan shows its raw score unsmoothed."""

    def apply(senders: Dict[str, Any], email: str) -> bool:
        return senders[email].pop("score_ema", None) is not None

    return _mutate(emails, apply)


def forget_senders(emails: Iterable[str]) -> int:
    """Delete senders entirely (email count and score); they become unknown again."""

    def apply(senders: Dict[str, Any], email: str) -> bool:
        senders.pop(email, None)
        return True

    return _mutate(emails, apply)


def senders_at_domain(domain: str) -> List[str]:
    return [r.email for r in load_sender_scores() if _domain_matches(r.domain, domain)]


def clear_threat_intel_for_domain(domain: str) -> int:
    """Drop cached threat-intel verdicts for a domain (and its subdomains/URLs).

    Scans merge only their own new results into the cache, so removed entries
    stay gone until a scan looks the indicator up again.
    """
    path = threat_intel_cache_path()
    if not domain or not path.is_file():
        return 0
    data = json.loads(path.read_text(encoding="utf-8"))
    results = data.get("results") if isinstance(data, dict) else None
    if not isinstance(results, dict):
        return 0
    doomed = []
    for key in results:
        parts = key.split("|", 2)
        if len(parts) != 3:
            continue
        _provider, kind, indicator = parts
        if kind == "url":
            try:
                host = urlsplit(indicator).hostname or ""
            except ValueError:
                host = ""
        else:
            host = indicator
        if _domain_matches(host, domain):
            doomed.append(key)
    for key in doomed:
        results.pop(key, None)
    if doomed:
        tmp = path.with_suffix(f".{os.getpid()}.tmp")
        tmp.write_text(json.dumps(data), encoding="utf-8")
        os.replace(tmp, path)
    return len(doomed)
