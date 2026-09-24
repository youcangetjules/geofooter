"""AES → GURI pending queue for broker-removal footer actions."""

from __future__ import annotations

import json
import os
import uuid
from datetime import datetime, timezone
from pathlib import Path
from typing import Any, Dict, List, Optional

from aura.catalog import find_broker, match_domain
from aura.settings import load_settings
from aura.store import RemovalStore


def data_dir() -> Path:
    local = os.environ.get("LOCALAPPDATA", "")
    return Path(local) / "GeoFooter" if local else Path(r"C:\GeoFooter")


def pending_path() -> Path:
    return data_dir() / "broker_pending_from_aes.json"


def _now_iso() -> str:
    return datetime.now(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ")


def _load() -> List[Dict[str, Any]]:
    path = pending_path()
    if not path.is_file():
        return []
    try:
        data = json.loads(path.read_text(encoding="utf-8"))
        items = data.get("items") if isinstance(data, dict) else data
        if isinstance(items, list):
            return [x for x in items if isinstance(x, dict)]
    except Exception:
        pass
    return []


def _save(items: List[Dict[str, Any]]) -> None:
    path = pending_path()
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(
        json.dumps({"items": items}, indent=2, ensure_ascii=False),
        encoding="utf-8",
    )


def _parse_when(value: str) -> Optional[datetime]:
    try:
        return datetime.fromisoformat((value or "").replace("Z", "+00:00"))
    except Exception:
        return None


def _in_window(item: Dict[str, Any]) -> bool:
    """Items without received_at predate the date gate (manual footer clicks) and are kept."""
    raw = str(item.get("received_at") or "")
    if not raw:
        return True
    return load_settings().is_in_window(_parse_when(raw))


def list_pending() -> List[Dict[str, Any]]:
    return [x for x in _load() if _in_window(x)]


def enqueue_from_aes(
    *,
    sender: str = "",
    domain: str = "",
    subject: str = "",
    guri: str = "",
    broker_id: str = "",
    received_at: str = "",
) -> Optional[Dict[str, Any]]:
    """Append a pending AES queue item (dedupe by domain+sender).

    ``received_at`` (ISO) is set by automatic scan detection; mail dated before
    ``detect_since`` is ignored and None is returned. Manual footer clicks pass
    no date and are always queued.
    """
    received_at = (received_at or "").strip()
    if received_at and not load_settings().is_in_window(_parse_when(received_at)):
        return None
    sender = (sender or "").strip().lower()
    domain = (domain or "").strip().lower()
    if domain in {"unknown"}:
        domain = ""
    if not domain and "@" in sender:
        domain = sender.rsplit("@", 1)[-1]

    broker = find_broker(broker_id) if broker_id else None
    if not broker:
        broker = match_domain(domain or sender)
    broker_id = str((broker or {}).get("id") or broker_id or "unknown")
    broker_name = str((broker or {}).get("name") or domain or sender or "Unknown broker")

    items = _load()
    for existing in items:
        if (
            str(existing.get("sender") or "").lower() == sender
            and str(existing.get("domain") or "").lower() == domain
            and str(existing.get("broker_id") or "") == broker_id
        ):
            existing["subject"] = subject or existing.get("subject") or ""
            existing["guri"] = guri or existing.get("guri") or ""
            existing["received_at"] = received_at or existing.get("received_at") or ""
            existing["scanned_at"] = _now_iso()
            _save(items)
            return existing

    item = {
        "id": uuid.uuid4().hex[:12],
        "broker_id": broker_id,
        "broker_name": broker_name,
        "sender": sender,
        "domain": domain,
        "subject": (subject or "").strip(),
        "guri": (guri or "").strip(),
        "received_at": received_at,
        "scanned_at": _now_iso(),
    }
    items.insert(0, item)
    _save(items[:200])
    return item


def dismiss_pending(pending_id: str) -> bool:
    items = _load()
    new = [x for x in items if str(x.get("id")) != pending_id]
    if len(new) == len(items):
        return False
    _save(new)
    return True


def accept_pending(pending_id: str, store: Optional[RemovalStore] = None):
    """Create a draft RemovalRequest from a pending AES item and remove it."""
    items = _load()
    match = None
    for x in items:
        if str(x.get("id")) == pending_id:
            match = x
            break
    if not match:
        return None
    store = store or RemovalStore()
    req = store.create(
        broker_id=str(match.get("broker_id") or "unknown"),
        broker_name=str(match.get("broker_name") or "Unknown"),
        status="draft",
        source="aes",
        sender=str(match.get("sender") or ""),
        domain=str(match.get("domain") or ""),
        subject=str(match.get("subject") or ""),
        guri=str(match.get("guri") or ""),
        notes="Queued from AES footer",
    )
    _save([x for x in items if str(x.get("id")) != pending_id])
    return req
