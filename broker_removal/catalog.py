"""Broker catalog load + domain match."""

from __future__ import annotations

import json
from functools import lru_cache
from pathlib import Path
from typing import Any, Dict, List, Optional

_CATALOG_PATH = Path(__file__).resolve().parent / "catalog.json"


def catalog_path() -> Path:
    return _CATALOG_PATH


@lru_cache(maxsize=1)
def load_catalog() -> List[Dict[str, Any]]:
    path = catalog_path()
    if not path.is_file():
        return []
    try:
        data = json.loads(path.read_text(encoding="utf-8"))
    except Exception:
        return []
    brokers = data.get("brokers") if isinstance(data, dict) else data
    if not isinstance(brokers, list):
        return []
    out: List[Dict[str, Any]] = []
    for row in brokers:
        if not isinstance(row, dict):
            continue
        bid = str(row.get("id") or "").strip()
        name = str(row.get("name") or "").strip()
        if not bid or not name:
            continue
        domains = [
            str(d).strip().lower().rstrip(".")
            for d in (row.get("domains") or [])
            if str(d).strip()
        ]
        out.append(
            {
                "id": bid,
                "name": name,
                "domains": domains,
                "method": str(row.get("method") or "unknown").strip().lower(),
                "opt_out_url": str(row.get("opt_out_url") or "").strip(),
                "opt_out_email": str(row.get("opt_out_email") or "").strip(),
                "jurisdiction": str(row.get("jurisdiction") or "").strip(),
                "template_id": str(row.get("template_id") or "generic").strip(),
                "notes": str(row.get("notes") or "").strip(),
            }
        )
    return out


def reload_catalog() -> List[Dict[str, Any]]:
    load_catalog.cache_clear()
    return load_catalog()


def find_broker(broker_id: str) -> Optional[Dict[str, Any]]:
    bid = (broker_id or "").strip().lower()
    if not bid:
        return None
    for b in load_catalog():
        if b["id"].lower() == bid:
            return b
    return None


def _domain_matches(host: str, domain: str) -> bool:
    host = (host or "").strip().lower().rstrip(".")
    domain = (domain or "").strip().lower().rstrip(".")
    return bool(domain) and (host == domain or host.endswith("." + domain))


def match_domain(domain_or_email: str) -> Optional[Dict[str, Any]]:
    """Return the first catalog broker whose domain matches the host."""
    raw = (domain_or_email or "").strip().lower()
    if not raw:
        return None
    host = raw.rsplit("@", 1)[-1].strip().rstrip(".")
    if host.startswith("www."):
        host = host[4:]
    for b in load_catalog():
        for d in b.get("domains") or []:
            if _domain_matches(host, d):
                return b
    return None
