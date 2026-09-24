"""Match mail senders against the broker catalog."""

from __future__ import annotations

from typing import Any, Dict, Optional

from aura.catalog import match_domain


def match_sender(
    sender_email: Optional[str] = None,
    sender_domain: Optional[str] = None,
) -> Optional[Dict[str, Any]]:
    """Return catalog broker if sender email or domain matches a known broker."""
    email = (sender_email or "").strip().lower()
    domain = (sender_domain or "").strip().lower()
    if domain in {"", "unknown"}:
        domain = ""
    hit = match_domain(email) if email else None
    if hit:
        return hit
    if domain:
        return match_domain(domain)
    return None
