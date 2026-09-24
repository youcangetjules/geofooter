"""Data-broker removal (Incogni-style) — catalog, identity profile, request tracking.

Persistence lives under ``%LOCALAPPDATA%\\GeoFooter\\`` so personal identity PII
is not mixed into GURI Postgres document storage.

v1: templates + tracking only — no web-form automation, no Outlook auto-send.
"""

from __future__ import annotations

from aura.catalog import load_catalog, find_broker, match_domain
from aura.detect import match_sender
from aura.pending import (
    accept_pending,
    dismiss_pending,
    enqueue_from_aes,
    list_pending,
)
from aura.profile import IdentityProfile, load_profile, save_profile
from aura.store import RemovalRequest, RemovalStore
from aura.templates import render_removal

__all__ = [
    "IdentityProfile",
    "RemovalRequest",
    "RemovalStore",
    "accept_pending",
    "dismiss_pending",
    "enqueue_from_aes",
    "find_broker",
    "list_pending",
    "load_catalog",
    "load_profile",
    "match_domain",
    "match_sender",
    "render_removal",
    "save_profile",
]
