#!/usr/bin/env python3
"""Persist GURI live-scrape cache, action statuses, and relevance feedback."""

from __future__ import annotations

import json
import os
import re
from datetime import datetime
from pathlib import Path
from typing import Any, Dict, Iterable, List, Optional, Set


def _local_geofooter() -> Path:
    local = os.environ.get("LOCALAPPDATA", "")
    if local:
        return Path(local) / "GeoFooter"
    return Path(r"C:\GeoFooter")


def cache_path() -> Path:
    return _local_geofooter() / "guri_action_inbox.json"


def status_path() -> Path:
    return _local_geofooter() / "guri_action_status.json"


def feedback_path() -> Path:
    return _local_geofooter() / "guri_relevance_feedback.json"


def deadline_accounts_path() -> Path:
    return _local_geofooter() / "guri_deadline_accounts.json"


def scrape_settings_path() -> Path:
    return _local_geofooter() / "guri_scrape_settings.json"


def load_scrape_settings() -> Dict[str, Any]:
    """Outlook scrape lookback / unread prefs (persisted separately from cache)."""
    data = load_json(
        scrape_settings_path(),
        {"lookback_months": 0, "lookback_locked": False, "unread_only": False},
    )
    if not isinstance(data, dict):
        data = {}
    try:
        months = max(0, min(60, int(data.get("lookback_months") or 0)))
    except (TypeError, ValueError):
        months = 0
    return {
        "lookback_months": months,
        "lookback_locked": bool(data.get("lookback_locked")) and months > 0,
        "unread_only": bool(data.get("unread_only")),
    }


def save_scrape_settings(settings: Dict[str, Any]) -> Path:
    path = scrape_settings_path()
    try:
        months = max(0, min(60, int(settings.get("lookback_months") or 0)))
    except (TypeError, ValueError):
        months = 0
    payload = {
        "lookback_months": months,
        "lookback_locked": bool(settings.get("lookback_locked")) and months > 0,
        "unread_only": bool(settings.get("unread_only")),
        "updated": datetime.now().isoformat(timespec="seconds"),
    }
    save_json(path, payload)
    return path


def lookback_months_to_days(months: int) -> int:
    """Approximate calendar months as days for Outlook Restrict windows."""
    months = max(0, int(months or 0))
    if months <= 0:
        return 7  # default week when unset
    return max(1, months * 30)


def load_deadline_account_keys() -> List[str]:
    """Store IDs / SMTP keys selected for the Deadlines tab."""
    data = load_json(deadline_accounts_path(), {"accounts": []})
    if isinstance(data, list):
        return [str(x) for x in data if str(x).strip()]
    accounts = data.get("accounts") if isinstance(data, dict) else []
    out: List[str] = []
    for item in accounts or []:
        if isinstance(item, str) and item.strip():
            out.append(item.strip())
        elif isinstance(item, dict):
            key = str(item.get("store_id") or item.get("smtp") or "").strip()
            if key:
                out.append(key)
    return out


def save_deadline_account_keys(keys: Iterable[str]) -> Path:
    path = deadline_accounts_path()
    path.parent.mkdir(parents=True, exist_ok=True)
    cleaned = sorted({str(k).strip() for k in keys if str(k).strip()})
    path.write_text(
        json.dumps({"accounts": cleaned, "updated": datetime.now().isoformat(timespec="seconds")}, indent=2),
        encoding="utf-8",
    )
    return path


def manual_deadlines_path() -> Path:
    return _local_geofooter() / "guri_manual_deadlines.json"


def load_manual_deadlines() -> List[Dict[str, Any]]:
    """User-created deadlines (not from Outlook scrape)."""
    data = load_json(manual_deadlines_path(), {"items": []})
    if isinstance(data, list):
        items = data
    else:
        items = (data or {}).get("items") if isinstance(data, dict) else []
    out: List[Dict[str, Any]] = []
    for item in items or []:
        if not isinstance(item, dict):
            continue
        title = str(item.get("title") or "").strip()
        due = str(item.get("due") or "").strip()
        if not title or not due:
            continue
        out.append(
            {
                "id": str(item.get("id") or "").strip() or f"md-{len(out)+1}",
                "title": title,
                "due": due[:10],
                "notes": str(item.get("notes") or ""),
                "created": str(item.get("created") or ""),
                "status": str(item.get("status") or "—"),
                "mail_key": str(item.get("mail_key") or ""),
                "linked_sender": str(item.get("linked_sender") or ""),
                "linked_subject": str(item.get("linked_subject") or ""),
            }
        )
    return out


def save_manual_deadlines(items: Iterable[Dict[str, Any]]) -> Path:
    path = manual_deadlines_path()
    cleaned: List[Dict[str, Any]] = []
    for item in items or []:
        if not isinstance(item, dict):
            continue
        title = str(item.get("title") or "").strip()
        due = str(item.get("due") or "").strip()[:10]
        if not title or not due:
            continue
        cleaned.append(
            {
                "id": str(item.get("id") or "").strip(),
                "title": title,
                "due": due,
                "notes": str(item.get("notes") or ""),
                "created": str(item.get("created") or ""),
                "status": str(item.get("status") or "—"),
                "mail_key": str(item.get("mail_key") or ""),
                "linked_sender": str(item.get("linked_sender") or ""),
                "linked_subject": str(item.get("linked_subject") or ""),
            }
        )
    save_json(
        path,
        {
            "updated": datetime.now().isoformat(timespec="seconds"),
            "items": cleaned,
        },
    )
    return path


def add_manual_deadline(
    items: List[Dict[str, Any]],
    *,
    title: str,
    due: str,
    notes: str = "",
    mail_key: str = "",
    sender: str = "",
    subject: str = "",
) -> List[Dict[str, Any]]:
    """Append a manual deadline and return the updated list."""
    import uuid

    title = (title or "").strip()
    due = (due or "").strip()[:10]
    if not title or not due:
        raise ValueError("Title and due date are required")
    entry = {
        "id": str(uuid.uuid4()),
        "title": title,
        "due": due,
        "notes": (notes or "").strip(),
        "created": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        "status": "—",
        "mail_key": (mail_key or "").strip(),
        "linked_sender": (sender or "").strip(),
        "linked_subject": (subject or "").strip(),
    }
    out = list(items or [])
    out.append(entry)
    save_manual_deadlines(out)
    return out


def delete_manual_deadline(items: List[Dict[str, Any]], deadline_id: str) -> List[Dict[str, Any]]:
    deadline_id = str(deadline_id or "").strip()
    out = [x for x in (items or []) if str(x.get("id") or "") != deadline_id]
    save_manual_deadlines(out)
    return out


def load_json(path: Path, default: Any) -> Any:
    try:
        if path.is_file():
            return json.loads(path.read_text(encoding="utf-8"))
    except Exception:
        pass
    return default


def save_json(path: Path, payload: Any) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(payload, indent=2), encoding="utf-8")


def load_action_status() -> Dict[str, str]:
    """EntryID -> status label (Need Reply / Done / Replied / —)."""
    data = load_json(status_path(), {})
    return data if isinstance(data, dict) else {}


def save_action_status(status: Dict[str, str]) -> None:
    save_json(status_path(), status)


def load_scrape_cache() -> Dict[str, Any]:
    data = load_json(cache_path(), {})
    return data if isinstance(data, dict) else {}


def save_scrape_cache(
    detected: List[Dict[str, Any]],
    *,
    days: int,
    unread_only: bool,
) -> None:
    save_json(
        cache_path(),
        {
            "scraped_at": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            "days": days,
            "unread_only": unread_only,
            "items": detected,
        },
    )


def apply_statuses(
    detected_dicts: List[Dict[str, Any]],
    status_map: Dict[str, str],
) -> List[Dict[str, Any]]:
    for item in detected_dicts:
        mail = item.get("mail") or {}
        entry_id = str(mail.get("entry_id") or "")
        guri = str(mail.get("guri") or item.get("guri") or "")
        status = "—"
        if entry_id and entry_id in status_map:
            status = status_map[entry_id]
        elif guri and guri in status_map:
            status = status_map[guri]
        item["status"] = status
    return detected_dicts


# ---------------------------------------------------------------------------
# Relevance / ignore feedback (preview ticks, crosses, Ignore buttons)
# ---------------------------------------------------------------------------

_EMPTY_FEEDBACK: Dict[str, Any] = {
    "ignored_emails": [],
    "ignored_senders": [],
    "spam_senders": [],
    "enforced_emails": {},  # mail_key -> int score (absolute pin)
    "enforced_senders": {},  # normalized sender -> int score (absolute pin)
    "endorsed_emails": {},  # mail_key -> cumulative boost
    "endorsed_senders": {},  # normalized sender -> cumulative boost
    "derated_content_types": {},  # action type -> penalty int
    "derated_senders": {},  # normalized sender -> penalty int (deprecate sender)
    "deprecated_emails": {},  # mail_key -> penalty int (deprecate email; stays visible)
}


def _as_str_list(value: Any) -> List[str]:
    if not isinstance(value, list):
        return []
    return [str(x) for x in value if str(x).strip()]


def _as_int_map(value: Any) -> Dict[str, int]:
    if not isinstance(value, dict):
        return {}
    out: Dict[str, int] = {}
    for key, raw in value.items():
        try:
            out[str(key)] = int(raw)
        except (TypeError, ValueError):
            continue
    return out


def normalize_feedback(data: Any) -> Dict[str, Any]:
    base = dict(_EMPTY_FEEDBACK)
    if not isinstance(data, dict):
        return base
    base["ignored_emails"] = _as_str_list(data.get("ignored_emails"))
    base["ignored_senders"] = _as_str_list(data.get("ignored_senders"))
    base["spam_senders"] = _as_str_list(data.get("spam_senders"))
    base["enforced_emails"] = _as_int_map(data.get("enforced_emails"))
    base["enforced_senders"] = _as_int_map(data.get("enforced_senders"))
    base["endorsed_emails"] = _as_int_map(data.get("endorsed_emails"))
    base["endorsed_senders"] = _as_int_map(data.get("endorsed_senders"))
    base["derated_content_types"] = _as_int_map(data.get("derated_content_types"))
    base["derated_senders"] = _as_int_map(data.get("derated_senders"))
    base["deprecated_emails"] = _as_int_map(data.get("deprecated_emails"))
    return base


def load_relevance_feedback() -> Dict[str, Any]:
    return normalize_feedback(load_json(feedback_path(), {}))


def save_relevance_feedback(feedback: Dict[str, Any]) -> None:
    save_json(feedback_path(), normalize_feedback(feedback))


def normalize_sender_address(sender: str) -> str:
    """Extract a stable lowercase email (or blob) from a From: line."""
    text = (sender or "").strip()
    if not text:
        return ""
    match = re.search(r"<([^>]+)>", text)
    if match:
        return match.group(1).strip().lower()
    # bare address or display name only
    if "@" in text:
        return text.strip("<>\"' ").lower()
    return text.lower()


def mail_feedback_key(mail: Optional[Dict[str, Any]], item: Optional[Dict[str, Any]] = None) -> str:
    """Stable id for ignore / enforce-email feedback."""
    mail = mail or {}
    item = item or {}
    entry_id = str(mail.get("entry_id") or "").strip()
    if entry_id:
        return f"eid:{entry_id}"
    guri = str(mail.get("guri") or item.get("guri") or "").strip()
    if guri:
        return f"guri:{guri}"
    sender = normalize_sender_address(str(mail.get("sender") or ""))
    subject = str(mail.get("subject") or "").strip().lower()
    received = str(mail.get("received") or "").strip()
    return f"fallback:{sender}|{subject}|{received}"


def _unique_append(values: List[str], value: str) -> List[str]:
    value = (value or "").strip()
    if not value:
        return values
    if value not in values:
        values.append(value)
    return values


def ignore_email(feedback: Dict[str, Any], mail_key: str) -> Dict[str, Any]:
    fb = normalize_feedback(feedback)
    fb["ignored_emails"] = _unique_append(list(fb["ignored_emails"]), mail_key)
    return fb


def ignore_sender(feedback: Dict[str, Any], sender: str) -> Dict[str, Any]:
    """Remove this counterparty (From:) from prioritisation lists."""
    fb = normalize_feedback(feedback)
    addr = normalize_sender_address(sender)
    if addr:
        fb["ignored_senders"] = _unique_append(list(fb["ignored_senders"]), addr)
    return fb


# Alias used by the Welcome preview label "Ignore Recipient"
ignore_recipient = ignore_sender


def deprecate_email(
    feedback: Dict[str, Any],
    mail_key: str,
    *,
    penalty: int = 25,
) -> Dict[str, Any]:
    """Lower this email's relevance score; keep it visible in lists."""
    fb = normalize_feedback(feedback)
    if mail_key:
        fb["deprecated_emails"][mail_key] = max(
            int(fb["deprecated_emails"].get(mail_key, 0)),
            int(penalty),
        )
        # Clear pins / endorsements so deprecate can take effect
        fb["enforced_emails"].pop(mail_key, None)
        fb["endorsed_emails"].pop(mail_key, None)
    return fb


def mark_spam(feedback: Dict[str, Any], *, mail_key: str, sender: str) -> Dict[str, Any]:
    fb = ignore_email(feedback, mail_key)
    addr = normalize_sender_address(sender)
    if addr:
        fb["spam_senders"] = _unique_append(list(fb["spam_senders"]), addr)
        fb["ignored_senders"] = _unique_append(list(fb["ignored_senders"]), addr)
    return fb


def enforce_email_score(feedback: Dict[str, Any], mail_key: str, score: int) -> Dict[str, Any]:
    fb = normalize_feedback(feedback)
    if mail_key:
        fb["enforced_emails"][mail_key] = max(0, min(100, int(score)))
    return fb


def enforce_sender_score(feedback: Dict[str, Any], sender: str, score: int) -> Dict[str, Any]:
    fb = normalize_feedback(feedback)
    addr = normalize_sender_address(sender)
    if addr:
        fb["enforced_senders"][addr] = max(0, min(100, int(score)))
    return fb


ENDORSE_STEP = 8


def endorse_email_score(
    feedback: Dict[str, Any],
    mail_key: str,
    *,
    step: int = ENDORSE_STEP,
) -> Dict[str, Any]:
    """Gradually raise this email's relevance (cumulative boost)."""
    fb = normalize_feedback(feedback)
    if not mail_key:
        return fb
    step = max(1, int(step))
    current = int(fb["endorsed_emails"].get(mail_key, 0))
    fb["endorsed_emails"][mail_key] = min(100, current + step)
    # Soften deprecate penalty if present
    if mail_key in fb["deprecated_emails"]:
        remaining = max(0, int(fb["deprecated_emails"].get(mail_key, 0)) - step)
        if remaining:
            fb["deprecated_emails"][mail_key] = remaining
        else:
            fb["deprecated_emails"].pop(mail_key, None)
    # Clear absolute pin so the boost is visible
    fb["enforced_emails"].pop(mail_key, None)
    return fb


def endorse_sender_score(
    feedback: Dict[str, Any],
    sender: str,
    *,
    step: int = ENDORSE_STEP,
) -> Dict[str, Any]:
    """Gradually raise this sender's relevance (cumulative boost)."""
    fb = normalize_feedback(feedback)
    addr = normalize_sender_address(sender)
    if not addr:
        return fb
    step = max(1, int(step))
    current = int(fb["endorsed_senders"].get(addr, 0))
    fb["endorsed_senders"][addr] = min(100, current + step)
    if addr in fb["derated_senders"]:
        remaining = max(0, int(fb["derated_senders"].get(addr, 0)) - step)
        if remaining:
            fb["derated_senders"][addr] = remaining
        else:
            fb["derated_senders"].pop(addr, None)
    fb["enforced_senders"].pop(addr, None)
    return fb


def derate_content_types(
    feedback: Dict[str, Any],
    action_types: Iterable[str],
    *,
    penalty: int = 20,
) -> Dict[str, Any]:
    fb = normalize_feedback(feedback)
    for action in action_types:
        key = str(action or "").strip().lower()
        if not key:
            continue
        fb["derated_content_types"][key] = max(
            int(fb["derated_content_types"].get(key, 0)),
            int(penalty),
        )
    return fb


def derate_sender(
    feedback: Dict[str, Any],
    sender: str,
    *,
    penalty: int = 30,
) -> Dict[str, Any]:
    fb = normalize_feedback(feedback)
    addr = normalize_sender_address(sender)
    if addr:
        fb["derated_senders"][addr] = max(
            int(fb["derated_senders"].get(addr, 0)),
            int(penalty),
        )
        fb["enforced_senders"].pop(addr, None)
        fb["endorsed_senders"].pop(addr, None)
    return fb


def deprecate_sender(
    feedback: Dict[str, Any],
    sender: str,
    *,
    penalty: int = 30,
) -> Dict[str, Any]:
    """Lower sender relevance for this address; keep messages visible."""
    return derate_sender(feedback, sender, penalty=penalty)


def is_mail_ignored(feedback: Dict[str, Any], mail: Dict[str, Any], item: Optional[Dict[str, Any]] = None) -> bool:
    fb = normalize_feedback(feedback)
    key = mail_feedback_key(mail, item)
    if key in set(fb["ignored_emails"]):
        return True
    addr = normalize_sender_address(str((mail or {}).get("sender") or ""))
    if not addr:
        return False
    blocked: Set[str] = set(fb["ignored_senders"]) | set(fb["spam_senders"])
    return addr in blocked


def is_mail_junked(
    feedback: Dict[str, Any],
    mail: Dict[str, Any],
    item: Optional[Dict[str, Any]] = None,
) -> bool:
    """True when ignored/spam or explicitly deprecated (email or sender)."""
    fb = normalize_feedback(feedback)
    if is_mail_ignored(fb, mail, item):
        return True
    key = mail_feedback_key(mail, item)
    if int(fb["deprecated_emails"].get(key, 0)) > 0:
        return True
    addr = normalize_sender_address(str((mail or {}).get("sender") or ""))
    if addr and int(fb["derated_senders"].get(addr, 0)) > 0:
        return True
    return False


def compute_relevance_scores(
    item: Dict[str, Any],
    feedback: Dict[str, Any],
    *,
    important_domains: Optional[Iterable[str]] = None,
    important_senders: Optional[Iterable[str]] = None,
) -> Dict[str, int]:
    """Return display/adjust scores: email_relevance, sender_relevance."""
    fb = normalize_feedback(feedback)
    mail = item.get("mail") or {}
    if not mail and item.get("sender"):
        mail = item
    key = mail_feedback_key(mail, item)
    sender = str(mail.get("sender") or "")
    addr = normalize_sender_address(sender)
    base_score = int(item.get("score") or 0)
    actions = [str(a).lower() for a in (item.get("actions") or [])]

    email_score = base_score
    for action in actions:
        email_score -= int(fb["derated_content_types"].get(action, 0))
    if addr:
        email_score -= int(fb["derated_senders"].get(addr, 0))
    email_score -= int(fb["deprecated_emails"].get(key, 0))
    email_score += int(fb["endorsed_emails"].get(key, 0))
    if key in fb["enforced_emails"]:
        email_score = int(fb["enforced_emails"][key])
    email_score = max(0, min(100, email_score))

    domains = {d.strip().lower() for d in (important_domains or []) if d and d.strip()}
    vip = {normalize_sender_address(s) for s in (important_senders or []) if normalize_sender_address(s)}
    domain = ""
    if "@" in addr:
        domain = addr.rsplit("@", 1)[-1]
    sender_score = 35
    if domain and any(domain == d or domain.endswith("." + d) for d in domains):
        sender_score = 75
    if addr and addr in vip:
        sender_score = max(sender_score, 88)
    # Blend with this message's adjusted score so the control feels connected
    sender_score = int(round(0.55 * sender_score + 0.45 * email_score))
    if addr:
        sender_score -= int(fb["derated_senders"].get(addr, 0))
        sender_score += int(fb["endorsed_senders"].get(addr, 0))
    if addr in fb["enforced_senders"]:
        sender_score = int(fb["enforced_senders"][addr])
    sender_score = max(0, min(100, sender_score))

    return {
        "email_relevance": email_score,
        "sender_relevance": sender_score,
    }


def apply_relevance_feedback(
    detected_dicts: List[Dict[str, Any]],
    feedback: Dict[str, Any],
    *,
    important_domains: Optional[Iterable[str]] = None,
    important_senders: Optional[Iterable[str]] = None,
    drop_ignored: bool = True,
) -> List[Dict[str, Any]]:
    """Annotate scores; optionally drop ignored/spam mail."""
    fb = normalize_feedback(feedback)
    out: List[Dict[str, Any]] = []
    for raw in detected_dicts:
        item = dict(raw)
        mail = dict(item.get("mail") or {})
        if not mail and item.get("sender"):
            mail = dict(item)
        if drop_ignored and is_mail_ignored(fb, mail, item):
            continue
        scores = compute_relevance_scores(
            item,
            fb,
            important_domains=important_domains,
            important_senders=important_senders,
        )
        item["email_relevance"] = scores["email_relevance"]
        item["sender_relevance"] = scores["sender_relevance"]
        # Keep list sort in sync with enforced / derated email relevance
        item["score"] = scores["email_relevance"]
        out.append(item)
    return out
