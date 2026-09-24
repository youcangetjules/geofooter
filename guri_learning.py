#!/usr/bin/env python3
"""Persisted personalization: important senders, deadline corrections, confirmations."""

from __future__ import annotations

import re
from datetime import datetime
from pathlib import Path
from typing import Any, Dict, Iterable, List, Optional, Sequence, Set

from guri_action_store import (
    _local_geofooter,
    load_json,
    mail_feedback_key,
    normalize_sender_address,
    save_json,
)

# Deadline timeline rows (source of the due item)
DL_CAT_EXTRACTED = "extracted"
DL_CAT_MANUAL = "manual"
DEADLINE_TIMELINE_CATEGORIES = (DL_CAT_EXTRACTED, DL_CAT_MANUAL)
DEADLINE_TIMELINE_LABELS = {
    DL_CAT_EXTRACTED: "Extracted from email",
    DL_CAT_MANUAL: "Added by you",
}


def learning_path() -> Path:
    return _local_geofooter() / "guri_learning.json"


_EMPTY_LEARNING: Dict[str, Any] = {
    "important_senders": [],  # lowercase email addresses
    "asked_important_senders": [],  # already prompted — don't nag
    "friend_senders": [],  # personal / friend addresses (soft relevant lean)
    "deadline_cues": [],  # extra phrases learned from corrections
    "confirmed_deadline_mails": {},  # mail_key -> {due, title, sender, subject, at}
    "rejected_deadline_mails": [],  # mail_key false positives
    "linked_manual": {},  # manual_id -> {mail_key, sender, subject, at}
    "suppress_email_link_prompt": False,
    # Timeline category corrections
    "category_overrides": {},  # mail_key -> {category, reasons, from_category, at}
    "sender_category_votes": {},  # sender -> {category: count}
    "sender_reason_tags": {},  # sender -> {reason_key: count}
    "demoted_actions": {},  # action type -> count of "not actionable" corrections
    "category_reason_counts": {},  # reason key -> count
    # Who "me" is — used to drop deadlines not addressed to you
    "my_names": [],  # e.g. Julian, Julian Garrett
    "my_emails": [],  # lowercase addresses
    "other_names": [],  # greetings learned as not-you (e.g. Lee)
    "filter_deadlines_not_for_me": True,
    # Natural-language rules for the LLM (and future heuristic application)
    "user_rules": [],  # [{id, text, scope, mail_key, sender, at, active}]
}

# Greeting at the start of the latest body block
_GREETING_RE = re.compile(
    r"(?im)^\s*(?:hi|hello|hey|dear|good\s+(?:morning|afternoon|evening))"
    r"(?:\s+there)?\s*[,:]?\s+"
    r"(?P<name>[A-Za-z][A-Za-z'\-]{1,40})"
    r"\b"
)

_GROUP_GREETINGS = {
    "all",
    "everyone",
    "team",
    "folks",
    "guys",
    "both",
    "colleagues",
    "friends",
    "list",
}

# Why-checkboxes (stable keys → UI labels)
# Positives teach VIP / friend / importance; negatives demote noise.
CATEGORY_CORRECTION_REASONS_POSITIVE = (
    ("from_boss", "Email from boss (definitely actionable)"),
    ("from_friend", "Email from a friend / personal"),
    ("important_email", "Important email"),
    ("needs_reply", "Needs a reply from me"),
    ("customer_client", "Customer / client"),
    ("deadline_real", "Real deadline / genuine urgency"),
)

CATEGORY_CORRECTION_REASONS_NEGATIVE = (
    ("not_actionable", "Not actionable"),
    ("not_important", "Not important"),
    ("marketing", "Marketing / newsletter"),
    ("no_action_needed", "No action needed from me"),
    ("wrong_deadline", "Wrong deadline / urgency signal"),
)

CATEGORY_CORRECTION_REASONS = (
    CATEGORY_CORRECTION_REASONS_POSITIVE + CATEGORY_CORRECTION_REASONS_NEGATIVE
)

CATEGORY_CORRECTION_REASON_GROUPS = (
    ("Why it matters (positives)", CATEGORY_CORRECTION_REASONS_POSITIVE),
    ("Why auto-category was wrong", CATEGORY_CORRECTION_REASONS_NEGATIVE),
)

# Positive reason → learning hints (applied gradually via votes / VIP lists)
POSITIVE_REASON_LEARNING: Dict[str, Dict[str, Any]] = {
    "from_boss": {
        "prefer": "actionable",
        "important_sender": True,
        "extra_votes": 1,
    },
    "from_friend": {
        "prefer": "relevant",
        "friend_sender": True,
        "extra_votes": 1,
    },
    "important_email": {
        "prefer": "actionable",
        "important_sender": True,
        "extra_votes": 1,
    },
    "needs_reply": {
        "prefer": "actionable",
        "extra_votes": 1,
    },
    "customer_client": {
        "prefer": "actionable",
        "important_sender": True,
        "extra_votes": 1,
    },
    "deadline_real": {
        "prefer": "actionable",
        "extra_votes": 1,
    },
}

# Targets offered in the wrong-categorisation dialog (importance ladder)
CATEGORY_RECAT_TARGETS = (
    ("important", "Important"),
    ("actionable", "Actionable"),
    ("relevant", "Relevant but not actionable"),
    ("less", "Less relevant"),
    ("deprecated", "Deprecated"),
)

# Ascending importance — corrections move one rung per save (no knee-jerk jumps)
CATEGORY_LADDER = (
    "deprecated",
    "less",
    "relevant",
    "actionable",
    "important",
)


def category_ladder_index(category: str) -> int:
    cat = str(category or "").strip().lower()
    try:
        return CATEGORY_LADDER.index(cat)
    except ValueError:
        return CATEGORY_LADDER.index("less")


def step_category_toward(
    current: str,
    target: str,
    *,
    reasons: Optional[Iterable[str]] = None,
) -> str:
    """Move one rung on the category ladder toward ``target``.

    Corrections stay gradual so GURI learns directionally — but never park an
    explicitly-actionable promotion in ``relevant`` ("Relevant but not
    actionable"). That middle rung is skipped when the user (or a positive
    why-reason) says the mail *is* actionable / important.
    """
    cur_i = category_ladder_index(current)
    tgt_i = category_ladder_index(target)
    intended = str(target or "").strip().lower()

    actionable_intent = intended in {"actionable", "important"}
    for raw in reasons or []:
        hint = POSITIVE_REASON_LEARNING.get(str(raw or "").strip()) or {}
        if str(hint.get("prefer") or "").strip().lower() == "actionable":
            actionable_intent = True
            break

    if cur_i == tgt_i:
        return CATEGORY_LADDER[cur_i]

    if cur_i < tgt_i:
        nxt = CATEGORY_LADDER[cur_i + 1]
        # Do not put an actionable item into a non-actionable category.
        if (
            actionable_intent
            and nxt == "relevant"
            and tgt_i > category_ladder_index("relevant")
        ):
            act_i = category_ladder_index("actionable")
            return CATEGORY_LADDER[min(tgt_i, act_i)]
        return nxt

    # Demoting: one rung down (relevant is a valid soft landing from actionable)
    return CATEGORY_LADDER[cur_i - 1]

# After this many "not actionable" hits on an action type, ignore it for triage
DEMOTE_ACTION_THRESHOLD = 3
# After this many votes for the same category on a sender, prefer it
SENDER_CATEGORY_VOTE_THRESHOLD = 2


def _as_str_list(value: Any) -> List[str]:
    if not isinstance(value, list):
        return []
    return [str(x).strip() for x in value if str(x).strip()]


_RULE_CATEGORIES = {"important", "actionable", "relevant", "less", "deprecated"}
_RULE_IMPORTANCE = {"high", "medium", "low"}


def normalize_rule_match(raw: Any) -> Dict[str, Any]:
    """Structured matchers compiled from a natural-language rule."""
    if not isinstance(raw, dict):
        raw = {}
    domains = sorted(
        {
            d.strip().lower().lstrip("@")
            for d in _as_str_list(raw.get("domains"))
            if d.strip()
        }
    )
    senders = sorted(
        {
            normalize_sender_address(s)
            for s in _as_str_list(raw.get("senders"))
            if normalize_sender_address(s)
        }
    )
    return {
        "domains": domains[:40],
        "senders": senders[:40],
        "subject_contains": [s.lower() for s in _as_str_list(raw.get("subject_contains"))][:40],
        "body_contains": [s.lower() for s in _as_str_list(raw.get("body_contains"))][:40],
        "any_keywords": [s.lower() for s in _as_str_list(raw.get("any_keywords"))][:40],
    }


def normalize_rule_effects(raw: Any) -> Dict[str, Any]:
    """Local enforcement effects for a compiled rule."""
    if not isinstance(raw, dict):
        raw = {}
    cat = str(raw.get("category") or "").strip().lower()
    if cat not in _RULE_CATEGORIES:
        cat = ""
    importance = str(raw.get("importance") or "").strip().lower()
    if importance not in _RULE_IMPORTANCE:
        importance = ""
    try:
        boost_score = int(raw.get("boost_score") or 0)
    except (TypeError, ValueError):
        boost_score = 0
    try:
        boost_sender = int(raw.get("boost_sender") or 0)
    except (TypeError, ValueError):
        boost_sender = 0
    demote = [str(a).strip().lower() for a in _as_str_list(raw.get("demote_actions"))]
    return {
        "category": cat,
        "importance": importance,
        "boost_score": max(-50, min(50, boost_score)),
        "boost_sender": max(-50, min(50, boost_sender)),
        "demote_actions": demote[:20],
        "hide_from_actions": bool(raw.get("hide_from_actions")),
        "hide_from_important": bool(raw.get("hide_from_important")),
        "skip_deadline": bool(raw.get("skip_deadline")),
        "force_vip": bool(raw.get("force_vip")),
    }


def rule_has_enforceable_effects(effects: Optional[Dict[str, Any]]) -> bool:
    eff = normalize_rule_effects(effects or {})
    return bool(
        eff.get("category")
        or eff.get("importance")
        or eff.get("boost_score")
        or eff.get("boost_sender")
        or eff.get("demote_actions")
        or eff.get("hide_from_actions")
        or eff.get("hide_from_important")
        or eff.get("skip_deadline")
        or eff.get("force_vip")
    )


def _as_str_map(value: Any) -> Dict[str, Any]:
    if not isinstance(value, dict):
        return {}
    return {str(k): v for k, v in value.items() if str(k).strip()}


def normalize_learning(data: Any) -> Dict[str, Any]:
    base = dict(_EMPTY_LEARNING)
    if not isinstance(data, dict):
        return base
    base["important_senders"] = sorted(
        {normalize_sender_address(x) for x in _as_str_list(data.get("important_senders")) if normalize_sender_address(x)}
    )
    base["asked_important_senders"] = sorted(
        {
            normalize_sender_address(x)
            for x in _as_str_list(data.get("asked_important_senders"))
            if normalize_sender_address(x)
        }
    )
    base["friend_senders"] = sorted(
        {
            normalize_sender_address(x)
            for x in _as_str_list(data.get("friend_senders"))
            if normalize_sender_address(x)
        }
    )
    cues: List[str] = []
    seen_cues: Set[str] = set()
    for cue in _as_str_list(data.get("deadline_cues")):
        key = cue.lower()
        if key in seen_cues or len(cue) < 3:
            continue
        seen_cues.add(key)
        cues.append(cue)
    base["deadline_cues"] = cues[:80]
    confirmed: Dict[str, Any] = {}
    for key, raw in _as_str_map(data.get("confirmed_deadline_mails")).items():
        if isinstance(raw, dict):
            confirmed[key] = {
                "due": str(raw.get("due") or "")[:10],
                "title": str(raw.get("title") or ""),
                "sender": str(raw.get("sender") or ""),
                "subject": str(raw.get("subject") or ""),
                "at": str(raw.get("at") or ""),
            }
        else:
            confirmed[key] = {"due": "", "title": "", "sender": "", "subject": "", "at": ""}
    base["confirmed_deadline_mails"] = confirmed
    base["rejected_deadline_mails"] = _as_str_list(data.get("rejected_deadline_mails"))
    linked: Dict[str, Any] = {}
    for mid, raw in _as_str_map(data.get("linked_manual")).items():
        if not isinstance(raw, dict):
            continue
        linked[mid] = {
            "mail_key": str(raw.get("mail_key") or ""),
            "sender": str(raw.get("sender") or ""),
            "subject": str(raw.get("subject") or ""),
            "at": str(raw.get("at") or ""),
        }
    base["linked_manual"] = linked
    base["suppress_email_link_prompt"] = bool(data.get("suppress_email_link_prompt"))

    overrides: Dict[str, Any] = {}
    for key, raw in _as_str_map(data.get("category_overrides")).items():
        if not isinstance(raw, dict):
            continue
        cat = str(raw.get("category") or "").strip().lower()
        if cat not in {"important", "actionable", "relevant", "less", "deprecated"}:
            continue
        reasons = _as_str_list(raw.get("reasons"))
        overrides[key] = {
            "category": cat,
            "intended_category": str(raw.get("intended_category") or cat).strip().lower(),
            "reasons": reasons,
            "from_category": str(raw.get("from_category") or ""),
            "at": str(raw.get("at") or ""),
            "sender": str(raw.get("sender") or ""),
        }
    base["category_overrides"] = overrides

    votes: Dict[str, Any] = {}
    for sender, raw in _as_str_map(data.get("sender_category_votes")).items():
        addr = normalize_sender_address(sender)
        if not addr or not isinstance(raw, dict):
            continue
        cleaned: Dict[str, int] = {}
        for cat, count in raw.items():
            c = str(cat).strip().lower()
            if c not in {"important", "actionable", "relevant", "less", "deprecated"}:
                continue
            try:
                cleaned[c] = max(0, int(count))
            except (TypeError, ValueError):
                continue
        if cleaned:
            votes[addr] = cleaned
    base["sender_category_votes"] = votes
    base["demoted_actions"] = _as_int_map(data.get("demoted_actions"))
    base["category_reason_counts"] = _as_int_map(data.get("category_reason_counts"))
    reason_tags: Dict[str, Any] = {}
    for sender, raw in _as_str_map(data.get("sender_reason_tags")).items():
        addr = normalize_sender_address(sender)
        if not addr or not isinstance(raw, dict):
            continue
        reason_tags[addr] = _as_int_map(raw)
    base["sender_reason_tags"] = reason_tags

    names: List[str] = []
    seen_names: Set[str] = set()
    for name in _as_str_list(data.get("my_names")):
        key = name.lower()
        if key in seen_names or len(name) < 2:
            continue
        seen_names.add(key)
        names.append(name)
    base["my_names"] = names[:40]

    emails: List[str] = []
    seen_emails: Set[str] = set()
    for raw in _as_str_list(data.get("my_emails")):
        addr = normalize_sender_address(raw)
        if not addr or addr in seen_emails:
            continue
        seen_emails.add(addr)
        emails.append(addr)
    base["my_emails"] = sorted(emails)

    others: List[str] = []
    seen_others: Set[str] = set()
    for name in _as_str_list(data.get("other_names")):
        key = name.lower()
        if key in seen_others or key in seen_names or len(name) < 2:
            continue
        seen_others.add(key)
        others.append(name)
    base["other_names"] = others[:80]

    if "filter_deadlines_not_for_me" in data:
        base["filter_deadlines_not_for_me"] = bool(data.get("filter_deadlines_not_for_me"))
    else:
        base["filter_deadlines_not_for_me"] = True

    rules: List[Dict[str, Any]] = []
    raw_rules = data.get("user_rules")
    if isinstance(raw_rules, list):
        for item in raw_rules:
            if not isinstance(item, dict):
                continue
            text = str(item.get("text") or "").strip()
            if len(text) < 3:
                continue
            scope = str(item.get("scope") or "always").strip().lower()
            if scope not in {"always", "sender", "email"}:
                scope = "always"
            rules.append(
                {
                    "id": str(item.get("id") or "").strip()
                    or f"rule-{len(rules)+1}",
                    "text": text[:2000],
                    "scope": scope,
                    "mail_key": str(item.get("mail_key") or ""),
                    "sender": str(item.get("sender") or ""),
                    "at": str(item.get("at") or ""),
                    "active": bool(item.get("active", True)),
                    "compiled": bool(item.get("compiled")),
                    "compiled_at": str(item.get("compiled_at") or ""),
                    "match": normalize_rule_match(item.get("match")),
                    "effects": normalize_rule_effects(item.get("effects")),
                }
            )
    base["user_rules"] = rules[:200]
    return base


def _as_int_map(value: Any) -> Dict[str, int]:
    if not isinstance(value, dict):
        return {}
    out: Dict[str, int] = {}
    for key, raw in value.items():
        try:
            out[str(key)] = max(0, int(raw))
        except (TypeError, ValueError):
            continue
    return out


def load_learning() -> Dict[str, Any]:
    return normalize_learning(load_json(learning_path(), {}))


def save_learning(learning: Dict[str, Any]) -> None:
    payload = normalize_learning(learning)
    payload["updated"] = datetime.now().isoformat(timespec="seconds")
    save_json(learning_path(), payload)


def is_important_sender(learning: Dict[str, Any], sender: str) -> bool:
    addr = normalize_sender_address(sender)
    if not addr:
        return False
    return addr in set(normalize_learning(learning)["important_senders"])


def is_friend_sender(learning: Dict[str, Any], sender: str) -> bool:
    addr = normalize_sender_address(sender)
    if not addr:
        return False
    return addr in set(normalize_learning(learning).get("friend_senders") or [])


def remember_important_sender(learning: Dict[str, Any], sender: str) -> Dict[str, Any]:
    fb = normalize_learning(learning)
    addr = normalize_sender_address(sender)
    if addr and addr not in fb["important_senders"]:
        fb["important_senders"].append(addr)
        fb["important_senders"] = sorted(set(fb["important_senders"]))
    if addr and addr not in fb["asked_important_senders"]:
        fb["asked_important_senders"].append(addr)
        fb["asked_important_senders"] = sorted(set(fb["asked_important_senders"]))
    return fb


def forget_important_sender(learning: Dict[str, Any], sender: str) -> Dict[str, Any]:
    """Remove a sender from the VIP / important-senders list."""
    fb = normalize_learning(learning)
    addr = normalize_sender_address(sender)
    if not addr:
        return fb
    fb["important_senders"] = sorted(
        {a for a in fb["important_senders"] if a != addr}
    )
    return fb


def remember_friend_sender(learning: Dict[str, Any], sender: str) -> Dict[str, Any]:
    fb = normalize_learning(learning)
    addr = normalize_sender_address(sender)
    if addr and addr not in fb["friend_senders"]:
        fb["friend_senders"].append(addr)
        fb["friend_senders"] = sorted(set(fb["friend_senders"]))
    return fb


def mark_asked_important_sender(learning: Dict[str, Any], sender: str) -> Dict[str, Any]:
    fb = normalize_learning(learning)
    addr = normalize_sender_address(sender)
    if addr and addr not in fb["asked_important_senders"]:
        fb["asked_important_senders"].append(addr)
        fb["asked_important_senders"] = sorted(set(fb["asked_important_senders"]))
    return fb


def should_ask_important_sender(learning: Dict[str, Any], sender: str) -> bool:
    fb = normalize_learning(learning)
    addr = normalize_sender_address(sender)
    if not addr:
        return False
    if addr in set(fb["important_senders"]):
        return False
    if addr in set(fb["asked_important_senders"]):
        return False
    return True


def learn_deadline_cue(learning: Dict[str, Any], phrase: str) -> Dict[str, Any]:
    """Store a free-text cue phrase (escaped later when used as regex)."""
    fb = normalize_learning(learning)
    phrase = re.sub(r"\s+", " ", (phrase or "").strip())
    if len(phrase) < 4 or len(phrase) > 80:
        return fb
    # Avoid storing pure dates as cues
    if re.fullmatch(r"[\d\s/\-.]+", phrase):
        return fb
    lower = phrase.lower()
    existing = {c.lower() for c in fb["deadline_cues"]}
    if lower not in existing:
        fb["deadline_cues"].append(phrase)
        fb["deadline_cues"] = fb["deadline_cues"][:80]
    return fb


def confirm_deadline_mail(
    learning: Dict[str, Any],
    *,
    mail_key: str,
    due: str = "",
    title: str = "",
    sender: str = "",
    subject: str = "",
) -> Dict[str, Any]:
    fb = normalize_learning(learning)
    mail_key = (mail_key or "").strip()
    if not mail_key:
        return fb
    fb["confirmed_deadline_mails"][mail_key] = {
        "due": (due or "")[:10],
        "title": title or "",
        "sender": sender or "",
        "subject": subject or "",
        "at": datetime.now().isoformat(timespec="seconds"),
    }
    # Clear rejection if user later confirms
    fb["rejected_deadline_mails"] = [k for k in fb["rejected_deadline_mails"] if k != mail_key]
    return fb


def reject_deadline_mail(learning: Dict[str, Any], mail_key: str) -> Dict[str, Any]:
    fb = normalize_learning(learning)
    mail_key = (mail_key or "").strip()
    if not mail_key:
        return fb
    if mail_key not in fb["rejected_deadline_mails"]:
        fb["rejected_deadline_mails"].append(mail_key)
    fb["confirmed_deadline_mails"].pop(mail_key, None)
    return fb


def is_deadline_rejected(learning: Dict[str, Any], mail_key: str) -> bool:
    return (mail_key or "") in set(normalize_learning(learning)["rejected_deadline_mails"])


def link_manual_deadline(
    learning: Dict[str, Any],
    *,
    manual_id: str,
    mail_key: str,
    sender: str = "",
    subject: str = "",
) -> Dict[str, Any]:
    fb = normalize_learning(learning)
    mid = (manual_id or "").strip()
    key = (mail_key or "").strip()
    if not mid or not key:
        return fb
    fb["linked_manual"][mid] = {
        "mail_key": key,
        "sender": sender or "",
        "subject": subject or "",
        "at": datetime.now().isoformat(timespec="seconds"),
    }
    return fb


def cue_patterns_from_learning(learning: Dict[str, Any]) -> List[str]:
    """Regex fragments safe to OR into deadline extraction."""
    fb = normalize_learning(learning)
    out: List[str] = []
    for cue in fb["deadline_cues"]:
        escaped = re.escape(cue)
        # Allow flexible whitespace inside multi-word cues
        escaped = re.sub(r"\\\s+", r"\\s+", escaped)
        out.append(escaped)
    return out


def domain_from_sender(sender: str) -> str:
    addr = normalize_sender_address(sender)
    if "@" not in addr:
        return ""
    return addr.rsplit("@", 1)[-1].strip().lower()


def candidate_deadline_emails(
    items: Sequence[Dict[str, Any]],
    *,
    limit: int = 12,
) -> List[Dict[str, Any]]:
    """Recent scrape items suitable for 'is this the email?' prompts."""
    scored: List[tuple] = []
    for det in items or []:
        mail = det.get("mail") or {}
        if not mail:
            continue
        key = mail_feedback_key(mail, det)
        subject = str(mail.get("subject") or "").strip()
        sender = str(mail.get("sender") or "").strip()
        if not subject and not sender:
            continue
        actions = [str(a).lower() for a in (det.get("actions") or [])]
        score = int(det.get("score") or 0)
        boost = 20 if "deadline" in actions else 0
        boost += 10 if det.get("deadlines") else 0
        scored.append(
            (
                score + boost,
                str(mail.get("received") or ""),
                {
                    "mail_key": key,
                    "sender": sender,
                    "subject": subject,
                    "received": str(mail.get("received") or ""),
                    "det": det,
                    "label": f"{sender[:40]} — {subject[:55]}"[:100],
                },
            )
        )
    scored.sort(key=lambda t: (t[0], t[1]), reverse=True)
    return [row for _s, _r, row in scored[: max(1, limit)]]


def record_category_correction(
    learning: Dict[str, Any],
    *,
    mail_key: str,
    sender: str,
    from_category: str,
    to_category: str,
    reasons: Iterable[str],
    actions: Optional[Iterable[str]] = None,
    instant: bool = False,
) -> Dict[str, Any]:
    """Persist a wrong-categorisation correction and update aggregate signals.

    The stored override is one ladder step toward the user's chosen target
    (gradual learning). ``intended_category`` keeps the full preference.
    Explicit actionable intent skips the non-actionable ``relevant`` rung.
    ``instant=True`` jumps straight to the intended category (Important overrides).
    """
    fb = normalize_learning(learning)
    mail_key = (mail_key or "").strip()
    intended = str(to_category or "").strip().lower()
    if intended not in {"important", "actionable", "relevant", "less", "deprecated"}:
        return fb
    reason_list = [str(r).strip() for r in (reasons or []) if str(r).strip()]
    if instant:
        applied = intended
    else:
        applied = step_category_toward(
            from_category, intended, reasons=reason_list
        )
    addr = normalize_sender_address(sender)

    if mail_key:
        fb["category_overrides"][mail_key] = {
            "category": applied,
            "intended_category": intended,
            "reasons": reason_list,
            "from_category": str(from_category or ""),
            "at": datetime.now().isoformat(timespec="seconds"),
            "sender": addr or (sender or ""),
        }

    for reason in reason_list:
        fb["category_reason_counts"][reason] = int(
            fb["category_reason_counts"].get(reason, 0)
        ) + 1

    # Not actionable / no action needed → demote the action types on this mail
    # Only when stepping down (or staying demoted), not when promoting.
    stepping_down = category_ladder_index(applied) < category_ladder_index(
        from_category
    )
    if stepping_down and any(
        r in {"not_actionable", "no_action_needed", "wrong_deadline"}
        for r in reason_list
    ):
        for action in actions or []:
            key = str(action or "").strip().lower()
            if not key:
                continue
            fb["demoted_actions"][key] = int(fb["demoted_actions"].get(key, 0)) + 1

    # Sender votes follow the applied (stepped) category — gradual preference
    if addr:
        votes = dict(fb["sender_category_votes"].get(addr) or {})
        votes[applied] = int(votes.get(applied, 0)) + 1
        tags = dict(fb.get("sender_reason_tags", {}).get(addr) or {})
        mark_important = False
        mark_friend = False

        for reason in reason_list:
            tags[reason] = int(tags.get(reason, 0)) + 1
            hint = POSITIVE_REASON_LEARNING.get(reason) or {}
            prefer = str(hint.get("prefer") or "").strip().lower()
            if prefer in {"important", "actionable", "relevant", "less", "deprecated"}:
                extra = max(1, int(hint.get("extra_votes") or 1))
                votes[prefer] = int(votes.get(prefer, 0)) + extra
            if hint.get("important_sender"):
                mark_important = True
            if hint.get("friend_sender"):
                mark_friend = True

        fb["sender_category_votes"][addr] = votes
        fb["sender_reason_tags"][addr] = tags
        if mark_important:
            fb = remember_important_sender(fb, addr)
        if mark_friend:
            fb = remember_friend_sender(fb, addr)

    return fb


def preferred_sender_category(learning: Dict[str, Any], sender: str) -> Optional[str]:
    """Return a learned preferred category for a sender, if votes are strong enough."""
    fb = normalize_learning(learning)
    addr = normalize_sender_address(sender)
    if not addr:
        return None
    votes = fb["sender_category_votes"].get(addr) or {}
    if not votes:
        return None
    best_cat = ""
    best_n = 0
    for cat, n in votes.items():
        try:
            count = int(n)
        except (TypeError, ValueError):
            continue
        if count > best_n:
            best_n = count
            best_cat = str(cat)
    if best_n >= SENDER_CATEGORY_VOTE_THRESHOLD and best_cat:
        return best_cat
    return None


def actions_for_categorization(
    actions: Iterable[str],
    learning: Dict[str, Any],
) -> List[str]:
    """Drop action types the user has repeatedly marked as not actionable."""
    fb = normalize_learning(learning)
    demoted = fb.get("demoted_actions") or {}
    out: List[str] = []
    for action in actions or []:
        key = str(action or "").strip().lower()
        if not key:
            continue
        if int(demoted.get(key, 0)) >= DEMOTE_ACTION_THRESHOLD:
            continue
        out.append(key)
    return out


def resolve_timeline_category(
    detected: Any,
    learning: Optional[Dict[str, Any]] = None,
    *,
    junked: bool = False,
    base_categorize=None,
    important_domains: Optional[Sequence[str]] = None,
) -> str:
    """Pick timeline category using overrides + sender votes + learned demotions."""
    from guri_action_detector import (
        CATEGORY_ACTIONABLE,
        CATEGORY_DEPRECATED,
        CATEGORY_IMPORTANT,
        CATEGORY_LESS,
        CATEGORY_RELEVANT,
        categorize_detected,
        sender_matches_important_domain,
    )

    if junked:
        return CATEGORY_DEPRECATED

    if isinstance(detected, dict):
        data = detected
    else:
        data = {
            "actions": getattr(detected, "actions", []) or [],
            "score": getattr(detected, "score", 0),
            "importance": getattr(detected, "importance", "low"),
            "mail": getattr(detected, "mail", {}) or {},
        }

    # Compiled user-rule hard override (local enforcement)
    forced = str(data.get("rule_forced_category") or "").strip().lower()
    if forced in {
        CATEGORY_IMPORTANT,
        CATEGORY_ACTIONABLE,
        CATEGORY_RELEVANT,
        CATEGORY_LESS,
        CATEGORY_DEPRECATED,
    }:
        return forced

    mail = data.get("mail") or {}
    if not mail and data.get("sender"):
        mail = data
    key = mail_feedback_key(mail, data)
    fb = normalize_learning(learning or {})

    override = (fb.get("category_overrides") or {}).get(key)
    if isinstance(override, dict):
        cat = str(override.get("category") or "").strip().lower()
        if cat in {
            CATEGORY_IMPORTANT,
            CATEGORY_ACTIONABLE,
            CATEGORY_RELEVANT,
            CATEGORY_LESS,
            CATEGORY_DEPRECATED,
        }:
            return cat

    preferred = preferred_sender_category(fb, str(mail.get("sender") or ""))
    if preferred in {
        CATEGORY_IMPORTANT,
        CATEGORY_ACTIONABLE,
        CATEGORY_RELEVANT,
        CATEGORY_LESS,
        CATEGORY_DEPRECATED,
    }:
        # Gradual sender preference — allow promotion and demotion once votes stick
        return preferred

    # VIP / important-domain mail sits on the top Important row (unless overridden)
    sender = str(mail.get("sender") or "")
    if data.get("rule_force_vip") or is_important_sender(fb, sender):
        return CATEGORY_IMPORTANT
    if sender_matches_important_domain(sender, list(important_domains or [])):
        return CATEGORY_IMPORTANT

    # Strip repeatedly demoted action types before base categorization
    filtered = dict(data)
    filtered["actions"] = actions_for_categorization(data.get("actions") or [], fb)
    categorize = base_categorize or categorize_detected
    return categorize(filtered)


def _name_tokens(name: str) -> Set[str]:
    return {t for t in re.split(r"[^A-Za-z]+", (name or "").lower()) if len(t) >= 2}


def _names_match(candidate: str, known_names: Sequence[str]) -> bool:
    cand = (candidate or "").strip().lower()
    if not cand:
        return False
    cand_tokens = _name_tokens(cand)
    for known in known_names or []:
        known_l = known.strip().lower()
        if not known_l:
            continue
        if cand == known_l or cand in known_l or known_l in cand:
            return True
        known_tokens = _name_tokens(known_l)
        # First-name match: "Julian" vs "Julian Garrett"
        if cand_tokens and known_tokens and (cand_tokens & known_tokens):
            return True
    return False


def extract_greeting_name(body: str) -> str:
    """Return the first-name greeting target from the latest body block, if any."""
    text = (body or "").replace("\r\n", "\n").replace("\r", "\n").strip()
    if not text:
        return ""
    # Prefer the top of the message (newest reply), ignore quoted history later
    head = "\n".join(text.split("\n")[:12])
    match = _GREETING_RE.search(head)
    if not match:
        return ""
    return str(match.group("name") or "").strip()


def ensure_identity_defaults(
    learning: Dict[str, Any],
    *,
    extra_emails: Optional[Iterable[str]] = None,
    extra_names: Optional[Iterable[str]] = None,
) -> Dict[str, Any]:
    """Seed my_emails / my_names without wiping user edits."""
    fb = normalize_learning(learning)
    emails = list(fb["my_emails"])
    seen = set(emails)
    for raw in extra_emails or []:
        addr = normalize_sender_address(str(raw or ""))
        if addr and addr not in seen:
            seen.add(addr)
            emails.append(addr)
    fb["my_emails"] = sorted(emails)

    names = list(fb["my_names"])
    seen_n = {n.lower() for n in names}
    for raw in extra_names or []:
        name = str(raw or "").strip()
        if name and name.lower() not in seen_n:
            seen_n.add(name.lower())
            names.append(name)
    fb["my_names"] = names[:40]
    return fb


def remember_other_name(learning: Dict[str, Any], name: str) -> Dict[str, Any]:
    fb = normalize_learning(learning)
    name = str(name or "").strip()
    if len(name) < 2:
        return fb
    if _names_match(name, fb["my_names"]):
        return fb
    existing = {n.lower() for n in fb["other_names"]}
    if name.lower() not in existing:
        fb["other_names"].append(name)
        fb["other_names"] = fb["other_names"][:80]
    return fb


def set_my_identity(
    learning: Dict[str, Any],
    *,
    names: Optional[Iterable[str]] = None,
    emails: Optional[Iterable[str]] = None,
    filter_not_for_me: Optional[bool] = None,
) -> Dict[str, Any]:
    fb = normalize_learning(learning)
    if names is not None:
        fb["my_names"] = [str(n).strip() for n in names if str(n).strip()]
    if emails is not None:
        cleaned = []
        seen: Set[str] = set()
        for raw in emails:
            addr = normalize_sender_address(str(raw or ""))
            if addr and addr not in seen:
                seen.add(addr)
                cleaned.append(addr)
        fb["my_emails"] = sorted(cleaned)
    if filter_not_for_me is not None:
        fb["filter_deadlines_not_for_me"] = bool(filter_not_for_me)
    return normalize_learning(fb)


def mail_is_for_me(
    mail: Optional[Dict[str, Any]],
    learning: Optional[Dict[str, Any]] = None,
) -> Dict[str, Any]:
    """Decide whether a scraped mail is addressed to the user.

    Returns ``{"for_me": bool, "reason": str, "greeting": str}``.
    When identity is empty, defaults to for_me=True (no filter yet).
    """
    fb = normalize_learning(learning or {})
    mail = mail or {}
    greeting = extract_greeting_name(str(mail.get("body_preview") or ""))

    my_names = list(fb.get("my_names") or [])
    my_emails = set(fb.get("my_emails") or [])
    other_names = list(fb.get("other_names") or [])

    if not my_names and not my_emails:
        return {
            "for_me": True,
            "reason": "identity not configured",
            "greeting": greeting,
        }

    to_blob = " ".join(
        [
            str(mail.get("to") or ""),
            str(mail.get("recipients") or ""),
        ]
    ).lower()
    cc_blob = str(mail.get("cc") or "").lower()
    role = str(mail.get("recipient_role") or "").strip().lower()
    account = normalize_sender_address(str(mail.get("account_smtp") or ""))

    in_to = bool(my_emails) and any(addr in to_blob for addr in my_emails)
    in_cc = bool(my_emails) and any(addr in cc_blob for addr in my_emails)
    account_is_mine = bool(account and account in my_emails)

    if greeting:
        g_low = greeting.lower()
        if g_low in _GROUP_GREETINGS:
            return {
                "for_me": True,
                "reason": f"group greeting ({greeting})",
                "greeting": greeting,
            }
        if _names_match(greeting, my_names):
            return {
                "for_me": True,
                "reason": f"greeting matches you ({greeting})",
                "greeting": greeting,
            }
        if _names_match(greeting, other_names) or (
            my_names and not _names_match(greeting, my_names)
        ):
            # Named greeting to someone else — even if it landed in your mailbox
            return {
                "for_me": False,
                "reason": f"greeting addresses {greeting}, not you",
                "greeting": greeting,
            }

    if in_to or (role == "to" and account_is_mine):
        return {"for_me": True, "reason": "addressed To you", "greeting": greeting}
    if in_cc or role == "cc":
        # Cc alone is weaker; still allow unless greeting said otherwise (handled above)
        return {"for_me": True, "reason": "on Cc", "greeting": greeting}
    if account_is_mine and role in {"", "unknown", "to"}:
        return {
            "for_me": True,
            "reason": "in your mailbox",
            "greeting": greeting,
        }

    # Identity configured but no positive signal — keep rather than over-filter
    return {
        "for_me": True,
        "reason": "no clear not-for-me signal",
        "greeting": greeting,
    }


def should_include_deadline_mail(
    mail: Optional[Dict[str, Any]],
    learning: Optional[Dict[str, Any]] = None,
    *,
    confirmed: bool = False,
    manual: bool = False,
) -> bool:
    """False when deadline lists should hide this mail as not-for-you."""
    if manual or confirmed:
        return True
    fb = normalize_learning(learning or {})
    if not fb.get("filter_deadlines_not_for_me", True):
        return True
    verdict = mail_is_for_me(mail, fb)
    return bool(verdict.get("for_me", True))


def add_user_rule(
    learning: Dict[str, Any],
    text: str,
    *,
    scope: str = "always",
    mail_key: str = "",
    sender: str = "",
    match: Optional[Dict[str, Any]] = None,
    effects: Optional[Dict[str, Any]] = None,
    compiled: bool = False,
) -> Dict[str, Any]:
    """Append a natural-language rule (optionally with compiled match/effects)."""
    fb = normalize_learning(learning)
    text = str(text or "").strip()
    if len(text) < 3:
        return fb
    scope = str(scope or "always").strip().lower()
    if scope not in {"always", "sender", "email"}:
        scope = "always"
    rid = f"rule-{datetime.now().strftime('%Y%m%d%H%M%S')}-{len(fb['user_rules'])+1}"
    eff = normalize_rule_effects(effects or {})
    is_compiled = bool(compiled) and rule_has_enforceable_effects(eff)
    fb["user_rules"].append(
        {
            "id": rid,
            "text": text[:2000],
            "scope": scope,
            "mail_key": str(mail_key or ""),
            "sender": normalize_sender_address(sender) if sender else str(sender or ""),
            "at": datetime.now().isoformat(timespec="seconds"),
            "active": True,
            "compiled": is_compiled,
            "compiled_at": datetime.now().isoformat(timespec="seconds") if is_compiled else "",
            "match": normalize_rule_match(match or {}),
            "effects": eff,
        }
    )
    fb["user_rules"] = fb["user_rules"][-200:]
    return fb


def update_user_rule(
    learning: Dict[str, Any],
    rule_id: str,
    *,
    match: Optional[Dict[str, Any]] = None,
    effects: Optional[Dict[str, Any]] = None,
    compiled: Optional[bool] = None,
    active: Optional[bool] = None,
    text: Optional[str] = None,
) -> Dict[str, Any]:
    """Patch a stored rule by id."""
    fb = normalize_learning(learning)
    rid = str(rule_id or "").strip()
    for rule in fb["user_rules"]:
        if str(rule.get("id") or "") != rid:
            continue
        if text is not None and str(text).strip():
            rule["text"] = str(text).strip()[:2000]
        if match is not None:
            rule["match"] = normalize_rule_match(match)
        if effects is not None:
            rule["effects"] = normalize_rule_effects(effects)
        if active is not None:
            rule["active"] = bool(active)
        if compiled is not None:
            rule["compiled"] = bool(compiled) and rule_has_enforceable_effects(
                rule.get("effects")
            )
            if rule["compiled"]:
                rule["compiled_at"] = datetime.now().isoformat(timespec="seconds")
        break
    return fb


def active_user_rules(
    learning: Optional[Dict[str, Any]] = None,
    *,
    mail_key: str = "",
    sender: str = "",
) -> List[Dict[str, Any]]:
    """Rules that apply globally or to this mail/sender (scope filter only)."""
    fb = normalize_learning(learning or {})
    sender_n = normalize_sender_address(sender) if sender else ""
    mail_key = str(mail_key or "")
    out: List[Dict[str, Any]] = []
    for rule in fb.get("user_rules") or []:
        if not isinstance(rule, dict) or not rule.get("active", True):
            continue
        text = str(rule.get("text") or "").strip()
        if not text:
            continue
        scope = str(rule.get("scope") or "always").lower()
        if scope == "always":
            out.append(rule)
        elif scope == "sender":
            rule_sender = normalize_sender_address(str(rule.get("sender") or ""))
            if rule_sender and sender_n and rule_sender == sender_n:
                out.append(rule)
            elif not rule_sender:
                out.append(rule)
        elif scope == "email":
            if mail_key and str(rule.get("mail_key") or "") == mail_key:
                out.append(rule)
    return out


def format_user_rules_for_llm(
    learning: Optional[Dict[str, Any]] = None,
    *,
    mail_key: str = "",
    sender: str = "",
) -> str:
    """Bullet list of active rules for system/user prompts."""
    rules = active_user_rules(learning, mail_key=mail_key, sender=sender)
    if not rules:
        return ""
    lines = ["User-defined rules (follow these when analysing mail):"]
    for i, rule in enumerate(rules, 1):
        scope = str(rule.get("scope") or "always")
        compiled = "compiled" if rule.get("compiled") else "text-only"
        lines.append(f"{i}. [{scope}|{compiled}] {rule.get('text')}")
    return "\n".join(lines)


def _match_haystack(mail: Dict[str, Any]) -> Dict[str, str]:
    sender = str(mail.get("sender") or "")
    addr = normalize_sender_address(sender)
    domain = addr.rsplit("@", 1)[-1].lower() if "@" in addr else ""
    subject = str(mail.get("subject") or "").lower()
    body = str(mail.get("body_preview") or "").lower()
    return {
        "sender": addr,
        "domain": domain,
        "subject": subject,
        "body": body,
        "blob": f"{subject}\n{body}",
    }


def rule_match_hits(rule: Dict[str, Any], mail: Dict[str, Any]) -> bool:
    """True when compiled matchers hit this mail (or no matchers → scope-only)."""
    match = normalize_rule_match((rule or {}).get("match"))
    hay = _match_haystack(mail or {})
    has_any = any(
        [
            match.get("domains"),
            match.get("senders"),
            match.get("subject_contains"),
            match.get("body_contains"),
            match.get("any_keywords"),
        ]
    )
    if not has_any:
        # Uncompiled / empty match → rely on scope alone
        return True
    if match["senders"] and hay["sender"] in set(match["senders"]):
        return True
    if match["domains"] and hay["domain"]:
        for d in match["domains"]:
            if hay["domain"] == d or hay["domain"].endswith("." + d):
                return True
    for needle in match["subject_contains"]:
        if needle and needle in hay["subject"]:
            return True
    for needle in match["body_contains"]:
        if needle and needle in hay["body"]:
            return True
    for needle in match["any_keywords"]:
        if needle and needle in hay["blob"]:
            return True
    return False


def matching_compiled_rules(
    learning: Optional[Dict[str, Any]],
    mail: Optional[Dict[str, Any]],
    *,
    mail_key: str = "",
) -> List[Dict[str, Any]]:
    """Active compiled rules that scope-match and content-match this mail."""
    mail = mail or {}
    sender = str(mail.get("sender") or "")
    if not mail_key:
        mail_key = mail_feedback_key(mail, {"mail": mail})
    out: List[Dict[str, Any]] = []
    for rule in active_user_rules(learning, mail_key=mail_key, sender=sender):
        if not rule.get("compiled"):
            continue
        if not rule_has_enforceable_effects(rule.get("effects")):
            continue
        if rule_match_hits(rule, mail):
            out.append(rule)
    return out


def apply_user_rules_to_item(
    item: Dict[str, Any],
    learning: Optional[Dict[str, Any]] = None,
) -> Dict[str, Any]:
    """Apply compiled rule effects onto a detection dict (local enforcement)."""
    out = dict(item or {})
    mail = dict(out.get("mail") or {})
    if not mail and out.get("sender"):
        mail = {
            "sender": out.get("sender"),
            "subject": out.get("subject"),
            "body_preview": out.get("body_preview"),
        }
    key = mail_feedback_key(mail, out)
    matched = matching_compiled_rules(learning, mail, mail_key=key)
    explanations: List[str] = []
    actions = [str(a).lower() for a in (out.get("actions") or [])]
    score = int(out.get("score") or 0)
    importance = str(out.get("importance") or "low").lower()
    forced_category = ""
    hide_actions = False
    hide_important = False
    skip_deadline = False
    force_vip = False
    boost_sender = 0

    for rule in matched:
        eff = normalize_rule_effects(rule.get("effects"))
        label = str(rule.get("text") or "")[:80]
        bits: List[str] = []
        if eff.get("category"):
            forced_category = str(eff["category"])
            bits.append(f"category->{forced_category}")
        if eff.get("importance"):
            importance = str(eff["importance"])
            bits.append(f"importance->{importance}")
        if eff.get("boost_score"):
            score += int(eff["boost_score"])
            bits.append(f"score{int(eff['boost_score']):+d}")
        if eff.get("boost_sender"):
            boost_sender += int(eff["boost_sender"])
            bits.append(f"sender{int(eff['boost_sender']):+d}")
        for act in eff.get("demote_actions") or []:
            if act in actions:
                actions = [a for a in actions if a != act]
                bits.append(f"drop:{act}")
        if eff.get("hide_from_actions"):
            hide_actions = True
            bits.append("hide actions")
        if eff.get("hide_from_important"):
            hide_important = True
            bits.append("hide important")
        if eff.get("skip_deadline"):
            skip_deadline = True
            bits.append("skip deadline")
        if eff.get("force_vip"):
            force_vip = True
            bits.append("VIP")
        if bits:
            explanations.append(f"Rule [{label}]: " + ", ".join(bits))

    out["mail"] = mail
    out["actions"] = actions
    out["score"] = max(0, min(100, score))
    out["importance"] = importance
    out["rule_matches"] = [
        {"id": r.get("id"), "text": r.get("text"), "effects": r.get("effects")}
        for r in matched
    ]
    out["rule_explanations"] = explanations
    if forced_category:
        out["rule_forced_category"] = forced_category
    if hide_actions:
        out["rule_hide_from_actions"] = True
        out["actions"] = []
    if hide_important:
        out["rule_hide_from_important"] = True
    if skip_deadline:
        out["rule_skip_deadline"] = True
        out["deadlines"] = []
    if force_vip:
        out["rule_force_vip"] = True
    if boost_sender:
        out["rule_boost_sender"] = boost_sender
    return out


def apply_user_rules_to_items(
    items: List[Dict[str, Any]],
    learning: Optional[Dict[str, Any]] = None,
) -> List[Dict[str, Any]]:
    return [apply_user_rules_to_item(it, learning) for it in (items or [])]


def summarize_rule_effects(effects: Optional[Dict[str, Any]]) -> str:
    eff = normalize_rule_effects(effects or {})
    bits: List[str] = []
    if eff.get("category"):
        bits.append(f"category={eff['category']}")
    if eff.get("importance"):
        bits.append(f"importance={eff['importance']}")
    if eff.get("boost_score"):
        bits.append(f"score{int(eff['boost_score']):+d}")
    if eff.get("boost_sender"):
        bits.append(f"sender{int(eff['boost_sender']):+d}")
    if eff.get("demote_actions"):
        bits.append("drop " + ",".join(eff["demote_actions"]))
    if eff.get("hide_from_actions"):
        bits.append("hide from Actions")
    if eff.get("hide_from_important"):
        bits.append("hide from Important")
    if eff.get("skip_deadline"):
        bits.append("skip deadlines")
    if eff.get("force_vip"):
        bits.append("force VIP colour")
    return "; ".join(bits) if bits else "(no local effects)"
