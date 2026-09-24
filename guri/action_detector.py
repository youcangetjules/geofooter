#!/usr/bin/env python3
"""Heuristic importance + action detection for scraped Outlook mail."""

from __future__ import annotations

import re
from dataclasses import asdict, dataclass, field
from datetime import datetime
from typing import Any, Dict, Iterable, List, Optional, Sequence, Set

from guri.outlook_scraper import ScrapedMail

# Adapted from AES body_scanner urgency cues.
URGENCY_PATTERNS = [
    r"\burgent\b",
    r"\bimmediately\b",
    r"\basap\b",
    r"\beod\b",
    r"\bclose of (?:business|play)\b",
    r"\bwithin\s+\d+\s+(?:hours?|days?)\b",
    r"\bact\s+now\b",
    r"\bfinal\s+notice\b",
    r"\bexpires?\s+today\b",
    r"\bby\s+(?:eod|cob|friday|monday|tuesday|wednesday|thursday|saturday|sunday)\b",
    r"\bdeadline\b",
]

# Stronger cues that something is due (for Deadlines tab extraction).
DEADLINE_CUE_PATTERNS = [
    r"\bdeadline\b",
    r"\bdue\s+date\b",
    r"\bdue\s+by\b",
    r"\bdue\s+on\b",
    r"\bpayment\s+due\b",
    r"\bsubmit(?:ted)?\s+by\b",
    r"\brespond\s+by\b",
    r"\breply\s+by\b",
    r"\bexpires?\b",
    r"\bexpir(?:y|ation)\b",
    r"\bclose of (?:business|play)\b",
    r"\beod\b",
    r"\bcob\b",
    r"\blast\s+chance\b",
    r"\bfinal\s+(?:notice|reminder|call)\b",
    r"\bmust\s+be\s+(?:received|submitted|returned)\s+by\b",
]

# Date / relative-due fragments often paired with deadline language.
DEADLINE_DATE_RE = re.compile(
    r"(?P<label>"
    r"\d{4}-\d{2}-\d{2}"
    r"|\d{1,2}[/-]\d{1,2}[/-]\d{2,4}"
    r"|\d{1,2}\s+(?:Jan(?:uary)?|Feb(?:ruary)?|Mar(?:ch)?|Apr(?:il)?|May|Jun(?:e)?|"
    r"Jul(?:y)?|Aug(?:ust)?|Sep(?:t(?:ember)?)?|Oct(?:ober)?|Nov(?:ember)?|Dec(?:ember)?)"
    r"\.?\s+\d{2,4}"
    r"|(?:Mon|Tue|Wed|Thu|Fri|Sat|Sun)[a-z]*"
    r"|(?:today|tomorrow|tonight)"
    r"|(?:this|next)\s+(?:week|monday|tuesday|wednesday|thursday|friday|saturday|sunday)"
    r"|(?:eod|cob)"
    r")",
    re.IGNORECASE,
)

REPLY_PATTERNS = [
    r"\bplease\s+reply\b",
    r"\bkindly\s+reply\b",
    r"\blooking\s+forward\s+to\s+(?:your\s+)?(?:reply|response)\b",
    r"\bawaiting\s+(?:your\s+)?(?:reply|response|feedback)\b",
    r"\bcan\s+you\s+(?:please\s+)?(?:confirm|advise|let\s+me\s+know)\b",
    r"\bplease\s+(?:confirm|advise|advise|respond)\b",
    r"\blet\s+me\s+know\b",
    r"\bneed\s+(?:your\s+)?(?:input|feedback|thoughts)\b",
    r"\bwaiting\s+on\s+you\b",
]

APPROVE_PATTERNS = [
    r"\bplease\s+approve\b",
    r"\bapproval\s+required\b",
    r"\bfor\s+your\s+approval\b",
    r"\bsign\s+(?:and\s+)?return\b",
    r"\bplease\s+sign\b",
    r"\bcounter[- ]?sign\b",
    r"\bauthori[sz]e\b",
]

RSVP_PATTERNS = [
    r"\brsvp\b",
    r"\bplease\s+(?:accept|decline)\b",
    r"\bmeeting\s+invitation\b",
    r"\byou('re|\s+are)\s+invited\b",
]

REVIEW_PATTERNS = [
    r"\bfor\s+your\s+(?:review|attention)\b",
    r"\bplease\s+review\b",
    r"\baction\s+required\b",
    r"\baction\s+needed\b",
    r"\bfyi\s*[—\-:]?\s*action\b",
]

QUESTION_RE = re.compile(r"\?")
RE_FW_RE = re.compile(r"^\s*(?:re|fw|fwd)\s*:", re.IGNORECASE)

# Combined patterns used to highlight action language in previews.
ACTION_HIGHLIGHT_PATTERNS: Sequence[str] = (
    *URGENCY_PATTERNS,
    *DEADLINE_CUE_PATTERNS,
    *REPLY_PATTERNS,
    *APPROVE_PATTERNS,
    *RSVP_PATTERNS,
    *REVIEW_PATTERNS,
)


def highlight_action_html(text: str) -> str:
    """Escape text and wrap action-language matches for QTextEdit HTML preview."""
    import html as html_mod

    raw = text or ""
    if not raw.strip():
        return "<i style='color:#888'>(No body text available)</i>"

    # Work on escaped text; match against original spans via iterative scan on raw.
    matches: List[tuple] = []
    for pat in ACTION_HIGHLIGHT_PATTERNS:
        for m in re.finditer(pat, raw, re.IGNORECASE):
            matches.append((m.start(), m.end()))
    # Also highlight standalone question marks in subject-like short lines
    for m in re.finditer(r"[^\n?]{0,80}\?", raw):
        q = raw.rfind("?", m.start(), m.end() + 1)
        if q >= 0:
            # highlight the question clause (last 60 chars up to ?)
            start = max(m.start(), q - 60)
            matches.append((start, q + 1))

    if not matches:
        return (
            "<pre style='white-space:pre-wrap;font-family:Segoe UI,sans-serif;"
            f"font-size:12px;color:#1a2e35;'>{html_mod.escape(raw)}</pre>"
        )

    # Merge overlapping ranges
    matches.sort()
    merged: List[List[int]] = []
    for start, end in matches:
        if not merged or start > merged[-1][1]:
            merged.append([start, end])
        else:
            merged[-1][1] = max(merged[-1][1], end)

    parts: List[str] = []
    cursor = 0
    for start, end in merged:
        if cursor < start:
            parts.append(html_mod.escape(raw[cursor:start]))
        snippet = html_mod.escape(raw[start:end])
        parts.append(
            "<mark style='background:#ffe08a;color:#1a1a1a;padding:0 2px;"
            f"border-radius:2px;'>{snippet}</mark>"
        )
        cursor = end
    if cursor < len(raw):
        parts.append(html_mod.escape(raw[cursor:]))

    return (
        "<div style='white-space:pre-wrap;font-family:Segoe UI,sans-serif;"
        f"font-size:12px;color:#1a2e35;line-height:1.45;'>{''.join(parts)}</div>"
    )


@dataclass
class DetectedMail:
    mail: Dict[str, Any]
    importance: str  # low | medium | high
    score: int
    actions: List[str] = field(default_factory=list)
    reasons: List[str] = field(default_factory=list)
    status: str = "—"
    deadlines: List[Dict[str, Any]] = field(default_factory=list)

    def to_dict(self) -> Dict[str, Any]:
        return {
            "mail": self.mail,
            "importance": self.importance,
            "score": self.score,
            "actions": list(self.actions),
            "reasons": list(self.reasons),
            "status": self.status,
            "deadlines": [dict(d) for d in self.deadlines],
        }


def _sender_domain(sender: str) -> str:
    sender = (sender or "").lower().strip()
    if "@" in sender:
        return sender.rsplit("@", 1)[-1].strip(">")
    return ""


def sender_matches_important_domain(
    sender: str,
    important_domains: Optional[Iterable[str]] = None,
) -> bool:
    """True when sender's email domain matches a configured important domain."""
    domains = {d.strip().lower() for d in (important_domains or []) if d and d.strip()}
    if not domains:
        return False
    domain = _sender_domain(sender)
    if not domain:
        # Fall back to substring only when address parsing failed
        blob = (sender or "").lower()
        return any(d in blob for d in domains)
    return any(domain == d or domain.endswith("." + d) for d in domains)


def _any_match(patterns: Sequence[str], text: str) -> Optional[str]:
    for pat in patterns:
        if re.search(pat, text, re.IGNORECASE):
            return pat
    return None


def _parse_deadline_label(label: str, *, now: Optional[datetime] = None) -> Optional[str]:
    """Best-effort ISO date (YYYY-MM-DD) from a deadline label, else None."""
    from datetime import datetime as _dt
    from datetime import timedelta

    now = now or _dt.now()
    raw = (label or "").strip()
    if not raw:
        return None
    low = raw.lower()

    if low in {"today", "tonight", "eod", "cob"}:
        return now.date().isoformat()
    if low == "tomorrow":
        return (now.date() + timedelta(days=1)).isoformat()
    if low.startswith("this week"):
        return (now.date() + timedelta(days=(4 - now.weekday()) % 7 or 7)).isoformat()
    if low.startswith("next week"):
        return (now.date() + timedelta(days=7 + ((4 - now.weekday()) % 7 or 7))).isoformat()

    weekdays = {
        "monday": 0,
        "tuesday": 1,
        "wednesday": 2,
        "thursday": 3,
        "friday": 4,
        "saturday": 5,
        "sunday": 6,
        "mon": 0,
        "tue": 1,
        "wed": 2,
        "thu": 3,
        "fri": 4,
        "sat": 5,
        "sun": 6,
    }
    for name, idx in weekdays.items():
        if low.startswith(name) or f" {name}" in low or low.endswith(name):
            if name in low.split() or low.startswith(name):
                delta = (idx - now.weekday()) % 7
                if delta == 0:
                    delta = 7
                if "next" in low:
                    delta += 7
                return (now.date() + timedelta(days=delta)).isoformat()

    for fmt in (
        "%Y-%m-%d",
        "%d/%m/%Y",
        "%d/%m/%y",
        "%d-%m-%Y",
        "%d-%m-%y",
        "%m/%d/%Y",
        "%m/%d/%y",
        "%d %b %Y",
        "%d %B %Y",
        "%d %b %y",
        "%d %B %y",
    ):
        try:
            return _dt.strptime(raw.replace(".", ""), fmt).date().isoformat()
        except ValueError:
            continue
    return None


def extract_deadlines(
    text: str,
    *,
    now: Optional[datetime] = None,
    extra_cues: Optional[Sequence[str]] = None,
) -> List[Dict[str, Any]]:
    """Find deadline cues (+ nearby dates) in subject/body text.

    Returns at most one hit per distinct due-date (or cue when undated),
    so a single mail does not explode into many near-identical rows.

    ``extra_cues`` — learned regex fragments from user corrections.
    """
    from datetime import datetime as _dt

    now = now or _dt.now()
    blob = text or ""
    hits: List[Dict[str, Any]] = []
    seen: Set[str] = set()
    patterns: List[str] = list(DEADLINE_CUE_PATTERNS)
    for cue in extra_cues or []:
        cue = str(cue or "").strip()
        if cue and cue not in patterns:
            patterns.append(cue)

    for pat in patterns:
        try:
            iterator = re.finditer(pat, blob, re.IGNORECASE)
        except re.error:
            continue
        for m in iterator:
            window = blob[max(0, m.start() - 20) : min(len(blob), m.end() + 48)]
            window_clean = re.sub(r"\s+", " ", window).strip()
            date_m = DEADLINE_DATE_RE.search(blob[m.start() : m.end() + 60])
            if not date_m:
                date_m = DEADLINE_DATE_RE.search(blob[max(0, m.start() - 40) : m.end()])
            label = (date_m.group("label") if date_m else "").strip()
            due = _parse_deadline_label(label, now=now) if label else None
            # Prefer due-date identity so "deadline" + "submit by" same day → one hit
            key = due or f"{m.group(0).lower()}|{label.lower()}"
            if key in seen:
                continue
            seen.add(key)
            hits.append(
                {
                    "cue": m.group(0),
                    "label": label or m.group(0),
                    "due": due,
                    "snippet": window_clean[:120],
                }
            )

    # Urgency phrases that imply a near-term deadline even without the word "deadline"
    if not hits and _any_match(URGENCY_PATTERNS, blob):
        date_m = DEADLINE_DATE_RE.search(blob)
        label = (date_m.group("label") if date_m else "urgent").strip()
        hits.append(
            {
                "cue": "urgency",
                "label": label,
                "due": _parse_deadline_label(label, now=now),
                "snippet": re.sub(r"\s+", " ", blob[:120]).strip(),
            }
        )
    return hits


def merge_deadline_hits(deadlines: Sequence[Dict[str, Any]]) -> Optional[Dict[str, Any]]:
    """Collapse multiple cue hits from one mail into a single display record."""
    hits = [dict(d) for d in (deadlines or []) if isinstance(d, dict)]
    if not hits:
        return None
    dated = [d for d in hits if d.get("due")]
    if dated:
        best = min(dated, key=lambda d: str(d.get("due") or "9999"))
    else:
        best = hits[0]
    cues: List[str] = []
    for d in hits:
        cue = str(d.get("cue") or "").strip()
        if cue and cue.lower() not in {c.lower() for c in cues}:
            cues.append(cue)
    snippet = str(best.get("snippet") or "")
    return {
        "cue": ", ".join(cues) if cues else str(best.get("cue") or "deadline"),
        "label": str(best.get("label") or best.get("cue") or "Deadline"),
        "due": best.get("due"),
        "snippet": snippet[:160],
        "all": hits,
    }

def detect_mail_actions(
    mail: ScrapedMail,
    *,
    important_domains: Optional[Iterable[str]] = None,
    important_senders: Optional[Iterable[str]] = None,
    extra_deadline_cues: Optional[Sequence[str]] = None,
    allow_responses: bool = True,
) -> DetectedMail:
    """Score a scraped mail and suggest action types."""
    domains = {d.strip().lower() for d in (important_domains or []) if d and d.strip()}
    vip_senders = {
        str(s).strip().lower()
        for s in (important_senders or [])
        if str(s).strip()
    }
    subject = mail.subject or ""
    body = mail.body_preview or ""
    blob = f"{subject}\n{body}"
    score = 0
    actions: List[str] = []
    reasons: List[str] = []

    domain = _sender_domain(mail.sender)
    if sender_matches_important_domain(mail.sender, domains):
        score += 25
        reasons.append(f"Important domain ({domain or mail.sender})")

    sender_addr = ""
    m_addr = re.search(r"<([^>]+)>", mail.sender or "")
    if m_addr:
        sender_addr = m_addr.group(1).strip().lower()
    elif "@" in (mail.sender or ""):
        sender_addr = (mail.sender or "").strip("<>\"' ").lower()
    if sender_addr and sender_addr in vip_senders:
        score += 28
        reasons.append(f"Important sender ({sender_addr})")

    if mail.importance >= 2:
        score += 15
        reasons.append("Outlook high importance")
    if mail.flag_status and int(mail.flag_status) > 0:
        score += 12
        reasons.append("Outlook flagged")
    if mail.unread and mail.recipient_role == "to":
        score += 8
        reasons.append("Unread and addressed To you")
    elif mail.unread:
        score += 3
        reasons.append("Unread")

    if mail.is_meeting_request or mail.message_class.lower().startswith("ipm.schedule.meeting"):
        score += 20
        actions.append("rsvp")
        reasons.append("Meeting request")

    if _any_match(RSVP_PATTERNS, blob):
        if "rsvp" not in actions:
            actions.append("rsvp")
        score += 12
        reasons.append("RSVP / invitation language")

    if _any_match(APPROVE_PATTERNS, blob):
        actions.append("approve")
        score += 18
        reasons.append("Approval / signature requested")

    if _any_match(REPLY_PATTERNS, blob):
        actions.append("reply")
        score += 14
        reasons.append("Reply / response requested")

    if _any_match(REVIEW_PATTERNS, blob):
        actions.append("review")
        score += 12
        reasons.append("Review / action required")

    deadlines = extract_deadlines(blob, extra_cues=extra_deadline_cues)
    learned_hit = bool(extra_deadline_cues) and any(
        re.search(pat, blob, re.IGNORECASE)
        for pat in (extra_deadline_cues or [])
        if pat
    )
    if (
        deadlines
        or _any_match(URGENCY_PATTERNS, blob)
        or _any_match(DEADLINE_CUE_PATTERNS, blob)
        or learned_hit
    ):
        if "deadline" not in actions:
            actions.append("deadline")
        score += 16
        if deadlines:
            labels = ", ".join(
                str(d.get("due") or d.get("label") or d.get("cue") or "") for d in deadlines[:3]
            )
            reasons.append(f"Deadline signal ({labels})" if labels else "Deadline language")
        elif learned_hit:
            reasons.append("Learned deadline cue")
        else:
            reasons.append("Urgency / deadline language")

    # Subject question or multiple body questions → likely needs reply
    if QUESTION_RE.search(subject) or len(QUESTION_RE.findall(body)) >= 2:
        if "reply" not in actions:
            actions.append("reply")
        score += 8
        reasons.append("Question(s) directed at you")

    if mail.recipient_role == "cc":
        # Cc is usually weaker unless explicit action language already found
        if not actions:
            score = max(0, score - 5)
            reasons.append("On Cc (lower priority unless action language)")
        else:
            reasons.append("On Cc but action language present")

    if not allow_responses and RE_FW_RE.match(subject):
        # Account opted out of Responses — demote thread noise
        score = max(0, score - 10)
        reasons.append("Responses disabled for account; demoted Re:/Fw:")

    # Deduplicate actions preserving order
    seen = set()
    uniq_actions: List[str] = []
    for a in actions:
        if a not in seen:
            seen.add(a)
            uniq_actions.append(a)

    if score >= 35 or (uniq_actions and score >= 20):
        importance = "high"
    elif score >= 15 or uniq_actions:
        importance = "medium"
    else:
        importance = "low"

    return DetectedMail(
        mail=mail.to_dict(),
        importance=importance,
        score=score,
        actions=uniq_actions,
        reasons=reasons,
        deadlines=deadlines,
    )


def detect_batch(
    mails: Sequence[ScrapedMail],
    *,
    important_domains: Optional[Iterable[str]] = None,
    important_senders: Optional[Iterable[str]] = None,
    extra_deadline_cues: Optional[Sequence[str]] = None,
    account_response_flags: Optional[Dict[str, bool]] = None,
) -> List[DetectedMail]:
    """Run detection on a batch; sort by score then received time."""
    account_response_flags = account_response_flags or {}
    out: List[DetectedMail] = []
    for mail in mails:
        allow = account_response_flags.get(mail.account_smtp.lower(), True)
        # Also try store-keyed flags if smtp missing
        if mail.account_smtp.lower() not in account_response_flags:
            allow = account_response_flags.get(mail.store_id, allow)
        out.append(
            detect_mail_actions(
                mail,
                important_domains=important_domains,
                important_senders=important_senders,
                extra_deadline_cues=extra_deadline_cues,
                allow_responses=allow,
            )
        )
    out.sort(
        key=lambda d: (
            d.score,
            str(d.mail.get("received") or ""),
        ),
        reverse=True,
    )
    return out


def filter_important(detected: Sequence[DetectedMail], *, min_score: int = 15) -> List[DetectedMail]:
    return [d for d in detected if d.score >= min_score or d.importance != "low"]


def filter_actions(detected: Sequence[DetectedMail]) -> List[DetectedMail]:
    return [d for d in detected if d.actions]


# Timeline / triage buckets (mutually exclusive)
CATEGORY_IMPORTANT = "important"
CATEGORY_ACTIONABLE = "actionable"
CATEGORY_RELEVANT = "relevant"
CATEGORY_LESS = "less"
CATEGORY_DEPRECATED = "deprecated"

TIMELINE_CATEGORIES = (
    CATEGORY_IMPORTANT,
    CATEGORY_ACTIONABLE,
    CATEGORY_RELEVANT,
    CATEGORY_LESS,
    CATEGORY_DEPRECATED,
)

TIMELINE_CATEGORY_LABELS = {
    CATEGORY_IMPORTANT: "Important",
    CATEGORY_ACTIONABLE: "Actionable",
    CATEGORY_RELEVANT: "Relevant but not actionable",
    CATEGORY_LESS: "Less relevant",
    CATEGORY_DEPRECATED: "Deprecated",
}


def categorize_detected(detected: Any) -> str:
    """Map a DetectedMail or detection dict into one timeline category."""
    if isinstance(detected, DetectedMail):
        actions = detected.actions
        score = detected.score
        importance = detected.importance
    else:
        data = detected or {}
        actions = data.get("actions") or []
        score = int(data.get("score") or 0)
        importance = str(data.get("importance") or "low")

    if actions:
        return CATEGORY_ACTIONABLE
    if score >= 15 or importance in {"medium", "high"}:
        return CATEGORY_RELEVANT
    return CATEGORY_LESS
