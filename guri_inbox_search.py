#!/usr/bin/env python3
"""Conversation-level inbox search: sentiment, subject matter, and theme.

Scores Outlook threads as a whole — not isolated keyword hits on one message.
Heuristic ranking always runs; Ollama can rerank the top conversations.
"""

from __future__ import annotations

import json
import logging
import os
import re
from collections import Counter, defaultdict
from dataclasses import asdict, dataclass, field
from datetime import datetime
from pathlib import Path
from typing import Any, Callable, Dict, Iterable, List, Optional, Sequence, Tuple

from guri_outlook_scraper import (
    ScrapedMail,
    expand_conversations,
    scrape_inbox_for_search,
)

logger = logging.getLogger(__name__)

ProgressFn = Optional[Callable[[str], None]]

STOPWORDS = {
    "a", "an", "the", "and", "or", "but", "if", "to", "of", "in", "on", "for",
    "from", "with", "at", "by", "as", "is", "are", "was", "were", "be", "been",
    "it", "this", "that", "these", "those", "i", "me", "my", "we", "our", "you",
    "your", "they", "them", "their", "he", "she", "his", "her", "not", "no",
    "so", "than", "then", "too", "very", "can", "could", "would", "should",
    "will", "just", "about", "into", "over", "after", "before", "also", "any",
    "all", "some", "more", "most", "other", "only", "own", "same", "such",
    "email", "emails", "mail", "mails", "message", "messages", "inbox",
    "thread", "threads", "conversation", "conversations", "find", "show",
    "search", "looking", "look", "please", "need", "get", "got", "see",
    "ones", "stuff", "thing", "things", "re", "fw", "fwd",
}

# Query language → requested sentiment (not the mail lexicon).
SENTIMENT_CUES: Dict[str, Tuple[str, ...]] = {
    "angry": ("angry", "anger", "furious", "outraged", "livid", "hostile"),
    "frustrated": (
        "frustrated", "frustration", "annoyed", "irritat", "fed up", "fed-up",
    ),
    "negative": (
        "negative", "unhappy", "upset", "complaint", "complaining",
        "dissatisfied", "unhappy", "bad news", "blame",
    ),
    "positive": (
        "positive", "happy", "pleased", "glad", "good news", "upbeat",
        "optimistic", "complimentary", "praise",
    ),
    "grateful": (
        "grateful", "thankful", "thanks", "thank you", "appreciation",
        "appreciate",
    ),
    "urgent": (
        "urgent", "urgency", "asap", "immediately", "pressing", "time-critical",
        "time critical",
    ),
    "worried": (
        "worried", "worry", "concerned", "concern", "anxious", "nervous",
        "uneasy", "risk",
    ),
    "confused": (
        "confused", "unclear", "ambigu", "don't understand", "clarif",
        "puzzled",
    ),
    "polite": ("polite", "courteous", "formal tone", "diplomatic"),
}

# Mail-body lexicons (compact; no extra NLP packages).
MAIL_LEXICONS: Dict[str, Tuple[str, ...]] = {
    "angry": (
        "unacceptable", "disgrace", "outraged", "furious", "angry",
        "appalling", "ridiculous", "incompetent",
    ),
    "frustrated": (
        "frustrated", "frustrating", "still waiting", "again and again",
        "not acceptable", "going round in circles", "chasing this",
        "why hasn't", "still not", "keeps happening",
    ),
    "negative": (
        "unfortunately", "disappointed", "disappointing", "complaint",
        "unhappy", "problem with", "does not work", "failed", "failure",
        "unable to", "missed", "overdue", "late payment",
    ),
    "positive": (
        "great news", "well done", "excellent", "pleased", "delighted",
        "congratulations", "looking good", "happy to", "glad to",
        "successfully", "approved",
    ),
    "grateful": (
        "thank you", "thanks", "much appreciated", "grateful",
        "appreciate your", "kind regards",
    ),
    "urgent": (
        "urgent", "asap", "immediately", "as soon as possible", "eod",
        "close of business", "today please", "by return", "critical",
    ),
    "worried": (
        "concerned", "worried", "at risk", "risk that", "afraid that",
        "might miss", "slipping", "not confident",
    ),
    "confused": (
        "not sure", "unclear", "confused", "can you clarify",
        "don't understand", "what do you mean", "which one",
    ),
}

ABOUT_RE = re.compile(
    r"\b(?:about|regarding|concerning|re(?:gards?)?(?:\s+the)?|"
    r"on the (?:topic|subject) of)\s+(.+)$",
    re.IGNORECASE,
)
TOKEN_RE = re.compile(r"[a-z0-9][a-z0-9+\-./]*[a-z0-9]|[a-z0-9]", re.IGNORECASE)


def _local_geofooter() -> Path:
    local = os.environ.get("LOCALAPPDATA", "")
    if local:
        return Path(local) / "GeoFooter"
    return Path(r"C:\GeoFooter")


def settings_path() -> Path:
    return _local_geofooter() / "guri_inbox_search.json"


def load_inbox_search_settings() -> Dict[str, Any]:
    path = settings_path()
    defaults = {
        "store_id": "",
        "smtp": "",
        "days": 30,
        "max_messages": 120,
        "include_sent": True,
        "use_ollama": True,
        "last_query": "",
    }
    try:
        if path.is_file():
            raw = json.loads(path.read_text(encoding="utf-8"))
            if isinstance(raw, dict):
                defaults.update({k: raw[k] for k in defaults if k in raw})
    except Exception:
        pass
    try:
        defaults["days"] = max(1, min(180, int(defaults.get("days") or 30)))
    except (TypeError, ValueError):
        defaults["days"] = 30
    try:
        defaults["max_messages"] = max(10, min(400, int(defaults.get("max_messages") or 120)))
    except (TypeError, ValueError):
        defaults["max_messages"] = 120
    defaults["include_sent"] = bool(defaults.get("include_sent", True))
    defaults["use_ollama"] = bool(defaults.get("use_ollama", True))
    return defaults


def save_inbox_search_settings(settings: Dict[str, Any]) -> Path:
    path = settings_path()
    path.parent.mkdir(parents=True, exist_ok=True)
    payload = {
        "store_id": str(settings.get("store_id") or ""),
        "smtp": str(settings.get("smtp") or ""),
        "days": max(1, min(180, int(settings.get("days") or 30))),
        "max_messages": max(10, min(400, int(settings.get("max_messages") or 120))),
        "include_sent": bool(settings.get("include_sent", True)),
        "use_ollama": bool(settings.get("use_ollama", True)),
        "last_query": str(settings.get("last_query") or "")[:2000],
        "updated": datetime.now().isoformat(timespec="seconds"),
    }
    path.write_text(json.dumps(payload, indent=2), encoding="utf-8")
    return path


def _normalize_token(word: str) -> str:
    w = (word or "").strip().lower()
    if w.endswith("'s"):
        w = w[:-2]
    if len(w) > 4 and w.endswith("s") and not w.endswith("ss"):
        if w.endswith("ies"):
            w = w[:-3] + "y"
        elif w.endswith(("ses", "xes", "zes", "ches", "shes")):
            w = w[:-2]
        else:
            w = w[:-1]
    return w


def tokenize(text: str) -> List[str]:
    return [_normalize_token(t) for t in TOKEN_RE.findall(text or "") if t]


def content_tokens(text: str) -> List[str]:
    return [t for t in tokenize(text) if t and t not in STOPWORDS and len(t) >= 2]


@dataclass
class SearchQuery:
    raw: str
    keywords: List[str] = field(default_factory=list)
    topics: List[str] = field(default_factory=list)
    sentiments: List[str] = field(default_factory=list)
    about: str = ""

    def to_dict(self) -> Dict[str, Any]:
        return asdict(self)

    def is_empty(self) -> bool:
        return not (self.keywords or self.topics or self.sentiments or self.about)


def parse_search_query(raw: str) -> SearchQuery:
    """Turn a free-text request into sentiment + subject-matter + keywords."""
    text = (raw or "").strip()
    lower = text.lower()
    sentiments: List[str] = []
    used_spans: List[str] = []
    for label, cues in SENTIMENT_CUES.items():
        for cue in cues:
            if cue in lower:
                if label not in sentiments:
                    sentiments.append(label)
                used_spans.append(cue)

    about = ""
    about_m = ABOUT_RE.search(text)
    if about_m:
        about = about_m.group(1).strip(" .,?!")

    remainder = lower
    for span in used_spans:
        remainder = remainder.replace(span, " ")
    remainder = ABOUT_RE.sub(" ", remainder)
    remainder = re.sub(
        r"\b(?:emails?|mails?|messages?|threads?|conversations?|inbox)\b",
        " ",
        remainder,
    )
    topics = content_tokens(about) if about else []
    if not topics:
        topics = content_tokens(remainder)
    keywords = list(dict.fromkeys(topics + content_tokens(text)))
    # Sentiment-only words should not dominate keyword matching
    sentiment_tokens = set()
    for cues in SENTIMENT_CUES.values():
        for cue in cues:
            sentiment_tokens.update(content_tokens(cue))
    keywords = [k for k in keywords if k not in sentiment_tokens]
    if not keywords and topics:
        keywords = list(topics)
    return SearchQuery(
        raw=text,
        keywords=keywords[:24],
        topics=topics[:16],
        sentiments=sentiments,
        about=about,
    )


def _mail_blob(mail: ScrapedMail) -> str:
    return " ".join(
        [
            mail.subject or "",
            mail.conversation or "",
            mail.body_preview or "",
            mail.sender or "",
        ]
    )


def score_mail_sentiment(text: str) -> Dict[str, float]:
    """Return per-label hit counts normalised 0–1."""
    blob = (text or "").lower()
    scores: Dict[str, float] = {}
    if not blob.strip():
        return scores
    for label, phrases in MAIL_LEXICONS.items():
        hits = 0
        for phrase in phrases:
            if phrase in blob:
                hits += 1
        if hits:
            scores[label] = min(1.0, hits / 3.0)
    return scores


def dominant_sentiment(scores: Dict[str, float]) -> str:
    if not scores:
        return "neutral"
    ranked = sorted(scores.items(), key=lambda kv: kv[1], reverse=True)
    top_label, top_val = ranked[0]
    if top_val < 0.25:
        return "neutral"
    if len(ranked) > 1 and ranked[1][1] >= top_val * 0.85 and ranked[1][0] != top_label:
        pos = {"positive", "grateful"}
        neg = {"angry", "frustrated", "negative", "worried"}
        labels = {ranked[0][0], ranked[1][0]}
        if labels & pos and labels & neg:
            return "mixed"
    return top_label


def conversation_key(mail: ScrapedMail) -> str:
    cid = (mail.conversation_id or "").strip()
    if cid:
        return f"id:{cid}"
    topic = (mail.conversation or mail.subject or "").strip().lower()
    topic = re.sub(r"^\s*(?:re|fw|fwd)\s*:\s*", "", topic, flags=re.IGNORECASE)
    topic = re.sub(r"\s+", " ", topic)
    if topic:
        return f"topic:{topic}"
    return f"item:{mail.entry_id}"


@dataclass
class ConversationHit:
    conversation_id: str
    topic: str
    sentiment: str
    sentiment_scores: Dict[str, float]
    participants: List[str]
    message_count: int
    latest: str
    score: float
    reasons: List[str]
    subject_matter: List[str]
    messages: List[Dict[str, Any]]
    summary: str = ""
    llm_why: str = ""

    def to_dict(self) -> Dict[str, Any]:
        return asdict(self)


def group_conversations(mails: Sequence[ScrapedMail]) -> List[List[ScrapedMail]]:
    buckets: Dict[str, List[ScrapedMail]] = defaultdict(list)
    order: List[str] = []
    for mail in mails:
        key = conversation_key(mail)
        if key not in buckets:
            order.append(key)
        buckets[key].append(mail)
    grouped: List[List[ScrapedMail]] = []
    for key in order:
        thread = buckets[key]
        thread.sort(key=lambda m: m.received_dt or datetime.min)
        grouped.append(thread)
    return grouped


def _participant_name(mail: ScrapedMail) -> str:
    sender = (mail.sender or "").strip()
    if sender:
        return sender
    return mail.account_smtp or "unknown"


def _thread_subject_matter(thread: Sequence[ScrapedMail], limit: int = 8) -> List[str]:
    counts: Counter = Counter()
    for mail in thread:
        counts.update(content_tokens(_mail_blob(mail)))
    common = [w for w, _n in counts.most_common(24) if w not in STOPWORDS]
    return common[:limit]


def _keyword_coverage(tokens: Sequence[str], keywords: Sequence[str]) -> float:
    if not keywords:
        return 0.0
    hay = set(tokens)
    hits = 0
    for kw in keywords:
        if kw in hay:
            hits += 1
            continue
        if any(kw in t or t in kw for t in hay if len(t) >= 3 and len(kw) >= 3):
            hits += 1
    return hits / max(1, len(keywords))


def _messages_hitting_keywords(
    thread: Sequence[ScrapedMail],
    keywords: Sequence[str],
) -> int:
    if not keywords:
        return 0
    n = 0
    for mail in thread:
        blob_tokens = content_tokens(_mail_blob(mail))
        if _keyword_coverage(blob_tokens, keywords) > 0:
            n += 1
    return n


def score_conversation(
    thread: Sequence[ScrapedMail],
    query: SearchQuery,
) -> ConversationHit:
    blobs = [_mail_blob(m) for m in thread]
    combined = "\n".join(blobs)
    tokens = content_tokens(combined)
    sent_scores = score_mail_sentiment(combined)
    sentiment = dominant_sentiment(sent_scores)
    topics = _thread_subject_matter(thread)
    latest_mail = max(thread, key=lambda m: m.received_dt or datetime.min)
    topic_label = (
        latest_mail.conversation
        or latest_mail.subject
        or "(no subject)"
    )
    topic_label = re.sub(r"^\s*(?:re|fw|fwd)\s*:\s*", "", topic_label, flags=re.IGNORECASE)

    reasons: List[str] = []
    kw_score = _keyword_coverage(tokens, query.keywords)
    topic_score = _keyword_coverage(tokens, query.topics)
    if query.about:
        about_tokens = content_tokens(query.about)
        about_score = _keyword_coverage(tokens, about_tokens)
        topic_score = max(topic_score, about_score)

    spread = 0.0
    if query.keywords:
        hitting = _messages_hitting_keywords(thread, query.keywords)
        spread = min(1.0, hitting / max(2, min(len(thread), 4)))
        if hitting >= 2:
            reasons.append(f"Theme appears across {hitting} messages in the thread")
        elif hitting == 1:
            reasons.append("Topic found in one message — weaker conversation match")

    sent_align = 0.5
    if query.sentiments:
        wanted = set(query.sentiments)
        # Treat angry/frustrated/negative as a family
        neg_family = {"angry", "frustrated", "negative"}
        pos_family = {"positive", "grateful"}
        present = set(k for k, v in sent_scores.items() if v >= 0.25)
        if sentiment in wanted or (wanted & present):
            sent_align = 1.0
            reasons.append(
                f"Thread sentiment is {sentiment} (asked for {', '.join(query.sentiments)})"
            )
        elif wanted & neg_family and (present & neg_family or sentiment in neg_family):
            sent_align = 0.85
            reasons.append(f"Negative / tense tone ({sentiment})")
        elif wanted & pos_family and (present & pos_family or sentiment in pos_family):
            sent_align = 0.85
            reasons.append(f"Positive tone ({sentiment})")
        elif sentiment == "neutral":
            sent_align = 0.2
            reasons.append("Asked for a tone, but the thread reads mostly neutral")
        else:
            sent_align = 0.05
            reasons.append(
                f"Tone mismatch: thread is {sentiment}, query asked for "
                f"{', '.join(query.sentiments)}"
            )
    elif sentiment != "neutral":
        reasons.append(f"Thread tone: {sentiment}")

    if kw_score > 0:
        reasons.append(
            f"Subject-matter overlap {int(kw_score * 100)}% "
            f"({', '.join(query.keywords[:6])})"
        )
    if topic_score > kw_score and query.topics:
        reasons.append(f"Topics in play: {', '.join(topics[:5]) or ', '.join(query.topics[:5])}")

    if query.sentiments and (query.keywords or query.topics):
        score = (
            0.32 * kw_score
            + 0.28 * sent_align
            + 0.25 * topic_score
            + 0.15 * spread
        )
    elif query.sentiments:
        score = 0.75 * sent_align + 0.15 * min(1.0, len(thread) / 4) + 0.10 * topic_score
    else:
        score = 0.45 * kw_score + 0.35 * topic_score + 0.20 * spread

    # Slight boost for real threads vs a lone message when the query is thematic
    if len(thread) >= 2 and (query.keywords or query.topics or query.sentiments):
        score = min(1.0, score + 0.04)
        if len(thread) >= 2:
            reasons.append(f"{len(thread)}-message conversation (not a one-off)")

    people: List[str] = []
    seen_p: set = set()
    for mail in thread:
        name = _participant_name(mail)
        key = name.lower()
        if key in seen_p:
            continue
        seen_p.add(key)
        people.append(name)

    snippets = []
    for mail in thread[-4:]:
        preview = re.sub(r"\s+", " ", (mail.body_preview or "")[:220]).strip()
        snippets.append(
            f"{mail.received or ''} | {mail.sender or ''}: "
            f"{mail.subject or ''} — {preview}"
        )
    summary = "\n".join(snippets)

    return ConversationHit(
        conversation_id=conversation_key(thread[0]),
        topic=topic_label[:180],
        sentiment=sentiment,
        sentiment_scores=sent_scores,
        participants=people[:8],
        message_count=len(thread),
        latest=latest_mail.received or "",
        score=round(float(score), 4),
        reasons=reasons[:8],
        subject_matter=topics,
        messages=[m.to_dict() for m in thread],
        summary=summary[:2500],
    )


def _seed_mails_from_hits(hits: Sequence[ConversationHit]) -> List[ScrapedMail]:
    seeds: List[ScrapedMail] = []
    for hit in hits:
        if not hit.messages:
            continue
        raw = hit.messages[-1]
        try:
            seeds.append(
                ScrapedMail(
                    entry_id=str(raw.get("entry_id") or ""),
                    store_id=str(raw.get("store_id") or ""),
                    account_smtp=str(raw.get("account_smtp") or ""),
                    subject=str(raw.get("subject") or ""),
                    sender=str(raw.get("sender") or ""),
                    to=str(raw.get("to") or ""),
                    cc=str(raw.get("cc") or ""),
                    received=str(raw.get("received") or ""),
                    received_dt=None,
                    unread=bool(raw.get("unread")),
                    importance=int(raw.get("importance") or 1),
                    flag_status=int(raw.get("flag_status") or 0),
                    message_class=str(raw.get("message_class") or ""),
                    conversation=str(raw.get("conversation") or ""),
                    body_preview=str(raw.get("body_preview") or ""),
                    recipient_role=str(raw.get("recipient_role") or "unknown"),
                    is_meeting_request=bool(raw.get("is_meeting_request")),
                    conversation_id=str(raw.get("conversation_id") or ""),
                    folder=str(raw.get("folder") or "Inbox"),
                )
            )
        except Exception:
            continue
    return [s for s in seeds if s.entry_id]


def _merge_mails(*groups: Iterable[ScrapedMail]) -> List[ScrapedMail]:
    by_id: Dict[str, ScrapedMail] = {}
    for group in groups:
        for mail in group:
            if mail.entry_id:
                by_id[mail.entry_id] = mail
    return list(by_id.values())


def _apply_ollama_rerank(
    hits: List[ConversationHit],
    query: SearchQuery,
    *,
    ollama_cfg: Optional[Dict[str, Any]],
    progress: ProgressFn = None,
) -> Tuple[List[ConversationHit], str]:
    if not ollama_cfg or not hits:
        return hits, ""
    model = str(ollama_cfg.get("model") or "").strip()
    if not model:
        return hits, "Ollama skipped — no model selected"

    from guri_ollama import (
        build_conversation_rank_prompt,
        chat as ollama_chat,
        parse_json_array,
    )

    top = hits[:12]
    payload = []
    for i, hit in enumerate(top, start=1):
        payload.append(
            {
                "id": i,
                "topic": hit.topic,
                "message_count": hit.message_count,
                "participants": hit.participants,
                "sentiment": hit.sentiment,
                "summary": hit.summary,
            }
        )
    prompt = build_conversation_rank_prompt(query.to_dict(), payload)
    if callable(progress):
        progress("Asking Ollama to rank conversations…")
    try:
        raw = ollama_chat(
            prompt,
            base_url=str(ollama_cfg.get("base_url") or ""),
            model=model,
            system=(
                "You rank email conversations for GURI. "
                "Judge the thread as a whole: tone, subject matter, and whether "
                "the discussion matches the query. JSON array only."
            ),
            timeout=float(ollama_cfg.get("timeout") or 90),
        )
    except Exception as exc:
        logger.warning("Ollama conversation rank failed: %s", exc)
        return hits, f"Ollama rank failed ({exc}); showing heuristic results"

    ranked = parse_json_array(raw)
    by_id: Dict[int, Dict[str, Any]] = {}
    for row in ranked:
        try:
            idx = int(row.get("id"))
        except (TypeError, ValueError):
            continue
        by_id[idx] = row
    if not by_id:
        return hits, "Ollama returned no usable ranking; showing heuristic results"

    blended: List[ConversationHit] = []
    for i, hit in enumerate(top, start=1):
        row = by_id.get(i) or {}
        try:
            llm_score = max(0.0, min(100.0, float(row.get("score") or 0))) / 100.0
        except (TypeError, ValueError):
            llm_score = hit.score
        new_score = round(0.55 * hit.score + 0.45 * llm_score, 4)
        why = str(row.get("why") or "").strip()
        llm_sent = str(row.get("sentiment") or "").strip()
        llm_topic = str(row.get("subject_matter") or "").strip()
        reasons = list(hit.reasons)
        if why:
            reasons.insert(0, why)
        subject_matter = list(hit.subject_matter)
        if llm_topic:
            extra = [t.strip() for t in re.split(r"[,;/]", llm_topic) if t.strip()]
            for t in extra:
                if t.lower() not in {s.lower() for s in subject_matter}:
                    subject_matter.append(t)
        blended.append(
            ConversationHit(
                conversation_id=hit.conversation_id,
                topic=hit.topic,
                sentiment=llm_sent or hit.sentiment,
                sentiment_scores=hit.sentiment_scores,
                participants=hit.participants,
                message_count=hit.message_count,
                latest=hit.latest,
                score=new_score,
                reasons=reasons[:8],
                subject_matter=subject_matter[:10],
                messages=hit.messages,
                summary=hit.summary,
                llm_why=why,
            )
        )
    blended.sort(key=lambda h: h.score, reverse=True)
    rest = hits[len(top):]
    return blended + rest, "Ollama ranked the top conversations"


def search_inbox_conversations(
    *,
    store_id: str,
    smtp: str = "",
    query_text: str,
    days: int = 30,
    max_messages: int = 120,
    include_sent: bool = True,
    use_ollama: bool = False,
    ollama_cfg: Optional[Dict[str, Any]] = None,
    progress: ProgressFn = None,
) -> Dict[str, Any]:
    """Scrape one mailbox and rank conversations against a thematic query."""
    query = parse_search_query(query_text)
    if query.is_empty():
        raise ValueError(
            "Describe what to look for — a topic, a tone (e.g. frustrated, "
            "positive), or both."
        )

    def _progress(msg: str) -> None:
        if callable(progress):
            try:
                progress(msg)
            except Exception:
                pass

    _progress("Scraping the selected inbox…")
    mails = scrape_inbox_for_search(
        store_id=store_id,
        smtp=smtp,
        days=days,
        max_messages=max_messages,
        include_sent=include_sent,
        unread_only=False,
        keywords=query.keywords or query.topics,
        progress=progress,
    )
    if not mails:
        return {
            "query": query.to_dict(),
            "hits": [],
            "mail_count": 0,
            "thread_count": 0,
            "used_ollama": False,
            "note": "No messages in that lookback window.",
        }

    _progress(f"Grouping {len(mails)} message(s) into conversations…")
    groups = group_conversations(mails)
    hits = [score_conversation(g, query) for g in groups]
    hits.sort(key=lambda h: h.score, reverse=True)

    # Expand the strongest threads so replies outside the recency cap join in
    expand_n = min(12, sum(1 for h in hits if h.score >= 0.12) or 8)
    seeds = _seed_mails_from_hits(hits[:expand_n])
    extra: List[ScrapedMail] = []
    if seeds:
        _progress("Pulling the rest of matching threads…")
        try:
            extra = expand_conversations(
                seeds,
                max_per_conversation=25,
                max_conversations=expand_n,
                progress=progress,
            )
        except Exception as exc:
            logger.debug("Thread expand skipped: %s", exc)
            extra = []
    if extra:
        mails = _merge_mails(mails, extra)
        groups = group_conversations(mails)
        hits = [score_conversation(g, query) for g in groups]
        hits.sort(key=lambda h: h.score, reverse=True)

    note = ""
    used_ollama = False
    if use_ollama:
        hits, note = _apply_ollama_rerank(
            hits, query, ollama_cfg=ollama_cfg, progress=progress
        )
        used_ollama = "ranked" in (note or "").lower()

    # Drop near-zero matches unless the query was sentiment-only and we need examples
    floor = 0.08 if (query.keywords or query.topics) else 0.18
    filtered = [h for h in hits if h.score >= floor]
    if not filtered:
        filtered = hits[:12]
        if not note:
            note = "No strong conversation matches — showing the closest threads."

    _progress(f"{len(filtered)} conversation(s) matched")
    return {
        "query": query.to_dict(),
        "hits": [h.to_dict() for h in filtered[:40]],
        "mail_count": len(mails),
        "thread_count": len(groups),
        "used_ollama": used_ollama,
        "note": note,
    }
