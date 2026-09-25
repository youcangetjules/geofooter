#!/usr/bin/env python3
"""Persistent manual threat ratings for message links.

A rating set from the AES Link Safety page is stored under
%LOCALAPPDATA%\\GeoFooter\\aes_link_ratings.json and applied on later scans
of the same URL. The scanner's own level is kept so it can be restored.
"""

from __future__ import annotations

import hashlib
import html
import json
import os
import re
from datetime import datetime, timezone
from pathlib import Path
from typing import Any, Dict, List, Optional
from urllib.parse import urlsplit, urlunsplit

LEVELS = ("low", "medium", "high")
LEVEL_LABEL = {"low": "Low", "medium": "Medium", "high": "High"}
MANUAL_PREFIX = "Manual rating"
_REPORT_STEM = re.compile(r"^links_report_[0-9a-f]{12}$")
_LINK_ID = re.compile(r"^[0-9a-f]{16}$")


def ratings_path() -> Path:
    local = os.environ.get("LOCALAPPDATA", "")
    base = Path(local) / "GeoFooter" if local else Path.home() / "GeoFooter"
    base.mkdir(parents=True, exist_ok=True)
    return base / "aes_link_ratings.json"


def normalize_url(url: str) -> str:
    """Stable key for one link. Fragment is dropped; the rest is lowercased."""
    text = (url or "").strip()
    if text.lower().startswith("www."):
        text = "http://" + text
    try:
        parts = urlsplit(text)
    except Exception:
        return text.lower()
    if not parts.scheme or not parts.netloc:
        return text.lower()
    return urlunsplit(
        (parts.scheme.lower(), parts.netloc.lower(), parts.path or "", parts.query or "", "")
    ).lower()


def link_id(url: str) -> str:
    return hashlib.sha256(normalize_url(url).encode("utf-8")).hexdigest()[:16]


def load_ratings() -> Dict[str, Dict[str, str]]:
    path = ratings_path()
    if not path.is_file():
        return {}
    try:
        raw = json.loads(path.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError):
        return {}
    items = raw.get("ratings") if isinstance(raw, dict) else None
    if not isinstance(items, dict):
        return {}
    out: Dict[str, Dict[str, str]] = {}
    for key, value in items.items():
        if not isinstance(value, dict):
            continue
        level = str(value.get("level") or "").lower()
        if level not in LEVELS:
            continue
        out[str(key)] = {
            "level": level,
            "updated": str(value.get("updated") or ""),
        }
    return out


def save_rating(url: str, level: str) -> None:
    """Store level, or remove the manual rating when level is 'clear'."""
    key = normalize_url(url)
    if not key:
        raise ValueError("missing url")
    chosen = (level or "").strip().lower()
    ratings = load_ratings()
    if chosen == "clear":
        ratings.pop(key, None)
    elif chosen in LEVELS:
        ratings[key] = {
            "level": chosen,
            "updated": datetime.now(timezone.utc).strftime("%Y-%m-%d %H:%M:%S UTC"),
        }
    else:
        raise ValueError("bad level")
    path = ratings_path()
    payload = {"ratings": ratings}
    tmp = path.with_suffix(".json.tmp")
    tmp.write_text(json.dumps(payload, indent=2), encoding="utf-8")
    os.replace(tmp, path)


def _get(item: Any, key: str) -> Any:
    if isinstance(item, dict):
        return item.get(key)
    return getattr(item, key, None)


def _set(item: Any, key: str, value: Any) -> None:
    if isinstance(item, dict):
        item[key] = value
    else:
        setattr(item, key, value)


def apply_saved_ratings(findings: Optional[List[Any]]) -> None:
    """Overlay saved ratings onto scanner findings. Manual level wins."""
    ratings = load_ratings()
    for item in findings or []:
        url = str(_get(item, "url") or "")
        scanner = str(_get(item, "scanner_level") or "") or str(_get(item, "risk_level") or "low")
        if scanner not in LEVELS:
            scanner = "low"
        if not str(_get(item, "scanner_level") or ""):
            _set(item, "scanner_level", scanner)
        reasons = [
            str(reason)
            for reason in (_get(item, "reasons") or [])
            if not str(reason).startswith(MANUAL_PREFIX)
        ]
        saved = ratings.get(normalize_url(url))
        if not saved:
            _set(item, "reasons", reasons)
            _set(item, "manual", False)
            continue
        reasons.insert(0, f"Manual rating (scanner said {scanner.upper()})")
        _set(item, "risk_level", saved["level"])
        _set(item, "reasons", reasons)
        _set(item, "manual", True)


def snapshots_from_findings(findings: Optional[List[Any]]) -> List[Dict[str, Any]]:
    """Scanner-side copy of each link, without the manual overlay."""
    snaps: List[Dict[str, Any]] = []
    for item in findings or []:
        url = str(_get(item, "url") or "")
        if not url:
            continue
        scanner = str(_get(item, "scanner_level") or "") or str(_get(item, "risk_level") or "low")
        if scanner not in LEVELS:
            scanner = "low"
        reasons = [
            str(reason)
            for reason in (_get(item, "reasons") or [])
            if not str(reason).startswith(MANUAL_PREFIX)
        ]
        snaps.append(
            {
                "id": link_id(url),
                "url": url,
                "source": str(_get(item, "source") or ""),
                "host": str(_get(item, "host") or ""),
                "scanner_level": scanner,
                "scanner_reasons": reasons,
            }
        )
    return snaps


def _rating_href(stem: str, item_id: str, level: str) -> str:
    from urllib.parse import urlencode

    query = urlencode({"report": stem, "id": item_id, "level": level})
    return f"aes://set-link-risk?{query}"


def _rate_cell(stem: str, item_id: str, level: str, manual: bool) -> str:
    parts: List[str] = []
    for name in LEVELS:
        label = LEVEL_LABEL[name]
        href = html.escape(_rating_href(stem, item_id, name), quote=True)
        label_html = f"<strong>{label}</strong>" if name == level else label
        parts.append(f"<a class='rate' href='{href}'>{label_html}</a>")
    if manual:
        href = html.escape(_rating_href(stem, item_id, "clear"), quote=True)
        parts.append(f"<a class='rate' href='{href}'>Use scanner</a>")
    return " ".join(parts)


def render_links_report(
    snapshots: List[Dict[str, Any]],
    *,
    stem: str,
    subject: str,
    sender_email: str,
    sender_domain: str,
    guri: str,
    scanned_at: str,
    body_available: bool,
) -> str:
    """HTML for the AES Link Safety page, including per-link rating controls."""
    ratings = load_ratings()
    displayed: List[Dict[str, Any]] = []
    for snap in snapshots[:200]:
        scanner = str(snap.get("scanner_level") or "low")
        if scanner not in LEVELS:
            scanner = "low"
        reasons = [str(r) for r in (snap.get("scanner_reasons") or [])]
        level = scanner
        manual = False
        saved = ratings.get(normalize_url(str(snap.get("url") or "")))
        if saved:
            level = saved["level"]
            manual = True
            reasons.insert(0, f"Manual rating (scanner said {scanner.upper()})")
        displayed.append(
            {
                "id": str(snap.get("id") or link_id(str(snap.get("url") or ""))),
                "url": str(snap.get("url") or ""),
                "source": str(snap.get("source") or ""),
                "host": str(snap.get("host") or ""),
                "risk_level": level,
                "reasons": reasons,
                "manual": manual,
            }
        )

    highs = [row for row in displayed if row["risk_level"] == "high"]
    meds = [row for row in displayed if row["risk_level"] == "medium"]
    lows = [row for row in displayed if row["risk_level"] not in {"high", "medium"}]
    ordered = highs + meds + lows

    rows = ""
    for idx, row in enumerate(ordered, 1):
        risk = str(row["risk_level"])
        risk_color = {"high": "#FF4444", "medium": "#CC8800", "low": "#008000"}.get(risk, "#333")
        reasons = "; ".join(row["reasons"]) or "—"
        rows += (
            "<tr>"
            f"<td style='padding:6px;border-bottom:1px solid #eee;'>{idx}</td>"
            f"<td style='padding:6px;border-bottom:1px solid #eee;color:{risk_color};"
            f"font-weight:bold;'>{html.escape(risk.upper())}</td>"
            f"<td style='padding:6px;border-bottom:1px solid #eee;'>"
            f"{html.escape(row['source'])}</td>"
            f"<td style='padding:6px;border-bottom:1px solid #eee;'>"
            f"{html.escape(row['host'])}</td>"
            f"<td style='padding:6px;border-bottom:1px solid #eee;word-break:break-all;'>"
            f"{html.escape(row['url'])}</td>"
            f"<td style='padding:6px;border-bottom:1px solid #eee;font-size:11px;'>"
            f"{html.escape(reasons)}</td>"
            f"<td style='padding:6px;border-bottom:1px solid #eee;white-space:nowrap;'>"
            f"{_rate_cell(stem, row['id'], risk, bool(row['manual']))}</td>"
            "</tr>"
        )
    if not rows:
        message = (
            "No links found in the message body."
            if body_available
            else "Body was not exported — links were not checked for this scan."
        )
        rows = f"<tr><td colspan='7' style='padding:8px;'>{message}</td></tr>"

    summary_colour = (
        "#FF4444" if highs else ("#CC8800" if meds else ("#008000" if displayed else "#666"))
    )
    return f"""<!DOCTYPE html>
<html lang="en"><head><meta charset="utf-8">
<meta http-equiv="Cache-Control" content="no-cache, no-store, must-revalidate">
<title>AES Link Safety</title>
<style>body{{font-family:Arial,sans-serif;margin:1.5em;color:#333;background:#fff}}
h1{{font-size:18px;margin-bottom:0.2em}}.meta{{color:#666;font-size:12px;margin-bottom:1em}}
.summary{{background:#f5f5f5;border:1px solid #ddd;border-radius:4px;padding:10px;margin-bottom:1em;font-size:13px}}
table{{width:100%;border-collapse:collapse;font-size:12px}}th{{text-align:left;background:#4c7a2c;color:#fff;padding:6px}}
.footer{{margin-top:1.5em;font-size:11px;color:#666;border-top:1px solid #eee;padding-top:0.8em}}
.note{{font-size:11px;color:#666;margin-top:0.8em;line-height:1.45}}
a.rate{{color:#1a5fb4;margin-right:8px;text-decoration:none}}
a.rate:hover{{text-decoration:underline}}
strong.rate{{margin-right:8px}}
</style></head><body>
<h1>AES Link Safety</h1>
<div class="meta">Subject: {html.escape(subject)}<br>
Sender: {html.escape(sender_email or 'Unknown')} ({html.escape(sender_domain or 'Unknown')})<br>
Scanned: {html.escape(scanned_at)} | GURI: {html.escape(guri or 'N/A')}</div>
<div class="summary" style="color:{summary_colour};">
<strong>{len(displayed)}</strong> link(s) checked —
<span style="color:#FF4444;font-weight:bold;">{len(highs)} high</span>,
<span style="color:#CC8800;font-weight:bold;">{len(meds)} suspicious</span>,
<span style="color:#008000;font-weight:bold;">{len(lows)} low</span>
</div>
<table>
<tr><th>#</th><th>Risk</th><th>Source</th><th>Host</th><th>URL</th><th>Notes</th><th>Your rating</th></tr>
{rows}
</table>
<p class="note">Heuristic checks only (IP hosts, shorteners, lookalikes, suspicious TLDs).
Live Safe Browsing lookups are optional and separate.
Choose Low, Medium, or High to keep that rating for this URL on later scans.
Use scanner clears the manual rating. Refresh this page after a choice is saved.</p>
<div class="footer">(C) Aliniant Labs | Aliniant Email Scanner (AES)</div>
</body></html>"""


def write_links_bundle(
    links_dir: Path,
    stem: str,
    findings: Optional[List[Any]],
    *,
    subject: str,
    sender_email: str,
    sender_domain: str,
    guri: str,
    scanned_at: str,
    body_available: bool,
) -> None:
    if not _REPORT_STEM.match(stem):
        raise ValueError("bad report name")
    links_dir.mkdir(parents=True, exist_ok=True)
    snapshots = snapshots_from_findings(findings)
    sidecar = {
        "subject": subject,
        "sender_email": sender_email,
        "sender_domain": sender_domain,
        "guri": guri,
        "scanned_at": scanned_at,
        "body_available": bool(body_available),
        "findings": snapshots,
    }
    html_text = render_links_report(
        snapshots,
        stem=stem,
        subject=subject,
        sender_email=sender_email,
        sender_domain=sender_domain,
        guri=guri,
        scanned_at=scanned_at,
        body_available=body_available,
    )
    json_path = links_dir / f"{stem}.json"
    html_path = links_dir / f"{stem}.html"
    json_tmp = json_path.with_suffix(".json.tmp")
    json_tmp.write_text(json.dumps(sidecar, indent=2), encoding="utf-8")
    os.replace(json_tmp, json_path)
    html_path.write_text(html_text, encoding="utf-8")


def valid_report_stem(stem: str) -> bool:
    return bool(_REPORT_STEM.match(stem or ""))


def valid_link_id(item_id: str) -> bool:
    return bool(_LINK_ID.match(item_id or ""))


def rewrite_report(html_path: Path, sidecar_path: Path) -> None:
    raw = json.loads(sidecar_path.read_text(encoding="utf-8"))
    if not isinstance(raw, dict):
        raise ValueError("bad sidecar")
    stem = html_path.stem
    html_text = render_links_report(
        list(raw.get("findings") or []),
        stem=stem,
        subject=str(raw.get("subject") or ""),
        sender_email=str(raw.get("sender_email") or ""),
        sender_domain=str(raw.get("sender_domain") or ""),
        guri=str(raw.get("guri") or ""),
        scanned_at=str(raw.get("scanned_at") or ""),
        body_available=bool(raw.get("body_available")),
    )
    html_path.write_text(html_text, encoding="utf-8")


def finding_url(sidecar: Dict[str, Any], item_id: str) -> str:
    for snap in sidecar.get("findings") or []:
        if isinstance(snap, dict) and str(snap.get("id") or "") == item_id:
            return str(snap.get("url") or "")
    return ""
