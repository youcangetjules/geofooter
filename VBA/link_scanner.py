#!/usr/bin/env python3
"""AES Deep Scan — extract and heuristically classify URLs from message body."""

from __future__ import annotations

import base64
import html as html_lib
import re
from dataclasses import dataclass, field
from typing import Iterable, List, Optional, Sequence, Set, Tuple
from urllib.parse import parse_qs, unquote, urlparse

# Quoted href="..." / href='...', or unquoted href=https://... (marketing HTML).
HREF_RE = re.compile(
    r"""(?is)<a\b[^>]*\bhref\s*=\s*(?:(?P<q>['"])(?P<urlq>.*?)(?P=q)|(?P<urlu>[^\s>]+))""",
)
BARE_URL_RE = re.compile(
    r"""(?i)\b((?:https?|ftp)://[^\s<>"'\)\]]+|www\.[^\s<>"'\)\]]+)""",
)
IP_HOST_RE = re.compile(
    r"""(?i)^(?:\d{1,3}\.){3}\d{1,3}$|^\[?[0-9a-f:]+\]?$""",
)

SUSPICIOUS_TLDS = frozenset({
    "zip", "mov", "click", "country", "tk", "ml", "ga", "cf", "gq", "top", "xyz",
    "work", "icu", "rest", "surf", "cfd", "sbs", "bond", "lol",
})

REDIRECTOR_HINTS = (
    "redirect", "redir", "r.php", "url=", "u=", "goto=", "out=", "away=",
    "click.", "link.", "track.", "email.", "e.mail",
)

LOOKALIKE_BRANDS = (
    "microsoft", "office365", "outlook", "paypal", "apple", "google", "amazon",
    "okta", "docusign", "dropbox", "linkedin", "facebook", "netflix", "adobe",
)

# Email-platform click-tracking hosts (registrable domains). Redirects through
# the sender's own domain or these ESPs are normal newsletter behaviour even
# when the decoded destination is a third-party site.
ESP_TRACKING_DOMAINS = frozenset({
    "sailthru.com", "list-manage.com", "mailchi.mp", "mcusercontent.com",
    "sendgrid.net", "mandrillapp.com", "mailgun.org", "exacttarget.com",
    "hubspotlinks.com", "klaviyo.com", "constantcontact.com", "rs6.net",
    "createsend.com", "cmail19.com", "cmail20.com", "sparkpostmail.com",
    "mailjet.com", "sendibt3.com", "customer.io", "substack.com", "beehiiv.com",
})

# Real brand properties — never treat these hosts as lookalikes of themselves.
OFFICIAL_BRAND_DOMAINS = frozenset({
    "google.com", "gmail.com", "youtube.com", "gstatic.com", "googleapis.com",
    "googleusercontent.com", "googlevideo.com",
    "microsoft.com", "live.com", "outlook.com", "office.com", "office365.com",
    "microsoftonline.com", "msn.com", "azure.com",
    "paypal.com", "apple.com", "icloud.com", "amazon.com", "amazonaws.com",
    "okta.com", "docusign.com", "dropbox.com", "linkedin.com", "facebook.com",
    "fbcdn.net", "netflix.com", "adobe.com",
})

_MULTI_PART_SUFFIXES = frozenset({
    "co.uk", "org.uk", "ac.uk", "gov.uk", "com.au", "net.au", "org.au",
    "co.nz", "co.jp", "com.br", "co.in", "co.za",
})

_REDIRECT_PARAMS = (
    "url", "u", "redirect", "redirect_url", "redirect_uri", "target", "dest",
    "destination", "goto", "out", "away", "link", "r", "to",
)

_B64_SEGMENT_RE = re.compile(r"^[A-Za-z0-9_\-+/]{16,}={0,2}$")


@dataclass
class LinkFinding:
    url: str
    source: str  # href | bare
    host: str = ""
    risk_level: str = "low"  # low | medium | high
    reasons: List[str] = field(default_factory=list)
    on_risktable: bool = False  # True when URL/domain hit threat intel / blocklists
    destination: str = ""  # decoded target of a redirect / click-tracking link

    def to_dict(self) -> dict:
        return {
            "url": self.url,
            "source": self.source,
            "host": self.host,
            "risk_level": self.risk_level,
            "reasons": list(self.reasons),
            "on_risktable": bool(self.on_risktable),
            "destination": self.destination,
        }


class LinkScanner:
    """Extract and classify links for Deep Scan reports."""

    def __init__(self, max_links: int = 200):
        self.max_links = max_links

    def scan(
        self,
        html_body: str = "",
        text_body: str = "",
        sender_domain: Optional[str] = None,
    ) -> List[LinkFinding]:
        findings: List[LinkFinding] = []
        seen: Set[str] = set()

        for url in self._extract_hrefs(html_body or ""):
            self._add(findings, seen, url, "href", sender_domain)
            if len(findings) >= self.max_links:
                return findings

        combined = f"{html_body or ''}\n{text_body or ''}"
        for url in self._extract_bare(combined):
            self._add(findings, seen, url, "bare", sender_domain)
            if len(findings) >= self.max_links:
                break

        return findings

    def risk_points(self, findings: Sequence[LinkFinding]) -> Tuple[int, List[str]]:
        """Return score points (higher=worse) and factor labels for SecurityAssessor fusion.

        Suspicious (medium) links: +2 each. High-risk links: +10 each.
        """
        points = 0
        factors: List[str] = []
        highs = sum(1 for f in findings if f.risk_level == "high")
        meds = sum(1 for f in findings if f.risk_level == "medium")
        if meds:
            add = 2 * meds
            points += add
            factors.append(f"Suspicious links ({meds}): +{add}")
        if highs:
            add = 10 * highs
            points += add
            factors.append(f"High-risk links ({highs}): +{add}")
        return points, factors

    def _add(
        self,
        findings: List[LinkFinding],
        seen: Set[str],
        url: str,
        source: str,
        sender_domain: Optional[str],
    ) -> None:
        cleaned = (url or "").strip()
        if not cleaned or cleaned.lower().startswith(("mailto:", "tel:", "cid:", "#")):
            return
        if cleaned.lower().startswith("www."):
            cleaned = "http://" + cleaned
        key = cleaned.lower()
        if key in seen:
            return
        seen.add(key)
        findings.append(self._classify(cleaned, source, sender_domain))

    def _classify(
        self,
        url: str,
        source: str,
        sender_domain: Optional[str],
    ) -> LinkFinding:
        finding = LinkFinding(url=url, source=source)
        try:
            parsed = urlparse(url)
        except Exception:
            finding.risk_level = "medium"
            finding.reasons.append("Malformed URL")
            return finding

        host = (parsed.hostname or "").lower().rstrip(".")
        finding.host = host
        if not host:
            finding.risk_level = "medium"
            finding.reasons.append("Missing host")
            return finding

        reasons: List[str] = []
        risk = "low"

        if IP_HOST_RE.match(host):
            reasons.append("IP-literal host")
            risk = "high"

        labels = host.split(".")
        tld = labels[-1] if labels else ""
        if tld in SUSPICIOUS_TLDS:
            reasons.append(f"Suspicious TLD .{tld}")
            risk = _max_risk(risk, "medium")

        path_q = f"{parsed.path or ''}?{parsed.query or ''}".lower()
        destination = decode_redirect_target(parsed)
        finding.destination = destination
        host_rd = registrable_domain(host)
        if destination or any(h in path_q or h in host for h in REDIRECTOR_HINTS):
            sd = registrable_domain((sender_domain or "").lstrip("@"))
            dest_host = (urlparse(destination).hostname or "").lower() if destination else ""
            dest_rd = registrable_domain(dest_host)
            # Sender tracking subdomain, known ESP, or official brand CDN/redirect.
            host_ok = (
                bool(sd and host_rd == sd)
                or host_rd in ESP_TRACKING_DOMAINS
                or host_rd in OFFICIAL_BRAND_DOMAINS
            )
            if host_ok:
                # Normal newsletter / ESP / Google redirect — not a risk signal.
                if dest_host:
                    reasons.append(f"Tracking / CDN redirect to {dest_host}")
                else:
                    reasons.append("Tracking / CDN redirect hop")
            elif dest_rd and sd and dest_rd == sd:
                reasons.append(f"Redirects to sender's site via {host}")
                risk = _max_risk(risk, "medium")
            elif dest_host:
                reasons.append(f"Redirects to {dest_host} via unknown hop {host}")
                risk = _max_risk(risk, "medium")
            else:
                reasons.append(f"Possible redirector / tracking hop on {host}")
                risk = _max_risk(risk, "medium")

        if "@" in (parsed.netloc or ""):
            reasons.append("Credentials embedded in URL netloc")
            risk = "high"

        if sender_domain:
            sd = sender_domain.lower().lstrip("@")
            if (
                sd
                and host
                and not (host == sd or host.endswith("." + sd))
                and host_rd not in OFFICIAL_BRAND_DOMAINS
            ):
                # Brand lookalike in host while From domain differs — skip real brand hosts.
                for brand in LOOKALIKE_BRANDS:
                    if brand in host.replace("-", "") and brand not in sd.replace("-", ""):
                        reasons.append(f"Possible {brand} lookalike vs From domain")
                        risk = _max_risk(risk, "high")
                        break

        # Punycode / homoglyph hint
        if "xn--" in host:
            reasons.append("Punycode (IDN) hostname")
            risk = _max_risk(risk, "medium")

        finding.reasons = reasons
        finding.risk_level = risk if reasons else "low"
        if not reasons:
            finding.reasons = ["No heuristic flags"]
        return finding

    @staticmethod
    def _extract_hrefs(html_body: str) -> Iterable[str]:
        for match in HREF_RE.finditer(html_body or ""):
            raw = html_lib.unescape(
                match.group("urlq") or match.group("urlu") or ""
            ).strip()
            if raw:
                yield raw

    @staticmethod
    def _extract_bare(text: str) -> Iterable[str]:
        for match in BARE_URL_RE.finditer(text or ""):
            raw = match.group(1).rstrip(".,;:)")
            if raw:
                yield raw


def _max_risk(a: str, b: str) -> str:
    order = {"low": 0, "medium": 1, "high": 2}
    return a if order.get(a, 0) >= order.get(b, 0) else b


def registrable_domain(host: str) -> str:
    """Approximate registrable domain (foo.bar.co.uk -> bar.co.uk)."""
    labels = [p for p in (host or "").lower().strip(".").split(".") if p]
    if len(labels) <= 2:
        return ".".join(labels)
    if ".".join(labels[-2:]) in _MULTI_PART_SUFFIXES:
        return ".".join(labels[-3:])
    return ".".join(labels[-2:])


def _b64_url(segment: str) -> str:
    """Decode a (URL-safe) base64 path/query segment that wraps an http(s) URL."""
    if not segment or not _B64_SEGMENT_RE.match(segment):
        return ""
    s = segment.replace("-", "+").replace("_", "/")
    s += "=" * (-len(s) % 4)
    try:
        raw = base64.b64decode(s).decode("utf-8").strip()
    except Exception:
        return ""
    if raw.lower().startswith(("http://", "https://")) and urlparse(raw).hostname:
        return raw
    return ""


def decode_redirect_target(parsed) -> str:
    """Destination URL carried by a redirect link (query param or base64 path segment)."""
    qs = parse_qs(parsed.query or "")
    for key in _REDIRECT_PARAMS:
        for value in qs.get(key, []):
            cand = unquote(value).strip()
            if cand.lower().startswith(("http://", "https://")) and urlparse(cand).hostname:
                return cand
            decoded = _b64_url(cand)
            if decoded:
                return decoded
    for seg in (parsed.path or "").split("/"):
        decoded = _b64_url(seg)
        if decoded:
            return decoded
    return ""
