#!/usr/bin/env python3
"""AES Deep Scan — body / phishing language heuristics."""

from __future__ import annotations

import re
from dataclasses import dataclass, field
from typing import List, Optional, Sequence, Tuple

URGENCY_PATTERNS = [
    r"\burgent\b",
    r"\bimmediately\b",
    r"\bwithin\s+\d+\s+hours?\b",
    r"\baccount\s+will\s+be\s+(?:suspended|locked|closed)\b",
    r"\bact\s+now\b",
    r"\bfinal\s+notice\b",
    r"\bexpires?\s+today\b",
]

CREDENTIAL_PATTERNS = [
    r"\bverify\s+(?:your\s+)?(?:password|account|identity)\b",
    r"\bconfirm\s+(?:your\s+)?(?:password|credentials|login)\b",
    r"\bupdate\s+(?:your\s+)?(?:password|payment|billing)\b",
    r"\blog\s*in\s+(?:to\s+)?(?:continue|verify|secure)\b",
    r"\bsecurity\s+alert\b",
    r"\bunusual\s+(?:sign[- ]?in|activity)\b",
    r"\bclick\s+(?:here|below)\s+to\s+(?:verify|confirm|secure)\b",
]

BRAND_PATTERNS = [
    (r"\bmicrosoft\b|\boffice\s*365\b|\boutlook\b", "microsoft"),
    (r"\bpaypal\b", "paypal"),
    (r"\bapple\s*id\b|\bicloud\b", "apple"),
    (r"\bgoogle\b|\bgmail\b", "google"),
    (r"\bamazon\b", "amazon"),
    (r"\bdocusign\b", "docusign"),
    (r"\bokta\b", "okta"),
]

DISPLAY_FROM_RE = re.compile(
    r'(?im)^From:\s*(?:"?([^"<]*)"?\s*)?<?([^>\s]+@[^>\s]+)>?',
)


@dataclass
class BodyFinding:
    category: str
    detail: str
    risk_level: str = "medium"  # low | medium | high

    def to_dict(self) -> dict:
        return {
            "category": self.category,
            "detail": self.detail,
            "risk_level": self.risk_level,
        }


@dataclass
class BodyScanResult:
    findings: List[BodyFinding] = field(default_factory=list)
    display_name: str = ""
    from_email: str = ""
    from_domain: str = ""

    def to_dict(self) -> dict:
        return {
            "findings": [f.to_dict() for f in self.findings],
            "display_name": self.display_name,
            "from_email": self.from_email,
            "from_domain": self.from_domain,
        }


class BodyScanner:
    """Heuristic scan of message body + From display-name consistency."""

    def scan(
        self,
        html_body: str = "",
        text_body: str = "",
        headers: str = "",
        sender_email: Optional[str] = None,
        sender_domain: Optional[str] = None,
    ) -> BodyScanResult:
        result = BodyScanResult()
        text = _strip_html(html_body or "")
        if text_body:
            text = f"{text}\n{text_body}"
        text_l = text.lower()

        self._parse_from(headers or "", result, sender_email, sender_domain)

        for pat in URGENCY_PATTERNS:
            if re.search(pat, text_l, re.IGNORECASE):
                result.findings.append(
                    BodyFinding("urgency", f"Matched urgency cue: {pat}", "medium")
                )
                break

        cred_hits = 0
        for pat in CREDENTIAL_PATTERNS:
            if re.search(pat, text_l, re.IGNORECASE):
                cred_hits += 1
        if cred_hits:
            level = "high" if cred_hits >= 2 else "medium"
            result.findings.append(
                BodyFinding(
                    "credential_harvest",
                    f"{cred_hits} credential / verify cue(s) in body",
                    level,
                )
            )

        sd = (sender_domain or result.from_domain or "").lower()
        for pat, brand in BRAND_PATTERNS:
            if re.search(pat, text_l, re.IGNORECASE):
                if sd and brand not in sd.replace("-", ""):
                    result.findings.append(
                        BodyFinding(
                            "brand_impersonation",
                            f"Body references {brand} but From domain is {sd or 'unknown'}",
                            "high",
                        )
                    )
                break

        if result.display_name and result.from_domain:
            name_l = result.display_name.lower()
            for _, brand in BRAND_PATTERNS:
                if brand in name_l.replace(" ", "") and brand not in result.from_domain.replace("-", ""):
                    result.findings.append(
                        BodyFinding(
                            "display_name_mismatch",
                            f'Display name "{result.display_name}" suggests {brand}; '
                            f"From domain is {result.from_domain}",
                            "high",
                        )
                    )
                    break

        return result

    def risk_points(self, result: BodyScanResult) -> Tuple[int, List[str]]:
        points = 0
        factors: List[str] = []
        highs = sum(1 for f in result.findings if f.risk_level == "high")
        meds = sum(1 for f in result.findings if f.risk_level == "medium")
        if highs:
            add = min(20, 7 * highs)
            points += add
            factors.append(f"Body phishing cues high ({highs}): +{add}")
        if meds and not highs:
            add = min(10, 3 * meds)
            points += add
            factors.append(f"Body phishing cues medium ({meds}): +{add}")
        return points, factors

    def _parse_from(
        self,
        headers: str,
        result: BodyScanResult,
        sender_email: Optional[str],
        sender_domain: Optional[str],
    ) -> None:
        match = DISPLAY_FROM_RE.search(headers or "")
        if match:
            result.display_name = (match.group(1) or "").strip().strip('"')
            result.from_email = (match.group(2) or "").strip()
        if sender_email and not result.from_email:
            result.from_email = sender_email
        if result.from_email and "@" in result.from_email:
            result.from_domain = result.from_email.rsplit("@", 1)[-1].lower()
        elif sender_domain:
            result.from_domain = sender_domain.lower()


def _strip_html(html_body: str) -> str:
    if not html_body:
        return ""
    text = re.sub(r"(?is)<script[^>]*>.*?</script>", " ", html_body)
    text = re.sub(r"(?is)<style[^>]*>.*?</style>", " ", text)
    text = re.sub(r"(?is)<[^>]+>", " ", text)
    text = re.sub(r"\s+", " ", text)
    return text.strip()
