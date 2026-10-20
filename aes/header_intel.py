"""Header intelligence: identity, sending platform, signing, transport, threads.

Each check takes the raw header block and returns an :class:`IntelSection`:
``rows`` are plain-text lines for the Full Scan report and ``findings`` are
``{label, points, note}`` dicts for ``SecurityAssessor.apply_extra_factors``.
Header-only checks are cheap enough for compact scans.
"""

from __future__ import annotations

import re
from dataclasses import dataclass, field
from email.utils import getaddresses, parseaddr
from typing import Any, Dict, Iterable, List, Optional, Sequence, Tuple

FREE_MAIL_DOMAINS = frozenset(
    {
        "gmail.com", "googlemail.com", "yahoo.com", "yahoo.co.uk", "ymail.com",
        "outlook.com", "hotmail.com", "hotmail.co.uk", "live.com", "live.co.uk",
        "msn.com", "aol.com", "icloud.com", "me.com", "mac.com", "proton.me",
        "protonmail.com", "pm.me", "gmx.com", "gmx.net", "gmx.de", "mail.com",
        "yandex.com", "yandex.ru", "zoho.com", "tutanota.com", "mail.ru",
        "qq.com", "163.com", "web.de",
    }
)

# Brand term -> registrable domains allowed to use it in a display name.
BRAND_DOMAINS: Dict[str, Tuple[str, ...]] = {
    "paypal": ("paypal.com", "paypal.co.uk"),
    "microsoft": ("microsoft.com", "office.com", "office365.com", "microsoftonline.com",
                  "sharepoint.com", "outlook.com", "live.com", "onmicrosoft.com"),
    "office 365": ("microsoft.com", "office.com", "office365.com", "microsoftonline.com"),
    "sharepoint": ("microsoft.com", "sharepoint.com"),
    "onedrive": ("microsoft.com", "onedrive.com", "sharepoint.com"),
    "docusign": ("docusign.com", "docusign.net"),
    "dropbox": ("dropbox.com", "dropboxmail.com"),
    "amazon": ("amazon.com", "amazon.co.uk", "amazon.de", "amazonses.com", "aws.com"),
    "apple": ("apple.com", "icloud.com"),
    "netflix": ("netflix.com",),
    "dhl": ("dhl.com", "dhl.co.uk", "dhl.de"),
    "fedex": ("fedex.com",),
    "hmrc": ("hmrc.gov.uk", "gov.uk"),
    "linkedin": ("linkedin.com",),
    "adobe": ("adobe.com", "adobesign.com", "echosign.com"),
    "coinbase": ("coinbase.com",),
    "binance": ("binance.com",),
    "barclays": ("barclays.com", "barclays.co.uk"),
    "hsbc": ("hsbc.com", "hsbc.co.uk"),
    "lloyds": ("lloydsbank.com", "lloydsbank.co.uk", "lloydsbankinggroup.com"),
    "natwest": ("natwest.com",),
    "santander": ("santander.co.uk", "santander.com"),
    "wells fargo": ("wellsfargo.com",),
}

_MULTI_PART_SUFFIXES = frozenset(
    {"co.uk", "org.uk", "gov.uk", "ac.uk", "com.au", "net.au", "co.nz", "co.jp",
     "com.br", "co.za", "com.cn", "co.in"}
)
_EMAIL_RE = re.compile(r"[A-Za-z0-9._%+'-]+@([A-Za-z0-9.-]+\.[A-Za-z]{2,})")
_DOMAIN_RE = re.compile(r"\b((?:[a-z0-9-]+\.)+(?:com|net|org|co\.uk|io|gov|uk|de|fr|info|biz))\b", re.I)


@dataclass
class IntelSection:
    title: str
    rows: List[str] = field(default_factory=list)
    findings: List[Dict[str, Any]] = field(default_factory=list)
    data: Dict[str, Any] = field(default_factory=dict)

    def add(self, label: str, points: int, note: str = "") -> None:
        self.findings.append({"label": label, "points": int(points), "note": note})


def unfold(headers: str) -> str:
    return re.sub(r"\r?\n[ \t]+", " ", headers or "")


def header_values(headers: str, name: str) -> List[str]:
    """Every unfolded value of header *name*, in file order."""
    pat = re.compile(rf"^{re.escape(name)}:[ \t]*(.*)$", re.I | re.M)
    return [m.group(1).strip() for m in pat.finditer(unfold(headers))]


def first_header(headers: str, name: str) -> str:
    vals = header_values(headers, name)
    return vals[0] if vals else ""


def registrable_domain(host: str) -> str:
    host = (host or "").strip().lower().rstrip(".").lstrip("@")
    if "@" in host:
        host = host.rsplit("@", 1)[1]
    if not host or re.match(r"^[\d.]+$", host) or ":" in host:
        return host
    labels = host.split(".")
    if len(labels) >= 3 and ".".join(labels[-2:]) in _MULTI_PART_SUFFIXES:
        return ".".join(labels[-3:])
    return ".".join(labels[-2:])


def _domain_of(addr: str) -> str:
    addr = (addr or "").strip().strip("<>")
    return addr.rsplit("@", 1)[1].lower() if "@" in addr else ""


def _addr(value: str) -> Tuple[str, str]:
    name, addr = parseaddr(value or "")
    return name.strip().strip('"').strip(), addr.strip().lower()


# --------------------------------------------------------------------------- #
# Identity (BEC / display-name spoofing)
# --------------------------------------------------------------------------- #

def _brand_hit(display: str, from_reg: str) -> Optional[str]:
    low = f" {display.lower()} "
    for brand, allowed in BRAND_DOMAINS.items():
        if not re.search(rf"(?<![a-z]){re.escape(brand)}(?![a-z])", low):
            continue
        if from_reg in allowed or brand.replace(" ", "") in from_reg:
            return None
        return brand
    return None


def identity_signals(
    headers: str,
    *,
    dmarc: Optional[str] = None,
    own_domains: Iterable[str] = (),
) -> IntelSection:
    sec = IntelSection("Identity")
    from_name, from_addr = _addr(first_header(headers, "From"))
    from_dom = _domain_of(from_addr)
    from_reg = registrable_domain(from_dom)
    sec.data.update(from_name=from_name, from_addr=from_addr)
    if not from_addr:
        return sec
    sec.rows.append(f"From: {from_name + ' ' if from_name else ''}<{from_addr}>")

    reply_list = [a.lower() for _n, a in getaddresses(header_values(headers, "Reply-To")) if a]
    for reply in reply_list:
        r_reg = registrable_domain(_domain_of(reply))
        if not r_reg or r_reg == from_reg:
            sec.rows.append(f"Reply-To: {reply} (same domain)")
            continue
        sec.rows.append(f"Reply-To: {reply} (different domain from From)")
        if r_reg in FREE_MAIL_DOMAINS and from_reg not in FREE_MAIL_DOMAINS:
            sec.add(
                f"Reply-To goes to free webmail ({r_reg}) while From is {from_reg}",
                15,
                "Replies would leave the sender's organisation; a common BEC setup",
            )
        else:
            sec.add(
                f"Reply-To domain {r_reg} differs from From domain {from_reg}",
                10,
                "Replies go somewhere other than the address shown",
            )
        break

    display_flagged = False
    if from_name:
        for m in _EMAIL_RE.finditer(from_name):
            if registrable_domain(m.group(1)) != from_reg:
                sec.add(
                    f"Display name shows {m.group(0)} but the mail is from {from_addr}",
                    15,
                    "Display-name spoofing: the visible address is not the sender",
                )
                display_flagged = True
                break
        if not display_flagged:
            for m in _DOMAIN_RE.finditer(from_name):
                if registrable_domain(m.group(1)) != from_reg:
                    sec.add(
                        f"Display name names {m.group(1).lower()} but the mail is from {from_reg}",
                        15,
                        "Display-name spoofing: the visible domain is not the sender",
                    )
                    display_flagged = True
                    break
        if not display_flagged:
            brand = _brand_hit(from_name, from_reg)
            if brand:
                sec.add(
                    f"Display name uses brand '{brand}' but the mail is from {from_reg}",
                    10,
                    "Brand impersonation in the display name",
                )
                display_flagged = True
        if not display_flagged:
            for own in own_domains:
                own_reg = registrable_domain(own)
                label = own_reg.split(".")[0] if own_reg else ""
                if len(label) < 4 or own_reg == from_reg:
                    continue
                if re.search(rf"(?<![a-z]){re.escape(label)}(?![a-z])", from_name.lower()):
                    sec.add(
                        f"Display name uses '{label}' but the mail is from {from_reg}",
                        10,
                        "Your organisation's name on an outside address",
                    )
                    break

    rp = first_header(headers, "Return-Path").strip().strip("<>").lower()
    if rp and "@" in rp:
        rp_reg = registrable_domain(_domain_of(rp))
        if rp_reg and rp_reg != from_reg:
            sec.rows.append(f"Return-Path: {rp} (bounce domain differs from From)")
            if (dmarc or "").lower() != "pass":
                sec.add(
                    f"Return-Path domain {rp_reg} differs from From domain {from_reg} without DMARC pass",
                    5,
                    "Bounce address belongs to someone else and DMARC did not vouch for From",
                )
        else:
            sec.rows.append(f"Return-Path: {rp}")

    sender_hdr = _addr(first_header(headers, "Sender"))[1]
    if sender_hdr and registrable_domain(_domain_of(sender_hdr)) != from_reg:
        sec.rows.append(f"Sender: {sender_hdr} (sending on behalf of From)")
    return sec


def load_own_domains() -> List[str]:
    """Important domains configured in GURI (datastore/important_domains.txt)."""
    try:
        from pathlib import Path

        from geofooter.paths import datastore_path

        path = Path(datastore_path("important_domains.txt"))
        if not path.is_file():
            return []
        out = []
        for line in path.read_text(encoding="utf-8", errors="ignore").splitlines():
            line = line.strip().lower().lstrip("@")
            if line and not line.startswith("#") and "." in line:
                out.append(line)
        return out
    except Exception:
        return []


def run_all(
    headers: str,
    *,
    dmarc: Optional[str] = None,
    own_domains: Optional[Sequence[str]] = None,
    lite: bool = True,
    hops: Optional[List[Dict[str, Any]]] = None,
) -> List[IntelSection]:
    """Every header check, in report order. ``lite`` skips DNS lookups."""
    if not headers:
        return []
    own = list(own_domains) if own_domains is not None else load_own_domains()
    return [identity_signals(headers, dmarc=dmarc, own_domains=own)]
