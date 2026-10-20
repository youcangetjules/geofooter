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


# --------------------------------------------------------------------------- #
# ESP / tenant fingerprinting
# --------------------------------------------------------------------------- #

ESP_ASNS: Dict[str, Tuple[str, ...]] = {
    "sendgrid": ("AS11377",),
    "mailgun": ("AS395085",),
    "amazonses": ("AS16509", "AS14618"),
    "postmark": ("AS395425",),
    "mailchimp": ("AS10868", "AS27357"),
    "salesforce": ("AS14340", "AS22606"),
}
ESP_NAMES = {
    "sendgrid": "SendGrid",
    "mailgun": "Mailgun",
    "amazonses": "Amazon SES",
    "postmark": "Postmark",
    "mailchimp": "Mailchimp / Mandrill",
    "salesforce": "Salesforce",
}


def _b64_digits(token: str) -> str:
    """SendGrid writes the account ID as base64 digits (MjE3MzIzODk -> 21732389)."""
    import base64
    import binascii

    tok = (token or "").strip()
    if not re.fullmatch(r"[A-Za-z0-9+/=]{4,24}", tok):
        return ""
    try:
        raw = base64.b64decode(tok + "=" * (-len(tok) % 4), validate=True)
        text = raw.decode("ascii")
    except (binascii.Error, ValueError, UnicodeDecodeError):
        return ""
    return text if text.isdigit() else ""


def _tenants_path():
    import os
    from pathlib import Path

    return Path(os.environ.get("LOCALAPPDATA", "")) / "GeoFooter" / "aes_esp_tenants.json"


def load_tenant_lists() -> Dict[str, List[str]]:
    """``{"trusted": ["sendgrid:21732389"], "blocked": []}`` from LocalAppData."""
    import json

    out: Dict[str, List[str]] = {"trusted": [], "blocked": []}
    try:
        path = _tenants_path()
        if path.is_file():
            data = json.loads(path.read_text(encoding="utf-8"))
            for key in out:
                out[key] = [str(v).strip().lower() for v in data.get(key) or [] if str(v).strip()]
    except Exception:
        pass
    return out


def esp_fingerprint(
    headers: str,
    hops: Optional[List[Dict[str, Any]]] = None,
    tenants: Optional[Dict[str, List[str]]] = None,
) -> IntelSection:
    sec = IntelSection("Sending platform")
    flat = unfold(headers)
    low = flat.lower()
    hops = hops or []
    esp = ""
    tenant = ""
    campaign = ""
    details: List[str] = []

    def hv(name: str) -> str:
        return first_header(headers, name)

    rp = hv("Return-Path").lower()
    feedback = hv("Feedback-ID") or hv("X-Feedback-ID")

    if (hv("X-SG-EID") or hv("X-SG-ID") or "sendgrid.net" in low
            or re.search(r"\bby\s+\S*ismtpd\S*\s+\(sg\)", low)
            or re.match(r"u\d{3}\.", hv("X-Entity-ID") or "")):
        esp = "sendgrid"
        m = re.search(r"bounces\+(\d+)-", rp)
        if m:
            tenant = m.group(1)
        for hop in hops:
            acct = _b64_digits(str(hop.get("fqdn") or ""))
            if acct and "ismtpd" in str(hop.get("by_host") or "").lower():
                tenant = tenant or acct
                hop["decoded_location"] = f"SendGrid API, account {acct}"
                if str(hop.get("asn_org") or "Unknown") in {
                    "Unknown", "", "—", "NA", "N/A", "Insufficient header data", "Internal / local hop",
                }:
                    hop["asn_org"] = "SendGrid (HTTP API)"
        if not tenant and feedback:
            m = re.match(r"\s*(\d+):", feedback)
            tenant = m.group(1) if m else ""
    elif any(hv(h) for h in ("X-Mailgun-Sid", "X-Mailgun-Variables", "X-Mailgun-Sending-Ip")):
        esp = "mailgun"
        tenant = (hv("X-Mailgun-Sid") or "")[:40]
        sip = hv("X-Mailgun-Sending-Ip")
        if sip:
            details.append(f"Mailgun sending IP {sip}")
    elif hv("X-SES-Outgoing") or "amazonses.com" in low:
        esp = "amazonses"
        m = re.match(r"\s*([^:]+(?::[^:]+)?):amazonses", feedback or "", re.I)
        tenant = m.group(1) if m else ""
    elif hv("X-PM-Message-Id") or hv("X-PM-Tag") or "mtasv.net" in low:
        esp = "postmark"
        campaign = hv("X-PM-Tag")
    elif hv("X-MC-User") or hv("X-Mandrill-User") or "mcsv.net" in low or "mandrillapp.com" in low:
        esp = "mailchimp"
        tenant = hv("X-MC-User") or hv("X-Mandrill-User")
    elif hv("X-SFDC-LK") or hv("X-SFDC-User") or "salesforce.com" in rp:
        esp = "salesforce"
        tenant = hv("X-SFDC-LK")

    if not esp:
        return sec
    if feedback and not campaign:
        campaign = feedback
    tenant = tenant.strip()
    name = ESP_NAMES.get(esp, esp)
    key = f"{esp}:{tenant}".lower() if tenant else ""
    sec.data.update(esp=esp, tenant=tenant, key=key)
    sec.rows.append(f"{name}" + (f", account {tenant}" if tenant else ""))
    if campaign:
        sec.rows.append(f"Campaign / feedback ID: {campaign[:120]}")
    sec.rows.extend(details)

    lists = tenants if tenants is not None else load_tenant_lists()
    if key and key in lists.get("blocked", []):
        sec.add(f"{name} account {tenant} is on your blocked tenant list", 25,
                "Blocked sender account on a shared sending platform")
    elif key and key in lists.get("trusted", []):
        sec.rows.append("Trusted tenant: shared-IP route risk not scored for its hops")
        asns = {a.upper() for a in ESP_ASNS.get(esp, ())}
        for hop in hops:
            org = str(hop.get("asn_org") or "").lower()
            if str(hop.get("asn") or "").upper() in asns or name.split()[0].lower() in org:
                hop["esp_trusted"] = True
    elif key:
        sec.rows.append(f"Tenant key {key} (add to aes_esp_tenants.json to trust or block)")
    return sec


# --------------------------------------------------------------------------- #
# DKIM signatures, identifier alignment, DMARC policy
# --------------------------------------------------------------------------- #

def _tags(value: str) -> Dict[str, str]:
    out: Dict[str, str] = {}
    for part in value.split(";"):
        if "=" in part:
            k, v = part.split("=", 1)
            out[k.strip().lower()] = re.sub(r"\s+", "", v)
    return out


def dkim_signatures(headers: str) -> List[Dict[str, Any]]:
    sigs = []
    for val in header_values(headers, "DKIM-Signature"):
        t = _tags(val)
        sigs.append(
            {
                "d": t.get("d", "").lower(),
                "s": t.get("s", ""),
                "a": t.get("a", "").lower(),
                "h": [h.strip().lower() for h in t.get("h", "").split(":") if h.strip()],
            }
        )
    return sigs


def _alignment(dom: str, from_dom: str) -> str:
    dom, from_dom = (dom or "").lower(), (from_dom or "").lower()
    if not dom or not from_dom:
        return ""
    if dom == from_dom:
        return "strict"
    if registrable_domain(dom) == registrable_domain(from_dom):
        return "relaxed"
    return "not aligned"


def _der_rsa_bits(der: bytes) -> int:
    """Modulus size of the first large INTEGER in a DER public key (SPKI or PKCS#1)."""
    best = 0

    def walk(buf: bytes, depth: int = 0) -> None:
        nonlocal best
        i = 0
        while i + 2 <= len(buf) and depth < 6:
            tag = buf[i]
            ln = buf[i + 1]
            i += 2
            if ln & 0x80:
                n = ln & 0x7F
                if n == 0 or i + n > len(buf):
                    return
                ln = int.from_bytes(buf[i:i + n], "big")
                i += n
            body = buf[i:i + ln]
            if tag == 0x02:
                val = body.lstrip(b"\x00")
                best = max(best, len(val) * 8 - (8 - val[0].bit_length() if val else 0))
            elif tag == 0x30:
                walk(body, depth + 1)
            elif tag == 0x03 and body:
                walk(body[1:], depth + 1)
            i += ln

    try:
        walk(der)
    except Exception:
        return 0
    return best


def _cache_path():
    import os
    from pathlib import Path

    return Path(os.environ.get("LOCALAPPDATA", "")) / "GeoFooter" / "dns_auth_cache.json"


def _dns_txt(name: str, *, ttl_s: int = 86400) -> Optional[List[str]]:
    """TXT strings for *name*, cached 24h. None when the lookup itself failed."""
    import json
    import time

    path = _cache_path()
    try:
        cache = json.loads(path.read_text(encoding="utf-8")) if path.is_file() else {}
    except Exception:
        cache = {}
    hit = cache.get(name)
    if hit and time.time() - float(hit.get("t", 0)) < ttl_s:
        return list(hit.get("txt") or [])
    try:
        import dns.resolver

        res = dns.resolver.Resolver()
        res.timeout, res.lifetime = 3, 5
        try:
            ans = res.resolve(name, "TXT")
            txt = ["".join(p.decode("utf-8", "ignore") for p in r.strings) for r in ans]
        except (dns.resolver.NXDOMAIN, dns.resolver.NoAnswer):
            txt = []
    except Exception:
        return None
    cache[name] = {"t": time.time(), "txt": txt}
    try:
        path.parent.mkdir(parents=True, exist_ok=True)
        tmp = path.with_suffix(".tmp")
        tmp.write_text(json.dumps(cache), encoding="utf-8")
        tmp.replace(path)
    except OSError:
        pass
    return txt


def dkim_key_bits(selector: str, domain: str, *, resolver=None) -> Tuple[int, str]:
    """(bits, key type) for ``selector._domainkey.domain``; (0, "") if unknown."""
    import base64

    txts = (resolver or _dns_txt)(f"{selector}._domainkey.{domain}")
    for txt in txts or []:
        t = _tags(txt)
        if "p" not in t:
            continue
        kind = t.get("k", "rsa").lower()
        if not t["p"]:
            return 0, "revoked"
        if kind == "ed25519":
            return 256, "ed25519"
        try:
            der = base64.b64decode(t["p"] + "=" * (-len(t["p"]) % 4))
        except Exception:
            return 0, kind
        try:
            from cryptography.hazmat.primitives.serialization import (
                load_der_public_key,
            )

            return int(getattr(load_der_public_key(der), "key_size", 0) or 0), kind
        except Exception:
            # Older selectors publish bare PKCS#1, which load_der_public_key rejects.
            return _der_rsa_bits(der), kind
    return 0, ""


def _dmarc_from_auth(headers: str) -> Dict[str, str]:
    for val in header_values(headers, "Authentication-Results"):
        m = re.search(r"\bdmarc=\w+\s*\(([^)]*)\)", val, re.I)
        if m:
            t = {k.lower(): v.lower() for k, v in re.findall(r"\b(p|sp|pct)=([\w]+)", m.group(1), re.I)}
            if t.get("p"):
                return t
    return {}


def signing_signals(
    headers: str,
    *,
    lite: bool = True,
    resolver=None,
) -> IntelSection:
    sec = IntelSection("Signing / alignment")
    _n, from_addr = _addr(first_header(headers, "From"))
    from_dom = _domain_of(from_addr)
    sigs = dkim_signatures(headers)

    if sigs:
        doms = []
        for s in sigs:
            if s["d"] and s["d"] not in doms:
                doms.append(s["d"])
        sec.rows.append("Signed by " + " and ".join(doms) if doms else "DKIM signature without d=")
        for s in sigs:
            sec.rows.append(
                f"d={s['d']} s={s['s']} a={s['a'] or '?'}"
                + (f" ({_alignment(s['d'], from_dom)} with From)" if from_dom else "")
            )
        main = next((s for s in sigs if _alignment(s["d"], from_dom) in {"strict", "relaxed"}), sigs[0])
        missing = [h for h in ("from", "subject") if h not in main["h"]]
        if first_header(headers, "Reply-To") and "reply-to" not in main["h"]:
            missing.append("reply-to")
        if missing:
            sec.add(
                f"DKIM signature for {main['d']} does not cover {', '.join(missing)}",
                5,
                "Those headers could be changed on a replayed copy without breaking the signature",
            )
    else:
        sec.rows.append("No DKIM signature")

    mailfrom = ""
    for val in header_values(headers, "Authentication-Results"):
        m = re.search(r"smtp\.mailfrom=\"?([^\s;\"]+)", val, re.I)
        if m:
            mailfrom = m.group(1).lower()
            break
    if mailfrom and from_dom:
        mf_dom = _domain_of(mailfrom) or mailfrom
        sec.rows.append(f"SPF domain {mf_dom} ({_alignment(mf_dom, from_dom)} with From)")

    policy = _dmarc_from_auth(headers)
    reg = registrable_domain(from_dom)
    if not lite and reg:
        for txt in (resolver or _dns_txt)(f"_dmarc.{reg}") or []:
            if txt.lower().startswith("v=dmarc1"):
                t = _tags(txt)
                policy = {k: t[k].lower() for k in ("p", "sp", "pct") if k in t} or policy
                break
        for s in sigs:
            if not (s["d"] and s["s"]):
                continue
            bits, kind = dkim_key_bits(s["s"], s["d"], resolver=resolver)
            if kind == "revoked":
                sec.rows.append(f"DKIM key {s['s']}._domainkey.{s['d']} is revoked (empty p=)")
            elif bits:
                sec.rows.append(f"DKIM key {s['s']}._domainkey.{s['d']}: {kind.upper()} {bits}-bit")
                if kind == "rsa" and bits < 1024:
                    sec.add(f"Weak DKIM key for {s['d']} ({bits}-bit RSA)", 5,
                            "Keys under 1024 bits can be factored; the signature proves little")
    if policy.get("p"):
        line = f"DMARC policy for {reg or from_dom}: p={policy['p']}"
        if policy.get("sp"):
            line += f" sp={policy['sp']}"
        if policy.get("pct"):
            line += f" pct={policy['pct']}"
        if policy["p"] == "none":
            line += " (monitoring only, spoofed mail is not rejected)"
        sec.rows.append(line)
    elif not lite and reg:
        sec.rows.append(f"No DMARC record published for {reg}")
    return sec


# --------------------------------------------------------------------------- #
# Transport (TLS), sending client, bulk / automated mail
# --------------------------------------------------------------------------- #

_PLAIN_PROTOCOLS = {"SMTP", "ESMTP", "ESMTPA"}


def _norm_tls(v: str) -> str:
    v = (v or "").upper().replace("TLSV", "TLS").replace("_", ".")
    m = re.match(r"(TLS|SSL)V?\s*(\d(?:\.\d)?)", v)
    return f"{m.group(1)} {m.group(2)}" if m else ""


def parse_tls(received: str) -> Tuple[str, str]:
    """(version, cipher) from one Received line, e.g. ('TLS 1.3', 'TLS_AES_256_GCM_SHA384')."""
    line = received or ""
    version = ""
    for pat in (
        r"version=(TLS[\w.]+|SSL\w+)",
        r"\busing\s+(TLSv?[\d.]+|SSLv\d)",
        r"\((TLS1[._]\d|TLSv1(?:\.\d)?)\)",
    ):
        m = re.search(pat, line, re.I)
        if m:
            version = _norm_tls(m.group(1))
            break
    m = re.search(r"\bcipher[=\s]+([A-Z0-9][A-Z0-9_-]{6,})", line, re.I) or re.search(
        r"\btls\s+(TLS_[A-Z0-9_]+|ECDHE[A-Z0-9_-]+)", line, re.I
    )
    return version, (m.group(1) if m else "")


def transport_signals(hops: Optional[List[Dict[str, Any]]]) -> IntelSection:
    sec = IntelSection("Transport")
    weak, plain = [], []
    for idx, hop in enumerate(hops or [], 1):
        if hop.get("hop_kind") != "public":
            continue
        ver = str(hop.get("tls_version") or "")
        proto = str(hop.get("protocol") or "").upper()
        if ver in {"TLS 1.0", "TLS 1.1", "TLS 1", "SSL 3"} or ver.startswith("SSL"):
            weak.append(f"hop {idx} {ver}")
        elif proto in _PLAIN_PROTOCOLS and not ver and hop.get("fqdn"):
            # "Received: by <host> with SMTP" with no "from" is a provider's
            # internal handoff (Gmail writes these), not a hop across the internet.
            plain.append(f"hop {idx}")
        if hop.get("tls"):
            sec.rows.append(f"Hop {idx}: {hop.get('tls')}")
    if weak:
        sec.add(f"Outdated TLS on an external hop ({', '.join(weak)})", 5,
                "TLS 1.0/1.1 points at a poorly maintained or bespoke mail server")
    if plain:
        sec.add(f"Unencrypted SMTP on an external hop ({', '.join(plain)})", 3,
                "Mail crossed the internet in plain text")
    return sec


_KIT_MAILERS = [
    (re.compile(r"phpmailer\s*\[?v?\s*([0-5])\.", re.I), "old PHPMailer"),
    (re.compile(r"swift\s*mailer\s*v?\s*([0-5])\.", re.I), "old SwiftMailer"),
    (re.compile(r"^\s*php/\d", re.I), "raw PHP mail()"),
    (re.compile(r"^\s*(python|perl|ruby|go-http-client|curl|powershell)\b", re.I), "script library"),
    (re.compile(r"microsoft cdo for windows 2000", re.I), "legacy CDO script"),
]
_WIN_HOST = re.compile(r"^(WIN|DESKTOP|LAPTOP)-[A-Z0-9]{5,}", re.I)


def client_signals(headers: str) -> IntelSection:
    sec = IntelSection("Sending client")
    mid = first_header(headers, "Message-ID").strip().strip("<>")
    mid_host = mid.rsplit("@", 1)[1] if "@" in mid else ""
    mailer = first_header(headers, "X-Mailer")
    agent = first_header(headers, "User-Agent")
    orig_ip = first_header(headers, "X-Originating-IP").strip("[] ")
    if mid_host:
        sec.rows.append(f"Message-ID host: {mid_host}")
    if mailer:
        sec.rows.append(f"X-Mailer: {mailer[:120]}")
    if agent:
        sec.rows.append(f"User-Agent: {agent[:120]}")
    if orig_ip:
        sec.rows.append(f"X-Originating-IP: {orig_ip}")
    for value in (mailer, agent):
        hit = next((label for rx, label in _KIT_MAILERS if value and rx.search(value)), "")
        if hit:
            sec.add(f"Sent with {hit} ({value[:60]})", 10,
                    "Phishing kits rely on these mailers; legitimate senders rarely do")
            break
    if mid_host and _WIN_HOST.match(mid_host):
        sec.add(f"Message-ID names a Windows machine ({mid_host})", 5,
                "Mail generated on a desktop PC, typical of a script or RAT")
    return sec


def bulk_signals(headers: str, *, esp: str = "") -> IntelSection:
    sec = IntelSection("Mail type")
    lu = " ".join(header_values(headers, "List-Unsubscribe"))
    one_click = bool(first_header(headers, "List-Unsubscribe-Post"))
    precedence = first_header(headers, "Precedence").lower()
    auto = first_header(headers, "Auto-Submitted").lower()
    urls = re.findall(r"<\s*(https?://[^>\s]+)\s*>", lu)
    mailtos = re.findall(r"<\s*(mailto:[^>\s]+)\s*>", lu)
    if lu or precedence in {"bulk", "list"}:
        kind = "Bulk / mailing list"
    elif auto and auto != "no" or precedence == "junk":
        kind = "Automated"
    elif esp:
        kind = f"Transactional (sent through {ESP_NAMES.get(esp, esp)}, no list headers)"
    else:
        kind = "Personal"
    sec.rows.append(kind)
    if lu:
        sec.rows.append("List-Unsubscribe" + (" with one-click (RFC 8058)" if one_click else ""))
        for u in urls[:3]:
            sec.rows.append(f"Unsubscribe link: {u[:140]}")
        for u in mailtos[:2]:
            sec.rows.append(f"Unsubscribe address: {u[:140]}")
    if auto and auto != "no":
        sec.rows.append(f"Auto-Submitted: {auto}")
    sec.data.update(kind=kind, unsubscribe_urls=urls)
    return sec


# --------------------------------------------------------------------------- #
# Conversation hijack (claims to continue a thread nobody saw)
# --------------------------------------------------------------------------- #

_REPLY_PREFIX = re.compile(r"^\s*(re|aw|sv|antw)\s*(\[\d+\])?\s*:", re.I)
_FWD_PREFIX = re.compile(r"^\s*(fw|fwd|wg|tr)\s*:", re.I)
HISTORY_MIN_DAYS = 14


def _decoded_subject(headers: str) -> str:
    from email.header import decode_header, make_header

    raw = first_header(headers, "Subject")
    try:
        return str(make_header(decode_header(raw)))
    except Exception:
        return raw


def message_id_key(msg_id: str) -> str:
    """Same key ``record_sender_email`` stores in sender history ``seen``."""
    import hashlib

    mid = (msg_id or "").strip().lower()
    return hashlib.sha1(mid.encode("utf-8")).hexdigest()[:16] if mid else ""


def _history_days(history: Dict[str, Any]) -> float:
    from datetime import datetime

    stamps = [str(v.get("first_seen") or "") for v in (history.get("senders") or {}).values()]
    stamps = [s for s in stamps if s]
    if not stamps:
        return 0.0
    try:
        oldest = datetime.strptime(min(stamps), "%Y-%m-%d %H:%M:%S")
    except ValueError:
        return 0.0
    return (datetime.now() - oldest).total_seconds() / 86400


def load_known_contacts() -> List[str]:
    """Addresses the user has sent to (GURI recipients list), lower-case."""
    try:
        from pathlib import Path

        from geofooter.paths import datastore_path

        path = Path(datastore_path("recipients.txt"))
        if not path.is_file():
            return []
        return [
            line.strip().lower()
            for line in path.read_text(encoding="utf-8", errors="ignore").splitlines()
            if "@" in line
        ]
    except Exception:
        return []


def thread_signals(
    headers: str,
    *,
    sender_email: str,
    history: Optional[Dict[str, Any]],
    known_contacts: Iterable[str] = (),
    own_domains: Iterable[str] = (),
) -> IntelSection:
    sec = IntelSection("Thread")
    subject = _decoded_subject(headers)
    is_reply = bool(_REPLY_PREFIX.match(subject))
    is_fwd = bool(_FWD_PREFIX.match(subject))
    ids = re.findall(r"<[^<>\s]+>", " ".join(header_values(headers, "In-Reply-To")
                                             + header_values(headers, "References")))
    ids = list(dict.fromkeys(i.lower() for i in ids))
    if not (is_reply or is_fwd or ids):
        return sec

    history = history or {}
    seen = history.get("seen") or {}
    sender = (sender_email or "").strip().lower()
    sender_count = int(((history.get("senders") or {}).get(sender) or {}).get("count", 0) or 0)
    if message_id_key(first_header(headers, "Message-ID")) in seen:
        # A rescan: this message is already counted, so leave it out.
        sender_count = max(0, sender_count - 1)
    known = {c.lower() for c in known_contacts}
    own = {registrable_domain(d) for d in own_domains}
    s_reg = registrable_domain(_domain_of(sender))
    contact = sender in known or (s_reg and s_reg in own)
    found = sum(1 for i in ids if message_id_key(i) in seen)

    if ids:
        sec.rows.append(
            f"{'Reply' if is_reply else 'Message'} references {len(ids)} earlier message(s); "
            f"{found} seen by AES"
        )
    elif is_reply:
        sec.rows.append("Subject says reply, but there are no In-Reply-To or References headers")
    if sender_count == 0:
        sec.rows.append("First message AES has seen from this sender")
    if contact:
        sec.rows.append("Sender is a known contact (recipients list or own domain)")

    days = _history_days(history)
    if days < HISTORY_MIN_DAYS:
        sec.rows.append(f"Sender history covers {days:.0f} days; thread checks are report-only until {HISTORY_MIN_DAYS}")
        return sec
    if sender_count or contact or not is_reply:
        # Newsletters and read receipts carry References too; only a
        # message that presents itself as a reply makes the claim.
        return sec
    auto = first_header(headers, "Auto-Submitted").lower()
    if header_values(headers, "List-Unsubscribe") or (auto and auto != "no"):
        return sec
    if ids and found == 0:
        sec.add(
            "Claims to continue a thread we never saw",
            15,
            "First-time sender replying to messages AES has no record of; a conversation-hijack pattern",
        )
    elif is_reply and not ids:
        sec.add(
            "'Re:' subject with no thread headers from a first-time sender",
            10,
            "Fake reply used to make a cold message look like an existing conversation",
        )
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
    sender_email: str = "",
    history: Optional[Dict[str, Any]] = None,
) -> List[IntelSection]:
    """Every header check, in report order. ``lite`` skips DNS lookups.

    ``history`` is the sender-history dict read *before* this message is
    recorded, so a first-time sender still has a count of zero.
    """
    if not headers:
        return []
    own = list(own_domains) if own_domains is not None else load_own_domains()
    esp = esp_fingerprint(headers, hops)
    sections = [
        identity_signals(headers, dmarc=dmarc, own_domains=own),
        esp,
        signing_signals(headers, lite=lite),
        transport_signals(hops),
        client_signals(headers),
        bulk_signals(headers, esp=str(esp.data.get("esp") or "")),
    ]
    if history is not None:
        sections.append(
            thread_signals(
                headers,
                sender_email=sender_email,
                history=history,
                known_contacts=load_known_contacts(),
                own_domains=own,
            )
        )
    return sections
