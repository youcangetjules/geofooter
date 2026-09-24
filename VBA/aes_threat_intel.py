#!/usr/bin/env python3
"""AES threat-intelligence lookups for IPs, domains, URLs and file hashes.

Providers (keyless ones always run; keyed ones run when a key is stored via
``aes_secrets.set_secret(<name>, key)`` or ``AES_<NAME>_API_KEY``):

  ip      internetdb (Shodan, keyless) · shodan · virustotal · greynoise · otx · threatfox
  domain  dnsbl (Spamhaus DBL / SURBL / URIBL, keyless) · virustotal · otx · urlhaus · threatfox
  url     gsb (Google Safe Browsing, batched) · urlhaus · virustotal · urlscan (deep only)
  hash    virustotal · malwarebazaar · otx

Privacy: nothing is uploaded. Files are looked up by SHA-256 only; URLs are
looked up, never submitted — except URLScan in deep mode, which submits
already-suspicious URLs with *private* visibility (configurable).

Everything runs in a thread pool under a per-scan time budget, results are
cached on disk, and VirusTotal's free-tier 4/min limit is enforced across
concurrent scan processes.
"""
from __future__ import annotations

import base64
import json
import logging
import os
import re
import threading
import time
from concurrent.futures import ThreadPoolExecutor, wait as futures_wait
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any, Callable, Dict, Iterable, List, Optional, Sequence, Tuple
from urllib.parse import quote, urlparse

logger = logging.getLogger("aes_threat_intel")

try:
    from aes_secrets import get_secret
except ImportError:  # pragma: no cover
    def get_secret(name: str) -> str:  # type: ignore[misc]
        return ""

PROVIDERS: Dict[str, Dict[str, Any]] = {
    "internetdb": {"label": "Shodan InternetDB", "key": None},
    "dnsbl": {"label": "Domain blocklists (DBL/SURBL/URIBL)", "key": None},
    "shodan": {"label": "Shodan", "key": "shodan", "signup": "https://account.shodan.io/"},
    "virustotal": {"label": "VirusTotal", "key": "virustotal", "signup": "https://www.virustotal.com/gui/my-apikey"},
    "gsb": {"label": "Google Safe Browsing", "key": "gsb", "signup": "https://developers.google.com/safe-browsing/v4/get-started"},
    "greynoise": {"label": "GreyNoise", "key": "greynoise", "signup": "https://viz.greynoise.io/account/api-key"},
    "otx": {"label": "AlienVault OTX", "key": "otx", "signup": "https://otx.alienvault.com/api"},
    "abusech": {"label": "abuse.ch (URLhaus / ThreatFox / MalwareBazaar)", "key": "abusech", "signup": "https://auth.abuse.ch/"},
    "urlscan": {"label": "urlscan.io", "key": "urlscan", "signup": "https://urlscan.io/user/profile/"},
}

MODE_BUDGET_S = {"compact": 8.0, "full": 20.0, "deep": 75.0}
VT_PER_SCAN = {"compact": 2, "full": 4, "deep": 4}
VT_PER_MINUTE = 4
VT_PER_DAY = 480

_VERDICT_ORDER = {"error": 0, "unknown": 1, "clean": 2, "suspicious": 3, "malicious": 4}

# Big shared platforms: blocklist DNS is cheap so it still runs, but quota-limited
# providers skip them (a verdict on google.com says nothing about this mail).
_QUOTA_SKIP_DOMAINS = frozenset({
    "google.com", "gmail.com", "microsoft.com", "outlook.com", "office.com", "live.com",
    "office365.com", "apple.com", "icloud.com", "linkedin.com", "facebook.com",
    "youtube.com", "twitter.com", "x.com", "instagram.com", "gstatic.com",
    "googleapis.com", "microsoftonline.com", "windows.com", "bing.com",
})

_MULTI_PART_SUFFIXES = frozenset({
    "co.uk", "org.uk", "ac.uk", "gov.uk", "me.uk", "ltd.uk", "plc.uk", "com.au", "net.au",
    "org.au", "co.nz", "co.jp", "ne.jp", "or.jp", "com.br", "com.cn", "com.hk", "com.sg",
    "co.za", "co.in", "com.mx", "com.tr", "co.kr", "com.tw", "com.ar", "co.il",
})

_BAD_SHODAN_TAGS = {"malware": 15, "c2": 15, "compromised": 15, "botnet": 15,
                    "tor": 10, "honeypot": 5, "scanner": 5, "proxy": 5, "vpn": 3}


# --------------------------------------------------------------------------- #
# Result types
# --------------------------------------------------------------------------- #

@dataclass
class ProviderResult:
    provider: str
    verdict: str = "unknown"  # clean | suspicious | malicious | unknown | error
    summary: str = ""
    detail: Dict[str, Any] = field(default_factory=dict)
    link: str = ""

    def to_dict(self) -> Dict[str, Any]:
        return {"provider": self.provider, "verdict": self.verdict,
                "summary": self.summary, "detail": self.detail, "link": self.link}


@dataclass
class IndicatorReport:
    indicator: str
    kind: str  # ip | domain | url | hash
    role: str = ""  # origin-ip, hop-ip, sender-domain, link-url, attachment …
    results: List[ProviderResult] = field(default_factory=list)

    @property
    def worst(self) -> str:
        best = "unknown"
        for r in self.results:
            if _VERDICT_ORDER.get(r.verdict, 0) > _VERDICT_ORDER.get(best, 0):
                best = r.verdict
        return best

    def lines(self) -> List[str]:
        return [f"{PROVIDERS.get(r.provider, {}).get('label', r.provider)}: {r.summary}"
                for r in self.results if r.summary]

    def to_dict(self) -> Dict[str, Any]:
        return {"indicator": self.indicator, "kind": self.kind, "role": self.role,
                "worst": self.worst, "results": [r.to_dict() for r in self.results]}


# --------------------------------------------------------------------------- #
# Keys, HTTP, cache, rate limits
# --------------------------------------------------------------------------- #

def provider_key(provider: str) -> str:
    meta = PROVIDERS.get(provider) or {}
    name = meta.get("key")
    if not name:
        return ""
    env = (os.environ.get(f"AES_{name.upper()}_API_KEY") or "").strip()
    return env or (get_secret(name) or "").strip()


def provider_enabled(provider: str) -> bool:
    meta = PROVIDERS.get(provider) or {}
    return meta.get("key") is None or bool(provider_key(provider))


def _http(method: str, url: str, *, timeout: float = 6.0, **kwargs: Any):
    """requests call that tolerates corporate TLS interception (same policy as aes_http_get)."""
    import requests
    from requests import exceptions as req_exc

    headers = dict(kwargs.pop("headers", None) or {})
    headers.setdefault("User-Agent", "Aliniant-AES/1.0")
    try:
        return requests.request(method, url, headers=headers, timeout=timeout, **kwargs)
    except req_exc.SSLError:
        try:
            import urllib3
            urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)
        except Exception:
            pass
        return requests.request(method, url, headers=headers, timeout=timeout, verify=False, **kwargs)


def _state_dir() -> Path:
    local = os.environ.get("LOCALAPPDATA", "")
    base = Path(local) / "GeoFooter" if local else Path(r"C:\GeoFooter")
    base.mkdir(parents=True, exist_ok=True)
    return base


class _Cache:
    """Small JSON cache shared across scan processes (merge-on-save, atomic replace)."""

    TTL = {"malicious": 6 * 3600, "suspicious": 12 * 3600, "clean": 24 * 3600,
           "unknown": 12 * 3600, "error": 600}

    def __init__(self, path: Optional[Path] = None):
        self.path = path or (_state_dir() / "threat_intel_cache.json")
        self.lock = threading.Lock()
        self.data: Dict[str, Any] = self._load()
        self.dirty: Dict[str, Any] = {}

    def _load(self) -> Dict[str, Any]:
        try:
            if self.path.is_file():
                raw = json.loads(self.path.read_text(encoding="utf-8"))
                if isinstance(raw, dict):
                    return raw
        except Exception:
            pass
        return {}

    def get(self, key: str) -> Optional[ProviderResult]:
        with self.lock:
            entry = (self.data.get("results") or {}).get(key)
        if not isinstance(entry, dict):
            return None
        age = time.time() - float(entry.get("ts") or 0)
        res = entry.get("r") or {}
        if age > self.TTL.get(res.get("verdict", "unknown"), 3600):
            return None
        result = ProviderResult(
            provider=res.get("provider", ""), verdict=res.get("verdict", "unknown"),
            summary=res.get("summary", ""), detail=res.get("detail") or {}, link=res.get("link", ""),
        )
        # Re-score cached VirusTotal engine stats with the current consensus thresholds
        # so a prior 1/91 "suspicious" does not stick until TTL expires.
        if result.provider == "virustotal":
            stats = (result.detail or {}).get("stats")
            if isinstance(stats, dict) and stats:
                kind = "hash"
                if "|url|" in key:
                    kind = "url"
                elif "|domain|" in key:
                    kind = "domain"
                elif "|ip|" in key:
                    kind = "ip"
                result.verdict, result.summary = _vt_stats_verdict(stats, kind=kind)
                label = (result.detail or {}).get("threat_label")
                if label:
                    result.summary = f"{result.summary} — {label}"
        return result

    def put(self, key: str, result: ProviderResult) -> None:
        entry = {"ts": time.time(), "r": result.to_dict()}
        with self.lock:
            self.data.setdefault("results", {})[key] = entry
            self.dirty[key] = entry

    def take_vt_slot(self) -> bool:
        """Cross-process VirusTotal quota: VT_PER_MINUTE / VT_PER_DAY."""
        with self.lock:
            fresh = self._load()
            calls = [t for t in (fresh.get("vt_calls") or []) if time.time() - t < 86400]
            last_min = [t for t in calls if time.time() - t < 60]
            if len(last_min) >= VT_PER_MINUTE or len(calls) >= VT_PER_DAY:
                return False
            calls.append(time.time())
            fresh["vt_calls"] = calls
            fresh.setdefault("results", {}).update(self.dirty)
            self._write(fresh)
            self.data = fresh
            return True

    def save(self) -> None:
        with self.lock:
            if not self.dirty:
                return
            fresh = self._load()
            results = fresh.setdefault("results", {})
            results.update(self.dirty)
            cutoff = time.time() - 3 * 86400
            for k in [k for k, v in results.items() if float((v or {}).get("ts") or 0) < cutoff]:
                results.pop(k, None)
            self._write(fresh)
            self.dirty = {}

    def _write(self, payload: Dict[str, Any]) -> None:
        tmp = self.path.with_suffix(f".{os.getpid()}.tmp")
        try:
            tmp.write_text(json.dumps(payload), encoding="utf-8")
            os.replace(tmp, self.path)
        except OSError as exc:
            logger.debug("threat intel cache write failed: %s", exc)
            try:
                tmp.unlink()
            except OSError:
                pass


# --------------------------------------------------------------------------- #
# Helpers
# --------------------------------------------------------------------------- #

def registrable_domain(host: str) -> str:
    host = (host or "").strip().lower().rstrip(".")
    if not host or re.match(r"^[\d.]+$", host) or ":" in host:
        return host
    labels = host.split(".")
    if len(labels) >= 3 and ".".join(labels[-2:]) in _MULTI_PART_SUFFIXES:
        return ".".join(labels[-3:])
    return ".".join(labels[-2:])


def _vt_url_id(url: str) -> str:
    return base64.urlsafe_b64encode(url.encode("utf-8")).decode("ascii").strip("=")


def _vt_stats_verdict(stats: Dict[str, Any], *, kind: str = "url") -> Tuple[str, str]:
    """Multi-engine consensus for VirusTotal — lone outliers are not suspicious.

    Popular sites often show 1/90+ detections. We score by absolute hits *and*
    share of engines; weak ratios stay clean (with the raw counts in the summary).
    File hashes stay slightly more sensitive than URL/domain/IP.
    """
    mal = int(stats.get("malicious") or 0)
    sus = int(stats.get("suspicious") or 0)
    total = sum(int(v or 0) for v in stats.values() if isinstance(v, (int, float)))
    if total <= 0:
        return "unknown", "no engine results"
    flagged = mal + sus
    ratio = flagged / float(total)
    mal_ratio = mal / float(total)
    summary = f"{mal} malicious, {sus} suspicious of {total} engines"

    if kind == "hash":
        if mal >= 5 or (mal >= 3 and mal_ratio >= 0.05):
            return "malicious", summary
        if mal >= 2 or (mal >= 1 and sus >= 1) or sus >= 3:
            return "suspicious", summary
        if flagged:
            return "clean", f"{summary} (below hash threshold)"
        return "clean", f"not flagged ({total} engines)"

    # URL / domain / IP (RiskTable): require real consensus. 1/91 stays clean.
    if mal >= 8 or (mal >= 5 and mal_ratio >= 0.05) or (mal >= 4 and mal_ratio >= 0.08):
        return "malicious", summary
    if (
        (flagged >= 5 and ratio >= 0.04)
        or (flagged >= 4 and ratio >= 0.06)
        or (mal >= 3 and ratio >= 0.03)
        or (mal >= 2 and sus >= 2)
    ):
        return "suspicious", summary
    if flagged:
        pct = round(ratio * 100, 1)
        return "clean", f"{summary} ({pct}% — below RiskTable threshold)"
    return "clean", f"not flagged ({total} engines)"


def _is_ipv6(ip: str) -> bool:
    return ":" in (ip or "")


# --------------------------------------------------------------------------- #
# Providers — each returns a ProviderResult and never raises
# --------------------------------------------------------------------------- #

def _internetdb_ip(ip: str) -> ProviderResult:
    r = ProviderResult("internetdb", link=f"https://internetdb.shodan.io/{ip}")
    resp = _http("GET", f"https://internetdb.shodan.io/{quote(ip)}", timeout=4)
    if resp.status_code == 404:
        r.verdict, r.summary = "clean", "no scan data (no exposed services seen)"
        return r
    resp.raise_for_status()
    d = resp.json()
    ports = d.get("ports") or []
    tags = [str(t).lower() for t in (d.get("tags") or [])]
    vulns = d.get("vulns") or []
    r.detail = {"ports": ports, "tags": tags, "vulns": vulns[:20],
                "hostnames": (d.get("hostnames") or [])[:5], "cpes": (d.get("cpes") or [])[:8]}
    bits = []
    if ports:
        bits.append("ports " + ",".join(str(p) for p in ports[:10]))
    if tags:
        bits.append("tags " + ",".join(tags))
    if vulns:
        bits.append(f"{len(vulns)} known CVE{'s' if len(vulns) != 1 else ''}")
    r.summary = "; ".join(bits) or "no open ports recorded"
    bad = [t for t in tags if t in {"malware", "c2", "compromised", "botnet"}]
    if bad:
        r.verdict = "malicious"
    elif any(t in {"tor", "proxy", "vpn", "scanner", "honeypot"} for t in tags) or vulns:
        r.verdict = "suspicious"
    else:
        r.verdict = "clean"
    return r


def _shodan_ip(ip: str, key: str) -> ProviderResult:
    r = ProviderResult("shodan", link=f"https://www.shodan.io/host/{ip}")
    resp = _http("GET", f"https://api.shodan.io/shodan/host/{quote(ip)}",
                 params={"key": key, "minify": "true"}, timeout=8)
    if resp.status_code == 404:
        r.verdict, r.summary = "clean", "not indexed"
        return r
    resp.raise_for_status()
    d = resp.json()
    tags = [str(t).lower() for t in (d.get("tags") or [])]
    vulns = list(d.get("vulns") or [])
    r.detail = {"org": d.get("org"), "isp": d.get("isp"), "os": d.get("os"),
                "ports": d.get("ports") or [], "tags": tags, "vulns": vulns[:20],
                "last_update": d.get("last_update"), "country": d.get("country_code"),
                "city": d.get("city")}
    bits = [str(d.get("org") or d.get("isp") or "")]
    if d.get("os"):
        bits.append(f"OS {d['os']}")
    if tags:
        bits.append("tags " + ",".join(tags))
    if vulns:
        bits.append(f"{len(vulns)} CVEs")
    r.summary = "; ".join(b for b in bits if b) or "indexed"
    r.verdict = ("malicious" if any(t in {"malware", "c2", "compromised"} for t in tags)
                 else "suspicious" if vulns or any(t in {"tor", "proxy", "vpn"} for t in tags)
                 else "clean")
    return r


def _vt(kind: str, indicator: str, key: str) -> ProviderResult:
    path = {"ip": "ip_addresses", "domain": "domains", "url": "urls", "hash": "files"}[kind]
    ident = _vt_url_id(indicator) if kind == "url" else indicator
    gui = {"ip": "ip-address", "domain": "domain", "url": "url", "hash": "file"}[kind]
    r = ProviderResult("virustotal", link=f"https://www.virustotal.com/gui/{gui}/{ident}")
    resp = _http("GET", f"https://www.virustotal.com/api/v3/{path}/{quote(ident, safe='')}",
                 headers={"x-apikey": key}, timeout=8)
    if resp.status_code == 404:
        r.verdict, r.summary = "unknown", "not in VirusTotal"
        return r
    if resp.status_code == 429:
        r.verdict, r.summary = "error", "rate limited"
        return r
    resp.raise_for_status()
    attrs = ((resp.json() or {}).get("data") or {}).get("attributes") or {}
    stats = attrs.get("last_analysis_stats") or {}
    r.verdict, r.summary = _vt_stats_verdict(stats, kind=kind)
    r.detail = {"stats": stats, "reputation": attrs.get("reputation")}
    if kind == "hash":
        label = ((attrs.get("popular_threat_classification") or {}).get("suggested_threat_label"))
        r.detail["name"] = attrs.get("meaningful_name")
        if label:
            r.detail["threat_label"] = label
            r.summary += f" — {label}"
    if kind == "domain":
        cats = attrs.get("categories") or {}
        if cats:
            r.detail["categories"] = sorted(set(cats.values()))[:4]
    return r


def _greynoise_ip(ip: str, key: str) -> ProviderResult:
    r = ProviderResult("greynoise", link=f"https://viz.greynoise.io/ip/{ip}")
    resp = _http("GET", f"https://api.greynoise.io/v3/community/{quote(ip)}",
                 headers={"key": key, "Accept": "application/json"}, timeout=5)
    if resp.status_code == 404:
        r.verdict, r.summary = "clean", "not seen scanning the internet"
        return r
    if resp.status_code == 429:
        r.verdict, r.summary = "error", "daily limit reached"
        return r
    resp.raise_for_status()
    d = resp.json()
    cls = str(d.get("classification") or "unknown").lower()
    r.detail = {k: d.get(k) for k in ("noise", "riot", "classification", "name", "last_seen")}
    if d.get("riot"):
        r.verdict, r.summary = "clean", f"known business service ({d.get('name') or 'RIOT'})"
    elif cls == "malicious":
        r.verdict, r.summary = "malicious", f"malicious internet scanner ({d.get('name') or 'unknown actor'})"
    elif d.get("noise"):
        r.verdict, r.summary = "suspicious", f"mass-scanning the internet ({cls})"
    else:
        r.verdict, r.summary = "clean", cls
    return r


def _otx(kind: str, indicator: str, key: str) -> ProviderResult:
    section = {"ip": "IPv6" if _is_ipv6(indicator) else "IPv4", "domain": "domain",
               "url": "url", "hash": "file"}[kind]
    r = ProviderResult("otx", link=f"https://otx.alienvault.com/indicator/{section.lower()}/{indicator}")
    resp = _http("GET", f"https://otx.alienvault.com/api/v1/indicators/{section}/{quote(indicator, safe='')}/general",
                 headers={"X-OTX-API-KEY": key}, timeout=6)
    if resp.status_code in (400, 404):
        r.verdict, r.summary = "unknown", "no OTX data"
        return r
    resp.raise_for_status()
    pulses = ((resp.json() or {}).get("pulse_info") or {})
    count = int(pulses.get("count") or 0)
    names = [p.get("name") for p in (pulses.get("pulses") or [])[:3] if p.get("name")]
    r.detail = {"pulses": count, "names": names}
    # OTX is community-curated and noisy (shared infra shows up in pulses);
    # it only ever raises "suspicious". Low counts stay clean — and the summary
    # must not sound like a threat finding when the verdict is clean.
    threshold = 3 if kind != "hash" else 1
    if count >= threshold:
        r.verdict = "suspicious"
        r.summary = (
            f"in {count} threat pulse{'s' if count != 1 else ''}"
            + (f" (e.g. {names[0]})" if names else "")
        )
    elif count > 0:
        r.verdict = "clean"
        r.summary = (
            f"{count} weak community mention{'s' if count != 1 else ''} "
            f"(below threshold; not a listing)"
            + (f" — e.g. {names[0]}" if names else "")
        )
    else:
        r.verdict, r.summary = "clean", "not in any OTX pulse"
    return r


def _urlhaus(kind: str, indicator: str, key: str) -> ProviderResult:
    r = ProviderResult("abusech")
    endpoint, field_name = ("url", "url") if kind == "url" else ("host", "host")
    resp = _http("POST", f"https://urlhaus-api.abuse.ch/v1/{endpoint}/",
                 headers={"Auth-Key": key}, data={field_name: indicator}, timeout=6)
    resp.raise_for_status()
    d = resp.json()
    status = d.get("query_status")
    if status != "ok":
        r.verdict, r.summary = "clean", "URLhaus: not listed"
        return r
    r.link = d.get("urlhaus_reference") or ""
    if kind == "url":
        r.verdict = "malicious"
        r.summary = f"URLhaus: malware URL ({d.get('threat') or 'malware'}, {d.get('url_status') or 'status?'})"
    else:
        online = sum(1 for u in (d.get("urls") or []) if u.get("url_status") == "online")
        count = int(d.get("url_count") or 0)
        r.verdict = "malicious" if online else "suspicious"
        r.summary = f"URLhaus: host served {count} malware URL(s), {online} online now"
    return r


def _threatfox(indicator: str, key: str) -> ProviderResult:
    r = ProviderResult("abusech")
    resp = _http("POST", "https://threatfox-api.abuse.ch/api/v1/",
                 headers={"Auth-Key": key}, json={"query": "search_ioc", "search_term": indicator}, timeout=6)
    resp.raise_for_status()
    d = resp.json()
    if d.get("query_status") != "ok" or not d.get("data"):
        r.verdict, r.summary = "clean", "ThreatFox: not listed"
        return r
    hit = d["data"][0]
    r.verdict = "malicious"
    r.summary = (f"ThreatFox: {hit.get('malware_printable') or 'malware'} "
                 f"{hit.get('threat_type') or 'IOC'} (confidence {hit.get('confidence_level')}%)")
    r.link = f"https://threatfox.abuse.ch/ioc/{hit.get('id')}/" if hit.get("id") else ""
    return r


def _malwarebazaar(sha256: str, key: str) -> ProviderResult:
    r = ProviderResult("abusech", link=f"https://bazaar.abuse.ch/sample/{sha256}/")
    resp = _http("POST", "https://mb-api.abuse.ch/api/v1/",
                 headers={"Auth-Key": key}, data={"query": "get_info", "hash": sha256}, timeout=6)
    resp.raise_for_status()
    d = resp.json()
    if d.get("query_status") != "ok" or not d.get("data"):
        r.verdict, r.summary = "clean", "MalwareBazaar: unknown sample"
        return r
    hit = d["data"][0]
    r.verdict = "malicious"
    r.summary = f"MalwareBazaar: known malware {hit.get('signature') or ''} ({hit.get('file_type') or 'file'})".strip()
    return r


# Spamhaus/SURBL/URIBL return special "refused" codes when queried via big
# public resolvers (8.8.8.8, 1.1.1.1) — those are errors, not listings.
_DNSBL_ZONES = (
    ("dbl.spamhaus.org", {
        "127.0.1.2": ("suspicious", "spam domain"), "127.0.1.4": ("malicious", "phishing domain"),
        "127.0.1.5": ("malicious", "malware domain"), "127.0.1.6": ("malicious", "botnet C&C domain"),
        "127.0.1.102": ("suspicious", "abused legit spam"), "127.0.1.103": ("suspicious", "abused spammed redirector"),
        "127.0.1.104": ("malicious", "abused legit phish"), "127.0.1.105": ("malicious", "abused legit malware"),
        "127.0.1.106": ("malicious", "abused legit botnet C&C"),
    }, {"127.255.255.252", "127.255.255.254", "127.255.255.255"}),
    ("multi.surbl.org", None, {"127.0.0.1"}),
    ("multi.uribl.com", {"127.0.0.2": ("suspicious", "URIBL black")}, {"127.0.0.1", "127.0.0.255"}),
)
_SURBL_BITS = {8: ("malicious", "phishing"), 16: ("malicious", "malware"),
               64: ("suspicious", "abuse/spam"), 128: ("malicious", "cracked site")}

# Permanent test entries each list publishes; used to prove a lookup path works.
_DNSBL_CANARY = {
    "dbl.spamhaus.org": "dbltest.com",
    "zen.spamhaus.org": "2.0.0.127",
    "multi.surbl.org": "test.surbl.org",
    "multi.uribl.com": "test.uribl.com",
    "bl.spamcop.net": "2.0.0.127",
    "dnsbl.sorbs.net": "2.0.0.127",
    "dnsbl-1.uceprotect.net": "2.0.0.127",
}
_NS_CACHE: Dict[str, List[str]] = {}
_PATH_OK: Dict[str, str] = {}  # zone -> "direct" | "system" | "broken"
_DNS_LOCK = threading.Lock()


def _zone_ns_ips(zone: str) -> List[str]:
    import dns.resolver

    with _DNS_LOCK:
        if zone in _NS_CACHE:
            return _NS_CACHE[zone]
    ips: List[str] = []
    try:
        res = dns.resolver.Resolver()
        res.lifetime = 3
        for ns in [a.to_text() for a in res.resolve(zone, "NS")][:4]:
            try:
                ips.extend(a.to_text() for a in res.resolve(ns, "A"))
            except Exception:
                continue
            if len(ips) >= 3:
                break
    except Exception:
        pass
    with _DNS_LOCK:
        _NS_CACHE[zone] = ips
    return ips


def _query_direct(qname: str, zone: str) -> Optional[List[str]]:
    """A-record answers from the zone's authoritative servers; [] = not listed, None = failed."""
    import dns.message
    import dns.query
    import dns.rcode
    import dns.rdatatype

    for ip in _zone_ns_ips(zone)[:3]:
        try:
            resp = dns.query.udp(dns.message.make_query(qname, "A"), ip, timeout=1.5)
        except Exception:
            continue
        if resp.rcode() == dns.rcode.NXDOMAIN:
            return []
        if resp.rcode() != dns.rcode.NOERROR:
            continue
        return [rr.to_text() for rrset in resp.answer if rrset.rdtype == dns.rdatatype.A for rr in rrset]
    return None


def _query_system(qname: str) -> Optional[List[str]]:
    import dns.resolver

    res = dns.resolver.Resolver()
    res.lifetime = 2.5
    try:
        return [a.to_text() for a in res.resolve(qname, "A")]
    except (dns.resolver.NXDOMAIN, dns.resolver.NoAnswer):
        return []
    except Exception:
        return None


def _lookup_path(zone: str) -> str:
    """Pick a lookup path whose canary answers; many ISP/corporate resolvers silently NXDOMAIN blocklists."""
    with _DNS_LOCK:
        if zone in _PATH_OK:
            return _PATH_OK[zone]
    canary = _DNSBL_CANARY.get(zone)
    path = "broken"
    if canary:
        if _query_direct(f"{canary}.{zone}", zone):
            path = "direct"
        elif _query_system(f"{canary}.{zone}"):
            path = "system"
    else:
        path = "direct"
    with _DNS_LOCK:
        _PATH_OK[zone] = path
    return path


def dnsbl_lookup(label: str, zone: str) -> Optional[List[str]]:
    """Query ``label.zone``. Returns answers ([] = not listed) or None when no working path exists.

    ``label`` is a domain, or a reversed IPv4 (``4.3.2.1``) / nibble-reversed IPv6.
    """
    path = _lookup_path(zone)
    if path == "broken":
        return None
    qname = f"{label}.{zone}"
    answers = _query_direct(qname, zone) if path == "direct" else _query_system(qname)
    if answers is None and path == "direct":
        answers = _query_system(qname)
    return answers


def _dnsbl_domain(domain: str) -> ProviderResult:
    r = ProviderResult("dnsbl")
    hits: List[Tuple[str, str, str]] = []
    errors: List[str] = []
    for zone, codes, refused in _DNSBL_ZONES:
        answers = dnsbl_lookup(domain, zone)
        if answers is None:
            errors.append(f"{zone}: no working lookup path")
            continue
        for addr in answers:
            if addr in refused:
                errors.append(f"{zone}: query refused (public resolver?)")
                continue
            if zone == "multi.surbl.org":
                last = int(addr.rsplit(".", 1)[-1])
                for bit, (verdict, why) in _SURBL_BITS.items():
                    if last & bit:
                        hits.append((verdict, zone, why))
            elif codes and addr in codes:
                verdict, why = codes[addr]
                hits.append((verdict, zone, why))
    r.detail = {"hits": hits, "errors": errors}
    if hits:
        r.verdict = "malicious" if any(v == "malicious" for v, _, _ in hits) else "suspicious"
        r.summary = "listed: " + "; ".join(f"{z.split('.')[0].upper()} {why}" for _, z, why in hits)
    elif errors and len(errors) == len(_DNSBL_ZONES):
        r.verdict, r.summary = "error", "blocklist DNS unavailable"
    else:
        r.verdict, r.summary = "clean", "not on DBL / SURBL / URIBL"
    return r


def _gsb_batch(urls: Sequence[str], key: str) -> Dict[str, ProviderResult]:
    body = {
        "client": {"clientId": "aliniant-aes", "clientVersion": "1.0"},
        "threatInfo": {
            "threatTypes": ["MALWARE", "SOCIAL_ENGINEERING", "UNWANTED_SOFTWARE",
                            "POTENTIALLY_HARMFUL_APPLICATION"],
            "platformTypes": ["ANY_PLATFORM"],
            "threatEntryTypes": ["URL"],
            "threatEntries": [{"url": u} for u in urls[:500]],
        },
    }
    resp = _http("POST", "https://safebrowsing.googleapis.com/v4/threatMatches:find",
                 params={"key": key}, json=body, timeout=6)
    resp.raise_for_status()
    matches: Dict[str, List[str]] = {}
    for m in (resp.json() or {}).get("matches") or []:
        u = ((m.get("threat") or {}).get("url")) or ""
        matches.setdefault(u, []).append(str(m.get("threatType") or "THREAT"))
    out: Dict[str, ProviderResult] = {}
    for u in urls:
        if u in matches:
            kinds = sorted(set(matches[u]))
            out[u] = ProviderResult("gsb", "malicious",
                                    "flagged: " + ", ".join(k.replace("_", " ").lower() for k in kinds),
                                    {"threat_types": kinds},
                                    f"https://transparencyreport.google.com/safe-browsing/search?url={quote(u, safe='')}")
        else:
            out[u] = ProviderResult("gsb", "clean", "not flagged")
    return out


def _urlscan_submit(url: str, key: str, visibility: str, deadline: float) -> ProviderResult:
    r = ProviderResult("urlscan")
    resp = _http("POST", "https://urlscan.io/api/v1/scan/",
                 headers={"API-Key": key, "Content-Type": "application/json"},
                 data=json.dumps({"url": url, "visibility": visibility}), timeout=8)
    if resp.status_code in (400, 429):
        r.verdict = "error"
        r.summary = (resp.json() or {}).get("message", f"HTTP {resp.status_code}") if resp.content else f"HTTP {resp.status_code}"
        return r
    resp.raise_for_status()
    uuid = (resp.json() or {}).get("uuid")
    if not uuid:
        r.verdict, r.summary = "error", "no scan id returned"
        return r
    r.link = f"https://urlscan.io/result/{uuid}/"
    time.sleep(10)
    while time.time() < deadline:
        res = _http("GET", f"https://urlscan.io/api/v1/result/{uuid}/", headers={"API-Key": key}, timeout=8)
        if res.status_code == 200:
            verdicts = ((res.json() or {}).get("verdicts") or {}).get("overall") or {}
            score = int(verdicts.get("score") or 0)
            cats = verdicts.get("categories") or []
            brands = verdicts.get("brands") or []
            r.detail = {"score": score, "categories": cats, "brands": brands}
            if verdicts.get("malicious"):
                r.verdict = "malicious"
            elif score > 0:
                r.verdict = "suspicious"
            else:
                r.verdict = "clean"
            bits = [f"score {score}"] + ([", ".join(cats)] if cats else []) + ([f"imitates {', '.join(brands)}"] if brands else [])
            r.summary = "; ".join(bits)
            return r
        time.sleep(4)
    r.verdict, r.summary = "error", "scan still running (open link for result)"
    return r


# --------------------------------------------------------------------------- #
# Orchestration
# --------------------------------------------------------------------------- #

@dataclass
class _Task:
    provider: str
    kind: str
    indicator: str
    fn: Callable[[], ProviderResult]
    quota: bool = False


class ThreatIntel:
    def __init__(self, mode: str = "compact", *, cache: Optional[_Cache] = None,
                 urlscan_visibility: Optional[str] = None):
        self.mode = mode if mode in MODE_BUDGET_S else "compact"
        self.cache = cache or _Cache()
        if urlscan_visibility is None:
            try:
                cfg = json.loads((_state_dir() / "aes_api.json").read_text(encoding="utf-8"))
                urlscan_visibility = str(cfg.get("urlscan_visibility") or "private")
            except Exception:
                urlscan_visibility = "private"
        self.urlscan_visibility = urlscan_visibility if urlscan_visibility in {"private", "unlisted", "public"} else "private"
        self.keys = {p: provider_key(p) for p in PROVIDERS if PROVIDERS[p].get("key")}
        self._vt_left = VT_PER_SCAN[self.mode]
        self._vt_lock = threading.Lock()

    def enabled_providers(self) -> List[str]:
        return [p for p in PROVIDERS if PROVIDERS[p].get("key") is None or self.keys.get(p)]

    def _vt_allowed(self) -> bool:
        with self._vt_lock:
            if self._vt_left <= 0:
                return False
            if not self.cache.take_vt_slot():
                return False
            self._vt_left -= 1
            return True

    def run(
        self,
        *,
        ips: Iterable[Tuple[str, str]] = (),
        domains: Iterable[Tuple[str, str]] = (),
        urls: Iterable[Tuple[str, str]] = (),
        hashes: Iterable[Tuple[str, str]] = (),
        deep_urls: Iterable[str] = (),
    ) -> List[IndicatorReport]:
        """Each iterable yields ``(indicator, role)`` in priority order."""
        start = time.time()
        deadline = start + MODE_BUDGET_S[self.mode]
        reports: Dict[Tuple[str, str], IndicatorReport] = {}
        tasks: List[Tuple[IndicatorReport, _Task]] = []

        def report(kind: str, ind: str, role: str) -> IndicatorReport:
            rep = reports.get((kind, ind))
            if rep is None:
                rep = reports[(kind, ind)] = IndicatorReport(ind, kind, role)
            return rep

        k = self.keys
        heavy = self.mode in {"full", "deep"}

        for ind, role in hashes:
            rep = report("hash", ind.lower(), role)
            if k.get("virustotal"):
                tasks.append((rep, _Task("virustotal", "hash", ind, lambda i=ind: _vt("hash", i, k["virustotal"]), True)))
            if k.get("abusech"):
                tasks.append((rep, _Task("abusech", "hash", ind, lambda i=ind: _malwarebazaar(i, k["abusech"]))))
            if k.get("otx") and heavy:
                tasks.append((rep, _Task("otx", "hash", ind, lambda i=ind: _otx("hash", i, k["otx"]))))

        for n, (ind, role) in enumerate(ips):
            rep = report("ip", ind, role)
            tasks.append((rep, _Task("internetdb", "ip", ind, lambda i=ind: _internetdb_ip(i))))
            primary = role == "origin-ip" or n == 0
            if k.get("greynoise"):
                tasks.append((rep, _Task("greynoise", "ip", ind, lambda i=ind: _greynoise_ip(i, k["greynoise"]))))
            if k.get("abusech"):
                tasks.append((rep, _Task("abusech", "ip", ind, lambda i=ind: _threatfox(i, k["abusech"]))))
            if k.get("shodan") and heavy and n < 3:
                tasks.append((rep, _Task("shodan", "ip", ind, lambda i=ind: _shodan_ip(i, k["shodan"]))))
            if k.get("otx") and (heavy or primary):
                tasks.append((rep, _Task("otx", "ip", ind, lambda i=ind: _otx("ip", i, k["otx"]))))
            if k.get("virustotal") and primary:
                tasks.append((rep, _Task("virustotal", "ip", ind, lambda i=ind: _vt("ip", i, k["virustotal"]), True)))

        for n, (ind, role) in enumerate(domains):
            rep = report("domain", ind, role)
            tasks.append((rep, _Task("dnsbl", "domain", ind, lambda i=ind: _dnsbl_domain(i))))
            skip_quota = ind in _QUOTA_SKIP_DOMAINS
            if k.get("abusech") and not skip_quota:
                tasks.append((rep, _Task("abusech", "domain", ind, lambda i=ind: _urlhaus("domain", i, k["abusech"]))))
            if k.get("otx") and heavy and not skip_quota and n < 5:
                tasks.append((rep, _Task("otx", "domain", ind, lambda i=ind: _otx("domain", i, k["otx"]))))
            if k.get("virustotal") and not skip_quota and (role == "sender-domain" or heavy):
                tasks.append((rep, _Task("virustotal", "domain", ind, lambda i=ind: _vt("domain", i, k["virustotal"]), True)))

        url_list = []
        for n, (ind, role) in enumerate(urls):
            rep = report("url", ind, role)
            url_list.append(ind)
            host_dom = registrable_domain(urlparse(ind).hostname or "")
            skip_quota = host_dom in _QUOTA_SKIP_DOMAINS
            if k.get("abusech") and n < 10:
                tasks.append((rep, _Task("abusech", "url", ind, lambda i=ind: _urlhaus("url", i, k["abusech"]))))
            if k.get("virustotal") and heavy and not skip_quota and n < 3:
                tasks.append((rep, _Task("virustotal", "url", ind, lambda i=ind: _vt("url", i, k["virustotal"]), True)))

        if self.mode == "deep" and k.get("urlscan"):
            for ind in list(deep_urls)[:2]:
                rep = report("url", ind, "link-url")
                tasks.append((rep, _Task("urlscan", "url", ind,
                                         lambda i=ind: _urlscan_submit(i, k["urlscan"], self.urlscan_visibility, deadline - 2))))

        # Google Safe Browsing: one batched call for every URL.
        if k.get("gsb") and url_list:
            pending = []
            for u in url_list:
                cached = self.cache.get(f"gsb|url|{u}")
                if cached:
                    reports[("url", u)].results.append(cached)
                else:
                    pending.append(u)
            if pending:
                try:
                    for u, res in _gsb_batch(pending, k["gsb"]).items():
                        reports[("url", u)].results.append(res)
                        self.cache.put(f"gsb|url|{u}", res)
                except Exception as exc:
                    logger.warning("Safe Browsing lookup failed: %s", type(exc).__name__)
                    for u in pending:
                        reports[("url", u)].results.append(ProviderResult("gsb", "error", "lookup failed"))

        self._execute(tasks, deadline)
        self.cache.save()
        return list(reports.values())

    def _execute(self, tasks: List[Tuple[IndicatorReport, _Task]], deadline: float) -> None:
        runnable: List[Tuple[IndicatorReport, _Task]] = []
        for rep, task in tasks:
            key = f"{task.provider}|{task.kind}|{task.indicator}"
            cached = self.cache.get(key)
            if cached is not None:
                rep.results.append(cached)
            else:
                runnable.append((rep, task))
        if not runnable:
            return

        def call(task: _Task) -> ProviderResult:
            if task.quota and not self._vt_allowed():
                return ProviderResult(task.provider, "error", "skipped (free-tier rate limit)")
            try:
                res = task.fn()
            except Exception as exc:
                res = ProviderResult(task.provider, "error", f"lookup failed ({type(exc).__name__})")
            return res

        pool = ThreadPoolExecutor(max_workers=8, thread_name_prefix="aes-ti")
        futures = {pool.submit(call, task): (rep, task) for rep, task in runnable}
        done, not_done = futures_wait(futures, timeout=max(0.5, deadline - time.time()))
        for fut in done:
            rep, task = futures[fut]
            res = fut.result()
            rep.results.append(res)
            if not (res.verdict == "error" and "rate limit" in res.summary):
                self.cache.put(f"{task.provider}|{task.kind}|{task.indicator}", res)
        for fut in not_done:
            rep, task = futures[fut]
            rep.results.append(ProviderResult(task.provider, "error", "timed out"))
        pool.shutdown(wait=False, cancel_futures=True)


def probe_provider(provider: str, key: Optional[str] = None) -> Tuple[bool, str]:
    """Check a key against a harmless endpoint. Returns (ok, message)."""
    key = (key if key is not None else provider_key(provider)).strip()
    if PROVIDERS.get(provider, {}).get("key") and not key:
        return False, "no key"
    try:
        if provider == "shodan":
            resp = _http("GET", "https://api.shodan.io/api-info", params={"key": key}, timeout=8)
            if resp.status_code == 200:
                d = resp.json()
                return True, f"OK — plan {d.get('plan', '?')}, {d.get('query_credits', '?')} query credits"
        elif provider == "virustotal":
            resp = _http("GET", "https://www.virustotal.com/api/v3/ip_addresses/8.8.8.8",
                         headers={"x-apikey": key}, timeout=8)
            if resp.status_code == 200:
                return True, "OK"
        elif provider == "gsb":
            resp = _http("POST", "https://safebrowsing.googleapis.com/v4/threatMatches:find",
                         params={"key": key}, timeout=8, json={
                             "client": {"clientId": "aliniant-aes", "clientVersion": "1.0"},
                             "threatInfo": {"threatTypes": ["MALWARE"], "platformTypes": ["ANY_PLATFORM"],
                                            "threatEntryTypes": ["URL"],
                                            "threatEntries": [{"url": "https://example.com/"}]}})
            if resp.status_code == 200:
                return True, "OK"
        elif provider == "greynoise":
            resp = _http("GET", "https://api.greynoise.io/v3/community/8.8.8.8",
                         headers={"key": key, "Accept": "application/json"}, timeout=8)
            if resp.status_code in (200, 404):
                return True, "OK"
        elif provider == "otx":
            resp = _http("GET", "https://otx.alienvault.com/api/v1/users/me",
                         headers={"X-OTX-API-KEY": key}, timeout=8)
            if resp.status_code == 200:
                return True, f"OK — {resp.json().get('username', 'account')}"
        elif provider == "abusech":
            resp = _http("POST", "https://mb-api.abuse.ch/api/v1/", headers={"Auth-Key": key},
                         data={"query": "get_info", "hash": "0" * 64}, timeout=8)
            if resp.status_code == 200 and resp.json().get("query_status") in {"hash_not_found", "ok", "illegal_hash"}:
                return True, "OK"
        elif provider == "urlscan":
            resp = _http("GET", "https://urlscan.io/user/quotas/", headers={"API-Key": key}, timeout=8)
            if resp.status_code == 200:
                return True, "OK"
        elif provider == "internetdb":
            resp = _http("GET", "https://internetdb.shodan.io/8.8.8.8", timeout=6)
            return resp.status_code in (200, 404), f"HTTP {resp.status_code}"
        elif provider == "dnsbl":
            paths = {z: _lookup_path(z) for z, _, _ in _DNSBL_ZONES}
            ok = any(p != "broken" for p in paths.values())
            return ok, ", ".join(f"{z.split('.')[0].upper()}: {p}" for z, p in paths.items())
        else:
            return False, "unknown provider"
        if resp.status_code in (401, 403):
            return False, "key rejected"
        if resp.status_code == 429:
            return True, "key accepted (rate limited right now)"
        return False, f"HTTP {resp.status_code}"
    except Exception as exc:
        return False, f"unreachable ({type(exc).__name__})"


# --------------------------------------------------------------------------- #
# Scoring
# --------------------------------------------------------------------------- #

MAX_TI_POINTS = 60


def score_reports(reports: Sequence[IndicatorReport]) -> List[Dict[str, Any]]:
    """Turn indicator reports into ``[{label, points, note}]`` risk findings."""
    findings: List[Dict[str, Any]] = []

    def add(label: str, pts: int, note: str) -> None:
        findings.append({"label": label, "points": pts, "note": note})

    for rep in reports:
        mal = [r for r in rep.results if r.verdict == "malicious"]
        sus = [r for r in rep.results if r.verdict == "suspicious"]
        names = lambda rs: ", ".join(sorted({PROVIDERS.get(r.provider, {}).get("label", r.provider).split(" (")[0] for r in rs}))
        if rep.kind == "hash":
            if mal:
                add(f"Attachment is known malware ({names(mal)})", 50, rep.lines()[0] if rep.lines() else "")
            elif sus:
                add(f"Attachment hash flagged ({names(sus)})", 10, "")
        elif rep.kind == "url":
            if mal:
                add(f"Link flagged malicious: {urlparse(rep.indicator).hostname} ({names(mal)})", 25,
                    "; ".join(r.summary for r in mal)[:160])
            elif sus:
                add(f"Link flagged suspicious: {urlparse(rep.indicator).hostname} ({names(sus)})", 5, "")
        elif rep.kind == "domain":
            role = "Sender domain" if rep.role == "sender-domain" else "Linked domain" if rep.role == "link-domain" else "Domain"
            if mal:
                add(f"{role} {rep.indicator} on threat lists ({names(mal)})", 20,
                    "; ".join(r.summary for r in mal)[:160])
            elif sus:
                add(f"{role} {rep.indicator} flagged ({names(sus)})", 5, "; ".join(r.summary for r in sus)[:160])
        elif rep.kind == "ip":
            where = "Origin IP" if rep.role == "origin-ip" else "Relay IP"
            pts = 0
            reasons: List[str] = []
            for r in rep.results:
                if r.provider in {"internetdb", "shodan"}:
                    tags = [t for t in (r.detail.get("tags") or []) if t in _BAD_SHODAN_TAGS]
                    if tags:
                        tag_pts = max(_BAD_SHODAN_TAGS[t] for t in tags)
                        if rep.role != "origin-ip":
                            tag_pts = min(tag_pts, 8)
                        pts = max(pts, tag_pts)
                        reasons.append(f"Shodan tags {','.join(tags)}")
                elif r.verdict == "malicious":
                    pts = max(pts, 15 if rep.role == "origin-ip" else 8)
                    reasons.append(f"{PROVIDERS[r.provider]['label']}: {r.summary}")
                elif r.verdict == "suspicious" and r.provider in {"greynoise", "virustotal"}:
                    pts = max(pts, 5 if rep.role == "origin-ip" else 3)
                    reasons.append(f"{PROVIDERS[r.provider]['label']}: {r.summary}")
            if pts:
                add(f"{where} {rep.indicator} flagged by threat intel", pts, "; ".join(reasons)[:180])

    total = 0
    out = []
    for f in sorted(findings, key=lambda x: -x["points"]):
        pts = max(0, min(f["points"], MAX_TI_POINTS - total))
        total += pts
        out.append({**f, "points": pts})
    return out


def collect_header_indicators(headers: str) -> Tuple[List[str], List[str]]:
    """URLs and domains present in the headers (compact scans have no body)."""
    urls: List[str] = []
    domains: List[str] = []
    if not headers:
        return urls, domains
    for m in re.finditer(r"https?://[^\s<>\"',;]+", headers):
        u = m.group(0).rstrip(">).")
        if u not in urls:
            urls.append(u)
    patterns = (
        r"^(?:From|Reply-To|Return-Path|Sender):[^\n]*?@([A-Za-z0-9.-]+\.[A-Za-z]{2,})",
        r"\bd=([A-Za-z0-9.-]+\.[A-Za-z]{2,})\s*;",
        r"^Message-ID:\s*<[^@>]+@([A-Za-z0-9.-]+\.[A-Za-z]{2,})>",
    )
    for pat in patterns:
        for m in re.finditer(pat, headers, re.IGNORECASE | re.MULTILINE):
            d = registrable_domain(m.group(1))
            if d and d not in domains:
                domains.append(d)
    for u in urls:
        d = registrable_domain(urlparse(u).hostname or "")
        if d and d not in domains:
            domains.append(d)
    return urls[:20], domains[:10]
