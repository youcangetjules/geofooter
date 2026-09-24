#!/usr/bin/env python3
"""
Email Security Analysis Tool - Fixed Version
Analyzes email headers for security indicators, geolocation, and authentication.
"""

import sys
import os
import re
import json
import time
import socket
import logging
import ipaddress
import math
from pathlib import Path
from typing import Optional, Dict, List, Tuple, Any, Union, Set
from dataclasses import dataclass, field
from datetime import datetime, timedelta
from logging.handlers import RotatingFileHandler

import urllib.request
import urllib.error
import urllib.parse
import threading
import subprocess

# Windows: whois.exe / powershell / MpCmdRun spawn console flashes unless
# CREATE_NO_WINDOW is set. Patch Popen before third-party imports (python-whois).
if sys.platform == "win32":
    _CREATE_NO_WINDOW = getattr(subprocess, "CREATE_NO_WINDOW", 0x08000000)
    _OrigPopen = subprocess.Popen

    class _SilentPopen(_OrigPopen):  # type: ignore[misc, valid-type]
        def __init__(self, *args, **kwargs):
            try:
                kwargs["creationflags"] = int(kwargs.get("creationflags", 0) or 0) | _CREATE_NO_WINDOW
            except Exception:
                pass
            super().__init__(*args, **kwargs)

    subprocess.Popen = _SilentPopen  # type: ignore[misc]

# VBA runs this file by path, which puts aes\ first on sys.path; suite packages
# (aes, guri, aura, geofooter) resolve from the install root instead.
_GEOFOOTER_ROOT = Path(__file__).resolve().parents[1]
_PKG_DIR = Path(__file__).resolve().parent
sys.path[:] = [p for p in sys.path if Path(p or ".").resolve() != _PKG_DIR]
if str(_GEOFOOTER_ROOT) not in sys.path:
    sys.path.insert(0, str(_GEOFOOTER_ROOT))

try:
    import aes.addr_decode as addr_decode
except ImportError:
    addr_decode = None  # type: ignore[assignment]

try:
    import aes.time_checks as time_checks
except ImportError:
    time_checks = None  # type: ignore[assignment]

try:
    import aes.route_checks as route_checks
except ImportError:
    route_checks = None  # type: ignore[assignment]

try:
    import aes.threat_intel as threat_intel
except ImportError:
    threat_intel = None  # type: ignore[assignment]

try:
    from aura.detect import match_sender as broker_match_sender
except ImportError:
    broker_match_sender = None  # type: ignore[assignment]

try:
    from aura.pending import enqueue_from_aes as broker_enqueue
except ImportError:
    broker_enqueue = None  # type: ignore[assignment]

# Test imports and provide helpful error messages
try:
    import dns.resolver
except ImportError as e:
    print(f"ERROR: Missing required module 'dnspython'. Install with: pip install dnspython", file=sys.stderr)
    print(f"Import error: {e}", file=sys.stderr)
    sys.exit(1)

try:
    import whois
except ImportError as e:
    print(f"ERROR: Missing required module 'python-whois'. Install with: pip install python-whois", file=sys.stderr)
    print(f"Import error: {e}", file=sys.stderr)
    sys.exit(1)

try:
    import requests
except ImportError as e:
    print(f"ERROR: Missing required module 'requests'. Install with: pip install requests", file=sys.stderr)
    print(f"Import error: {e}", file=sys.stderr)
    sys.exit(1)

try:
    import pytz
except ImportError as e:
    print(f"ERROR: Missing required module 'pytz'. Install with: pip install pytz", file=sys.stderr)
    print(f"Import error: {e}", file=sys.stderr)
    sys.exit(1)

import html
import random
import string
import email.utils
import sqlite3
import hashlib

try:
    import pycountry
except ImportError as e:
    print(f"ERROR: Missing required module 'pycountry'. Install with: pip install pycountry", file=sys.stderr)
    print(f"Import error: {e}", file=sys.stderr)
    sys.exit(1)

try:
    from ipwhois import IPWhois, exceptions as ipwhois_exceptions
except ImportError as e:
    print(f"ERROR: Missing required module 'ipwhois'. Install with: pip install ipwhois", file=sys.stderr)
    print(f"Import error: {e}", file=sys.stderr)
    sys.exit(1)

# Import GURI module
try:
    from guri.core import GURIDatabase, connect_guri_database
except ImportError as e:
    print(f"ERROR: Missing required module 'guri.core'. Make sure the guri\\ package is under the install root.", file=sys.stderr)
    print(f"Import error: {e}", file=sys.stderr)
    sys.exit(1)

try:
    from aes.scanners.attachment import AttachmentScanner, AttachmentVerdict
except ImportError as e:
    print(f"ERROR: Missing required module 'aes.scanners.attachment'. Make sure the aes\\ package is under the install root.", file=sys.stderr)
    print(f"Import error: {e}", file=sys.stderr)
    sys.exit(1)

try:
    from aes.scanners.link import LinkScanner, LinkFinding
except ImportError:
    LinkScanner = None  # type: ignore
    LinkFinding = None  # type: ignore

try:
    from aes.scanners.body import BodyScanner, BodyScanResult
except ImportError:
    BodyScanner = None  # type: ignore
    BodyScanResult = None  # type: ignore

try:
    from aes.secret_store import (
        abuseipdb_check_url,
        aes_http_get,
        get_abuseipdb_api_key,
        migrate_plaintext_if_present,
    )
except ImportError:
    def get_abuseipdb_api_key() -> str:
        return (os.environ.get("AES_ABUSEIPDB_API_KEY") or "").strip()

    def migrate_plaintext_if_present(name: str = "abuseipdb") -> str:
        return "missing"

    def abuseipdb_check_url() -> str:
        return "https://api.abuseipdb.com/api/v2/check"

    def aes_http_get(
        url: str,
        *,
        headers: Optional[Dict[str, str]] = None,
        params: Optional[Dict[str, Any]] = None,
        timeout: float = 10,
    ) -> tuple[Any, bool]:
        import requests

        resp = requests.get(
            url, headers=headers, params=params, timeout=timeout, verify=False
        )
        return resp, True

# Set up logging at the very first step (debug → debuglog\, crashes → crashlogs\)
try:
    from geofooter.crashlog import install_sys_excepthook, operational_log_path

    install_sys_excepthook(prefix="geolocate")
    log_file = str(operational_log_path("geolocate_debug.log"))
except Exception:
    try:
        from geofooter.paths import debuglog_dir

        log_file = str(debuglog_dir() / "geolocate_debug.log")
    except Exception:
        log_file = str(Path(__file__).resolve().parent.parent / "debuglog" / "geolocate_debug.log")
    try:
        os.makedirs(os.path.dirname(log_file), exist_ok=True)
    except Exception:
        pass
logging.basicConfig(filename=log_file, level=logging.DEBUG,
                    format='%(asctime)s - %(levelname)s - %(message)s')
logging.debug("Script started with arguments: %s", sys.argv)

# Configuration
try:
    from geofooter.paths import get_install_root as _gf_root

    _BASE = str(_gf_root()).replace("\\", "/")
except Exception:
    _BASE = ""

CONFIG = {
    "base_path": _BASE,
    "log_file": "debuglog/geolocate_debug.log",
    "log_max_bytes": 5 * 1024 * 1024,
    "log_backup_count": 5,
    "timezone": "Europe/London",
    "api_keys": {
        # AbuseIPDB is loaded from DPAPI-secured settings (see aes/secret_store.py).
        # Never store the raw key here. Optional override: AES_ABUSEIPDB_API_KEY.
        "abuseipdb": "",
        "ipinfo": "1a9a711593217b",
        "ipgeolocation": "82dad7d806b54c158da274d5576e7702"
    },
    "blocklists": {
        "ipv4": ["dnsbl.sorbs.net", "zen.spamhaus.org", "bl.spamcop.net", "dnsbl-1.uceprotect.net"],
        "ipv6": ["zen.spamhaus.org"]
    },
    # Mail risk bands (upper-inclusive until High+mitigate):
    #   0–25 Low, 26–50 Raised, 51–70 High, >70 High + strip/text-only.
    "risk_thresholds": {
        "low_max": 25,
        "raised_max": 50,
        "high_max": 70,
    },
    "domain_age_thresholds": {
        "extremely_new": 10,
        "very_new": 30,
        "new": 180
    },
    "footer_compact_default": True
}


def _suite_root_path() -> Path:
    base = (CONFIG.get("base_path") or "").strip()
    if not base:
        try:
            from geofooter.paths import get_install_root
            base = str(get_install_root())
        except Exception:
            base = str(Path(__file__).resolve().parent.parent)
    return Path(base)


def _suite_output(*parts: str) -> Path:
    p = _suite_root_path().joinpath("output", *parts)
    p.parent.mkdir(parents=True, exist_ok=True)
    return p


def _suite_file_url(*parts: str) -> str:
    """file:// URL for a path under <install_root>/output/…"""
    return _suite_output(*parts).resolve().as_uri()



def _resolve_abuseipdb_api_key() -> str:
    """DPAPI/env AbuseIPDB key (never from plaintext source)."""
    try:
        migrate_plaintext_if_present("abuseipdb")
    except Exception:
        pass
    key = (get_abuseipdb_api_key() or "").strip()
    if key:
        return key
    # Legacy CONFIG slot — left empty in source; kept only for rare overrides.
    return str((CONFIG.get("api_keys") or {}).get("abuseipdb") or "").strip()


# --- External IP API helpers (retries, rate limits, DNSBL) -------------------

_ABUSEIPDB_LOCK = threading.Lock()
_ABUSEIPDB_LAST_CALL = 0.0
_ABUSEIPDB_MIN_INTERVAL = 0.4  # ~2.5 req/s — avoids deep-scan timeouts


def _rate_limit_abuseipdb() -> None:
    """Serialize AbuseIPDB calls so deep scans do not stampede the API."""
    global _ABUSEIPDB_LAST_CALL
    with _ABUSEIPDB_LOCK:
        now = time.time()
        wait = _ABUSEIPDB_MIN_INTERVAL - (now - _ABUSEIPDB_LAST_CALL)
        if wait > 0:
            time.sleep(wait)
        _ABUSEIPDB_LAST_CALL = time.time()


def _http_get_with_retry(
    url: str,
    *,
    headers: Optional[Dict[str, str]] = None,
    params: Optional[Dict[str, Any]] = None,
    timeout: float = 12,
    retries: int = 3,
    backoff: float = 1.5,
) -> Tuple[Any, bool]:
    """GET with retries on timeout/connection errors; uses aes_http_get TLS fallback."""
    last_exc: Optional[Exception] = None
    for attempt in range(max(1, retries)):
        try:
            resp, insecure = aes_http_get(
                url,
                headers=headers,
                params=params,
                timeout=timeout,
            )
            return resp, insecure
        except Exception as exc:
            last_exc = exc
            if attempt + 1 < retries:
                time.sleep(backoff * (attempt + 1))
    if last_exc is not None:
        raise last_exc
    raise RuntimeError("HTTP GET failed with no exception")


def _ipv6_dnsbl_label(ip: str) -> str:
    """Nibble-reversed expanded IPv6 for DNSBL lookups."""
    addr = ipaddress.ip_address(ip)
    if addr.version != 6:
        return ""
    hex_chars = addr.exploded.replace(":", "").lower()
    return ".".join(reversed(hex_chars))


def _blocklist_dns_query(ip: str, bl_zone: str) -> str:
    """Build the DNS query name for an IPv4 or IPv6 blocklist zone."""
    addr = ipaddress.ip_address(ip)
    if addr.version == 4:
        rev = ".".join(reversed(str(ip).split(".")))
        return f"{rev}.{bl_zone}"
    nibbles = _ipv6_dnsbl_label(ip)
    if not nibbles:
        return ""
    # Spamhaus and most major DNSBLs use the ip6.* subdomain for IPv6.
    if "spamhaus" in bl_zone or bl_zone.startswith("zen."):
        return f"{nibbles}.ip6.{bl_zone}"
    return f"{nibbles}.{bl_zone}"


def _proxycheck_entry(payload: Dict[str, Any], ip: str) -> Optional[Dict[str, Any]]:
    """Extract the per-IP object from a proxycheck.io v2 response."""
    if not isinstance(payload, dict):
        return None
    for key in (ip, ip.lower(), ip.upper()):
        val = payload.get(key)
        if isinstance(val, dict):
            return val
    for key, val in payload.items():
        if key == "status" or not isinstance(val, dict):
            continue
        return val
    return None


def _apply_proxycheck(result: "GeoLocationResult", data: Dict[str, Any]) -> None:
    """Merge proxycheck.io privacy signals (free tier, no API key required)."""
    ptype = str(data.get("type") or "").strip()
    ptype_u = ptype.upper()
    ptype_l = ptype.lower()
    proxy_raw = str(data.get("proxy") or "").lower()
    vpn_raw = str(data.get("vpn") or "").lower()
    hosting_already = bool(result.security_data.get("Hosting"))

    if ptype_u == "TOR":
        _set_security_flag(result, "Tor", True)

    if vpn_raw == "yes" or (ptype_u == "VPN" and proxy_raw == "yes"):
        if not hosting_already:
            _set_security_flag(result, "VPN", True)
            provider = str(data.get("provider") or "").strip()
            if provider:
                result.vpn_service = result.vpn_service or provider
    elif vpn_raw == "no":
        _set_security_flag(result, "VPN", False)

    if proxy_raw == "yes":
        # Enterprise mail/CDN egress is often tagged proxy=yes; Hosting wins.
        _set_security_flag(result, "Proxy", False if hosting_already else True)
        if hosting_already and result.security_data.get("Anonymous") is None:
            _set_security_flag(result, "Anonymous", False)
    elif proxy_raw == "no":
        _set_security_flag(result, "Proxy", False)

    if ptype_l in {"open proxy", "public proxy", "anonymous"}:
        _set_security_flag(result, "Anonymous", True)
    elif proxy_raw == "no" and result.security_data.get("Anonymous") is None:
        _set_security_flag(result, "Anonymous", False)

    if ptype_l in {"hosting", "datacenter"}:
        _set_security_flag(result, "Hosting", True)

    if not result.org or result.org == "Unknown":
        org = str(data.get("provider") or data.get("organisation") or "").strip()
        if org:
            result.org = org
    if not result.timezone and data.get("timezone"):
        result.timezone = str(data["timezone"])
    if not result.asn or result.asn == "Unknown":
        asn = str(data.get("asn") or "").strip()
        if asn:
            result.asn = asn
    if (not result.city or result.city == "Unknown") and data.get("city"):
        result.city = str(data.get("city"))
    if not result.country and data.get("isocode"):
        result.country = str(data.get("isocode"))
    if not result.success and (result.city or result.org):
        result.success = True


# Default route / ASN risk settings (overridden by aes_risk_route.json).
DEFAULT_ROUTE_RISK_CONFIG: Dict[str, Any] = {
    "provider": "abuseipdb",
    "abuse_threshold": 25,
    "extra_high_risk_asns": [],
    "enabled": True,
}

_ROUTE_RISK_CONFIG_CACHE: Optional[Dict[str, Any]] = None


def _normalize_asn(value: Any) -> str:
    """Normalize ASN to canonical AS12345 form (empty if unknown)."""
    if value is None:
        return ""
    text = str(value).strip().upper()
    if not text or text in {"UNKNOWN", "N/A", "—", "-"}:
        return ""
    match = re.search(r"AS\s*(\d+)", text)
    if match:
        return f"AS{match.group(1)}"
    if text.isdigit():
        return f"AS{text}"
    return ""


def load_route_risk_config(force_reload: bool = False) -> Dict[str, Any]:
    """Load route/ASN risk config from LocalAppData or GeoFooter; cache result."""
    global _ROUTE_RISK_CONFIG_CACHE
    if _ROUTE_RISK_CONFIG_CACHE is not None and not force_reload:
        return dict(_ROUTE_RISK_CONFIG_CACHE)

    cfg = dict(DEFAULT_ROUTE_RISK_CONFIG)
    cfg["extra_high_risk_asns"] = list(DEFAULT_ROUTE_RISK_CONFIG["extra_high_risk_asns"])

    candidates = [
        Path(os.environ.get("LOCALAPPDATA", "")) / "GeoFooter" / "aes_risk_route.json",
        Path(CONFIG["base_path"]) / "aes_risk_route.json",
    ]
    for path in candidates:
        try:
            if not path.is_file():
                continue
            data = json.loads(path.read_text(encoding="utf-8"))
            if not isinstance(data, dict):
                continue
            provider = str(data.get("provider") or cfg["provider"]).strip().lower()
            if provider:
                cfg["provider"] = provider
            if "enabled" in data:
                cfg["enabled"] = bool(data.get("enabled"))
            try:
                threshold = int(data.get("abuse_threshold", cfg["abuse_threshold"]))
                cfg["abuse_threshold"] = max(0, min(100, threshold))
            except (TypeError, ValueError):
                pass
            extra = data.get("extra_high_risk_asns") or []
            if isinstance(extra, str):
                extra = re.split(r"[\s,;]+", extra)
            normalized = []
            for item in extra:
                asn = _normalize_asn(item)
                if asn and asn not in normalized:
                    normalized.append(asn)
            cfg["extra_high_risk_asns"] = normalized
            break
        except Exception as exc:
            logging.getLogger(__name__).warning("Failed to load route risk config %s: %s", path, exc)

    _ROUTE_RISK_CONFIG_CACHE = dict(cfg)
    cfg["extra_high_risk_asns"] = list(cfg["extra_high_risk_asns"])
    return cfg


# Sender status settings (overridden by aes_sender_status.json).
DEFAULT_SENDER_STATUS_CONFIG: Dict[str, Any] = {
    # Emails from a sender before they count as Known rather than Unknown.
    "known_threshold": 5,
}

# A sender the user has explicitly trusted keeps this share of their score.
TRUSTED_SENDER_DISCOUNT = 0.33
# Link scoring (Basic scan and Deep Scan fusion).
SUSPICIOUS_LINK_POINTS = 2   # medium-risk / suspicious links
HIGH_RISK_LINK_POINTS = 10   # high-risk links
# Beacons and bad attachments still use a flat per-item charge.
HIGH_RISK_ELEMENT_POINTS = 2

RISK_LEVEL_COLORS: Dict[str, str] = {
    "LOW": "#4CAF50",
    "RAISED": "#FFC107",
    "HIGH": "#F44336",
}


def risk_level_for_score(score: int, domain_age_flag: Optional[str] = None) -> str:
    """Map a 0–100 score onto LOW / RAISED / HIGH."""
    if domain_age_flag:
        return "HIGH"
    thresholds = CONFIG["risk_thresholds"]
    s = max(0, min(100, int(score)))
    if s <= int(thresholds.get("low_max", 25)):
        return "LOW"
    if s <= int(thresholds.get("raised_max", 50)):
        return "RAISED"
    return "HIGH"


def risk_requires_mitigation(score: int) -> bool:
    """Scores above high_max trigger attachment strip + text-only conversion."""
    thresholds = CONFIG["risk_thresholds"]
    return int(score) > int(thresholds.get("high_max", 70))


def risk_color_for_level(level: str) -> str:
    return RISK_LEVEL_COLORS.get(str(level or "").upper(), "#888888")


# Footer scan strip: one dark slate background, white type, muted dividers.
AES_STRIP_BG = "#0f1f2a"
AES_STRIP_TEXT = "#ffffff"
AES_STRIP_MUTED = "#9fb3bf"
AES_STRIP_RULE = "#24394a"
AES_STRIP_DIVIDER = "#4d6575"
# Quick Action chips on that strip. Default is a white pill with slate text.
AES_CHIP_BG = "#ffffff"
AES_CHIP_FG = "#0f1f2a"
AES_CHIP_BLOCKED_BG = "#d92d20"
AES_CHIP_TRUSTED_BG = "#1b7a3d"


def strip_separators(inner_html: str) -> str:
    """Turn ' | ' text separators into muted hairline dividers."""
    divider = (
        f"<span style='color:{AES_STRIP_DIVIDER}; padding:0 8px; font-weight:normal;'>|</span>"
    )
    return (inner_html or "").replace(" | ", divider)


def contrasting_text_color(background: str) -> str:
    """Near-black or white, whichever meets contrast on this banner colour.

    The scan strip runs from medium green to red. White on that green is
    about 2.8:1, so a fixed white label disappears.
    """
    text = str(background or "").strip().lower()
    rgb = None
    if text.startswith("rgb"):
        nums = re.findall(r"\d+", text)
        if len(nums) >= 3:
            rgb = tuple(int(n) for n in nums[:3])
    elif text.startswith("#") and len(text) >= 7:
        try:
            rgb = (int(text[1:3], 16), int(text[3:5], 16), int(text[5:7], 16))
        except ValueError:
            rgb = None
    if rgb is None:
        return "#1a1a1a"

    def channel(value: int) -> float:
        s = value / 255.0
        if s <= 0.04045:
            return s / 12.92
        return ((s + 0.055) / 1.055) ** 2.4

    r, g, b = (channel(c) for c in rgb)
    luminance = 0.2126 * r + 0.7152 * g + 0.0722 * b
    return "#1a1a1a" if luminance > 0.179 else "#ffffff"

_SENDER_STATUS_CONFIG_CACHE: Optional[Dict[str, Any]] = None


def _geofooter_data_dir() -> Path:
    return Path(os.environ.get("LOCALAPPDATA", CONFIG["base_path"])) / "GeoFooter"


def load_sender_status_config(force_reload: bool = False) -> Dict[str, Any]:
    """Load sender status settings written by the AES settings dialog."""
    global _SENDER_STATUS_CONFIG_CACHE
    if _SENDER_STATUS_CONFIG_CACHE is not None and not force_reload:
        return dict(_SENDER_STATUS_CONFIG_CACHE)

    cfg = dict(DEFAULT_SENDER_STATUS_CONFIG)
    candidates = [
        _geofooter_data_dir() / "aes_sender_status.json",
        Path(CONFIG["base_path"]) / "aes_sender_status.json",
    ]
    for path in candidates:
        try:
            if not path.is_file():
                continue
            data = json.loads(path.read_text(encoding="utf-8"))
            if not isinstance(data, dict):
                continue
            try:
                cfg["known_threshold"] = max(
                    0, min(1000, int(data.get("known_threshold", cfg["known_threshold"])))
                )
            except (TypeError, ValueError):
                pass
            break
        except Exception as exc:
            logging.getLogger(__name__).warning(
                "Failed to load sender status config %s: %s", path, exc
            )

    _SENDER_STATUS_CONFIG_CACHE = dict(cfg)
    return cfg


# Bound the dedup index so the history file cannot grow without limit.
_SENDER_HISTORY_MAX_SEEN = 20000
_SENDER_HISTORY_LOCK_STALE_SEC = 30


def _sender_history_path() -> Path:
    return _geofooter_data_dir() / "aes_sender_history.json"


def _acquire_sender_history_lock(timeout: float = 5.0) -> Optional[int]:
    """Guard the read-modify-write; several scans can run at once."""
    lock_path = str(_sender_history_path()) + ".lock"
    try:
        # Without this the very first open fails as a missing path, which is
        # indistinguishable here from a lock we could not take.
        _geofooter_data_dir().mkdir(parents=True, exist_ok=True)
    except OSError:
        return None
    deadline = time.time() + timeout
    while True:
        try:
            return os.open(lock_path, os.O_CREAT | os.O_EXCL | os.O_RDWR)
        except FileExistsError:
            try:
                # A killed scan leaves the lock behind; a stale one is not held.
                if time.time() - os.path.getmtime(lock_path) > _SENDER_HISTORY_LOCK_STALE_SEC:
                    os.unlink(lock_path)
                    continue
            except OSError:
                pass
            if time.time() >= deadline:
                return None
            time.sleep(0.05)
        except OSError:
            return None


def _release_sender_history_lock(fd: Optional[int]) -> None:
    if fd is None:
        return
    try:
        os.close(fd)
    except OSError:
        pass
    try:
        os.unlink(str(_sender_history_path()) + ".lock")
    except OSError:
        pass


def _read_sender_history() -> Dict[str, Any]:
    path = _sender_history_path()
    try:
        if path.is_file():
            data = json.loads(path.read_text(encoding="utf-8"))
            if isinstance(data, dict):
                return data
    except Exception as exc:
        logging.getLogger(__name__).warning("Failed to read sender history: %s", exc)
    return {}


def _write_sender_history(data: Dict[str, Any]) -> None:
    path = _sender_history_path()
    path.parent.mkdir(parents=True, exist_ok=True)
    data["updated"] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    tmp = Path(str(path) + ".tmp")
    tmp.write_text(json.dumps(data), encoding="utf-8")
    os.replace(str(tmp), str(path))


def sender_email_count(sender_email: Optional[str]) -> int:
    """How many distinct emails we have recorded from this sender."""
    sender = (sender_email or "").strip().lower()
    if not sender:
        return 0
    entry = (_read_sender_history().get("senders") or {}).get(sender) or {}
    try:
        return int(entry.get("count", 0))
    except (TypeError, ValueError):
        return 0


def record_sender_email(sender_email: Optional[str], message_id: Optional[str]) -> int:
    """Record one email from a sender and return the running total.

    Deduplicated on Message-ID so rescanning the same message never inflates
    the count. Returns the count already on file if the store cannot be
    updated, so a locked or unreadable file degrades to a stale reading
    rather than a wrong one.
    """
    sender = (sender_email or "").strip().lower()
    if not sender or "@" not in sender:
        return 0

    msg_id = (message_id or "").strip().lower()
    key = hashlib.sha1(msg_id.encode("utf-8")).hexdigest()[:16] if msg_id else ""

    fd = _acquire_sender_history_lock()
    if fd is None:
        # Another scan holds the lock. Report what is on file rather than an
        # increment we are not going to persist.
        return sender_email_count(sender)

    try:
        data = _read_sender_history()
        senders = data.setdefault("senders", {})
        seen = data.setdefault("seen", {})
        entry = senders.setdefault(sender, {"count": 0, "first_seen": "", "last_seen": ""})

        # Without a Message-ID there is nothing to deduplicate against, so the
        # mail is counted once and cannot be recognised again on a rescan.
        if key and key in seen:
            return int(entry.get("count", 0))

        stamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        entry["count"] = int(entry.get("count", 0) or 0) + 1
        if not entry.get("first_seen"):
            entry["first_seen"] = stamp
        entry["last_seen"] = stamp

        if key:
            seen[key] = stamp
            if len(seen) > _SENDER_HISTORY_MAX_SEEN:
                for old in sorted(seen, key=lambda k: seen[k])[
                    : len(seen) - _SENDER_HISTORY_MAX_SEEN
                ]:
                    seen.pop(old, None)

        _write_sender_history(data)
        return int(entry.get("count", 0))
    except Exception as exc:
        logging.getLogger(__name__).warning("Sender history update failed: %s", exc)
        return sender_email_count(sender)
    finally:
        _release_sender_history_lock(fd)


# Score smoothing: blend each scan's raw score with the sender's running
# average so soft-signal jitter (API wobble, lite vs full scans) does not
# bounce the displayed number. Rises track faster than falls, tiny moves
# pass through unchanged, and hard evidence (>= threshold) bypasses it.
SCORE_SMOOTHING_BYPASS = 70   # raw scores at/above this always show as-is
SCORE_SMOOTHING_DEADBAND = 3  # moves this small are not worth damping
SCORE_SMOOTHING_RISE = 0.6    # weight of the new score when rising
SCORE_SMOOTHING_FALL = 0.4    # weight of the new score when falling


def update_sender_score_ema(sender_email: Optional[str], raw_score: int) -> Optional[int]:
    """Blend raw_score with the sender's stored average and persist it.

    Returns the smoothed score, or None when smoothing does not apply
    (unknown sender, first sighting, hard-evidence bypass, or a locked /
    unreadable history file — all of which fall back to the raw score).
    """
    sender = (sender_email or "").strip().lower()
    if not sender or "@" not in sender:
        return None

    raw = max(0, min(100, int(raw_score)))
    fd = _acquire_sender_history_lock()
    if fd is None:
        return None

    try:
        data = _read_sender_history()
        senders = data.setdefault("senders", {})
        entry = senders.setdefault(sender, {"count": 0, "first_seen": "", "last_seen": ""})
        prev = entry.get("score_ema")

        if raw >= SCORE_SMOOTHING_BYPASS or prev is None:
            # Hard evidence shows immediately; a first sighting has no
            # history to smooth against. Either way the memory tracks raw.
            entry["score_ema"] = raw
            _write_sender_history(data)
            return None

        prev = max(0, min(100, int(prev)))
        delta = raw - prev
        if abs(delta) <= SCORE_SMOOTHING_DEADBAND:
            smoothed = raw
        else:
            weight = SCORE_SMOOTHING_RISE if delta > 0 else SCORE_SMOOTHING_FALL
            smoothed = int(round(prev + weight * delta))
        smoothed = max(0, min(100, smoothed))

        entry["score_ema"] = smoothed
        _write_sender_history(data)
        return smoothed if smoothed != raw else None
    except Exception as exc:
        logging.getLogger(__name__).warning("Sender score smoothing failed: %s", exc)
        return None
    finally:
        _release_sender_history_lock(fd)


# CompAuth Reason Codes
COMPAUTH_REASON_CODES = {
    "100": "Unknown reason (generic pass or no specific reason provided)",
    "101": "SPF aligned, no DKIM, no DMARC policy",
    "102": "SPF/DKIM aligned, no DMARC policy",
    "103": "SPF/DKIM aligned, DMARC policy quarantine",
    "104": "SPF/DKIM aligned, DMARC policy reject",
    "105": "DKIM aligned, no SPF, no DMARC policy",
    "106": "DKIM aligned, SPF failed, no DMARC policy",
    "107": "SPF failed, DKIM aligned, DMARC policy quarantine",
    "108": "SPF failed, DKIM aligned, DMARC policy reject",
    "109": "SPF/DKIM aligned, DMARC policy quarantine (duplicate for clarity)",
    "110": "SPF/DKIM aligned, DMARC policy reject (duplicate for clarity)",
    "200": "Authentication failed, no SPF/DKIM alignment",
    "201": "SPF failed, DKIM failed, no DMARC policy",
    "202": "SPF failed, DKIM failed, DMARC policy quarantine",
    "203": "SPF failed, DKIM failed, DMARC policy reject",
    "300": "Temporary authentication issue (e.g., timeout or server error)",
    "400": "Invalid or malformed authentication data",
    "601": "SPF/DKIM not aligned, DMARC policy quarantine",
    "602": "SPF/DKIM not aligned, no DMARC policy",
    "610": "SPF/DKIM not aligned, DMARC policy reject",
    "700": "Policy override (e.g., manual whitelist)",
    "800": "Authentication bypassed (e.g., internal relay)",
}

# Boolean security flags shown in footers / deep reports (order = display order).
SECURITY_FLAG_ORDER: Tuple[str, ...] = (
    "Tor",
    "Proxy",
    "Anonymous",
    "VPN",
    "Hosting",
    "Relay",
    "CDN",
    "Mobile",
    "Whitelisted",
)

# Flags that contribute to the shared "+10 once" active-threat cluster.
SECURITY_THREAT_FLAGS: Tuple[str, ...] = (
    "Tor",
    "Proxy",
    "Anonymous",
    "VPN",
    "Relay",
)

# No-data penalty (+2 each) applies only to these primary privacy flags.
SECURITY_NO_DATA_FLAGS: Tuple[str, ...] = (
    "Tor",
    "Proxy",
    "Anonymous",
    "VPN",
)

_CDN_ORG_RE = re.compile(
    r"\b(cdn|cloudflare|akamai|fastly|cloudfront|edgecast|limelight|"
    r"incapsula|stackpath|keycdn|bunnycdn|azure\s*cdn|google\s*cloud)\b",
    re.IGNORECASE,
)
_HOSTING_USAGE_RE = re.compile(
    r"data\s*cent(?:er|re)|web\s*hosting|\bhosting\b|\bdatacenter\b",
    re.IGNORECASE,
)
_CDN_USAGE_RE = re.compile(r"\bcdn\b|content\s*delivery", re.IGNORECASE)
_MOBILE_USAGE_RE = re.compile(
    r"\bmobile\b|\bcellular\b|\bwireless\b|\blte\b|\b5g\b",
    re.IGNORECASE,
)


def _default_security_data() -> Dict[str, Optional[bool]]:
    return {name: None for name in SECURITY_FLAG_ORDER}


def _set_security_flag(
    result: "GeoLocationResult",
    name: str,
    value: Any,
    *,
    upgrade_only: bool = True,
) -> None:
    """Set a boolean security flag. True wins over False/None when upgrade_only."""
    if name not in result.security_data:
        return
    if value is None:
        return
    flag = bool(value)
    current = result.security_data.get(name)
    if upgrade_only and current is True and not flag:
        return
    if current is None or not upgrade_only or flag:
        result.security_data[name] = flag


def _infer_usage_flags(result: "GeoLocationResult", usage_type: str, org: str = "") -> None:
    """Derive Hosting / CDN / Mobile from AbuseIPDB usageType and org heuristics."""
    text = f"{usage_type or ''} {org or ''}".strip()
    if not text:
        return
    if _HOSTING_USAGE_RE.search(text):
        _set_security_flag(result, "Hosting", True)
    if _CDN_USAGE_RE.search(text) or _CDN_ORG_RE.search(text):
        _set_security_flag(result, "CDN", True)
    if _MOBILE_USAGE_RE.search(usage_type or ""):
        _set_security_flag(result, "Mobile", True)


@dataclass
class GeoLocationResult:
    """Represents geolocation information for an IP address."""
    ip: str
    success: bool = False
    city: str = "Unknown"
    country: str = ""
    org: str = "Unknown"
    asn: str = "Unknown"
    blocklist: str = "Not Listed"
    lat: Optional[float] = None
    lon: Optional[float] = None
    security_data: Dict[str, Optional[bool]] = field(default_factory=_default_security_data)
    abuse_confidence: Optional[int] = None
    vpn_service: Optional[str] = None
    usage_type: Optional[str] = None
    timezone: Optional[str] = None
    geo_source: str = ""

@dataclass
class AuthenticationInfo:
    """Represents email authentication information."""
    auth_lines: List[str] = field(default_factory=list)
    spf: Optional[str] = None
    dkim: Optional[str] = None
    dmarc: Optional[str] = None
    compauth: Optional[Dict[str, str]] = None
    arc: Optional[Dict[str, Optional[str]]] = None
    arc_seal_valid: bool = False
    arc_chain_info: List[Dict[str, Any]] = field(default_factory=list)
    auth_type: str = "None"

@dataclass
class RiskScoreComponent:
    """One contribution used to compile the AES risk score (higher points = worse)."""
    category: str
    label: str
    points: int
    note: str = ""


@dataclass
class SecurityAssessment:
    """Represents the security risk assessment results."""
    score: int
    risk_level: str
    risk_color: str
    factors: List[str]
    domain_age_flag: Optional[str] = None
    components: List[RiskScoreComponent] = field(default_factory=list)

class LoggingSetup:
    """Handles logging configuration."""
    
    @staticmethod
    def setup():
        """Set up logging with rotation."""
        os.makedirs(CONFIG["base_path"], exist_ok=True)
        
        logger = logging.getLogger()
        logger.setLevel(logging.DEBUG)
        
        log_path = Path(CONFIG["base_path"]) / CONFIG["log_file"]
        handler = RotatingFileHandler(
            log_path,
            maxBytes=CONFIG["log_max_bytes"],
            backupCount=CONFIG["log_backup_count"]
        )
        
        tz = pytz.utc
        
        class TimezoneFormatter(logging.Formatter):
            def formatTime(self, record, datefmt=None):
                dt = datetime.fromtimestamp(record.created, tz)
                if datefmt:
                    return dt.strftime(datefmt)
                return dt.strftime('%Y-%m-%d %H:%M:%S,%f')[:-3] + f' {tz.tzname(dt)}'
        
        formatter = TimezoneFormatter(
            '%(asctime)s - %(name)s - %(levelname)s - %(message)s',
            datefmt='%Y-%m-%d %H:%M:%S,%f'
        )
        handler.setFormatter(formatter)
        logger.addHandler(handler)

class IPValidator:
    """Handles IP address validation."""
    
    @staticmethod
    def is_private(ip: str) -> bool:
        """Check if IP is private."""
        try:
            ip_obj = ipaddress.ip_address(ip)
            return (ip_obj.is_private or ip_obj.is_loopback or 
                   ip_obj.is_link_local or ip_obj.is_multicast)
        except ValueError:
            return False
    
    @staticmethod
    def is_valid_public(ip: str) -> bool:
        """Check if IP is valid and public."""
        if not ip:
            return False
        try:
            ip_obj = ipaddress.ip_address(ip)
            return (ip_obj.is_global and not ip_obj.is_multicast and
                   not ip_obj.is_reserved and not ip_obj.is_private)
        except ValueError:
            return False

class HeaderParser:
    """Parses email headers for various information."""
    
    def __init__(self):
        self.logger = logging.getLogger(__name__)

    @staticmethod
    def unfold_headers(headers: str) -> str:
        """Join RFC5322 folded continuation lines into single logical lines."""
        if not headers:
            return ""
        joined: List[str] = []
        for line in headers.splitlines():
            if line.startswith((" ", "\t")) and joined:
                joined[-1] += " " + line.strip()
            else:
                joined.append(line.rstrip())
        return "\n".join(joined)

    @staticmethod
    def extract_received_blocks(headers: Optional[str]) -> List[str]:
        """Return full unfolded Received: header values (one string per hop)."""
        if not headers:
            return []
        flat = HeaderParser.unfold_headers(headers)
        return re.findall(r"^Received:\s*.+$", flat, re.MULTILINE | re.IGNORECASE)

    @staticmethod
    def parse_export_file(raw_content: str) -> Tuple[str, Dict[str, str]]:
        """Split VBA export preamble from transport headers."""
        metadata: Dict[str, str] = {}
        if "----------------------------------------" not in raw_content:
            return raw_content, metadata

        preamble, _, body = raw_content.partition("----------------------------------------")
        for line in preamble.splitlines():
            line = line.strip()
            if not line or ":" not in line:
                continue
            key, value = line.split(":", 1)
            metadata[key.strip().lower()] = value.strip()

        return body.strip(), metadata

    @staticmethod
    def _normalize_email(raw: Optional[str]) -> Optional[str]:
        if not raw:
            return None
        email = raw.strip().strip("<>").strip('"').strip()
        if "@" not in email:
            return None
        return email.lower()

    @staticmethod
    def _domain_from_email(email: Optional[str]) -> Optional[str]:
        if not email or "@" not in email:
            return None
        return email.split("@", 1)[1].lower().strip(">.")

    def extract_sender_info(self, headers: str) -> Tuple[Optional[str], Optional[str]]:
        """Extract sender email and domain from headers."""
        patterns = [
            r'^From:\s*<([^>]+)>',
            r'^From:\s*"[^"]*"\s*<([^>]+)>',
            r'^From:\s*([^<\s]+@[^>\s]+)',
            r'^From:\s*[^<]*<([^>]+)>',
            r'^Reply-To:\s*<([^>]+)>',
            r'^Reply-To:\s*([^<\s]+@[^>\s]+)',
            r'^Return-Path:\s*<([^>]+)>',
            r'^Return-Path:\s*([^<\s]+@[^>\s]+)',
            r'^Sender:\s*<([^>]+)>',
            r'^Sender:\s*([^<\s]+@[^>\s]+)',
            r'^Received-SPF:[^\n]*\bmailfrom=([^;\s>]+)',
            r'^Authentication-Results:[^\n]*\bsmtp\.mailfrom=([^;\s>]+)',
            r'^X-MS-Exchange-Organization-OriginalArrivalTime:[^\n]*\bfrom=([^;\s>]+@[^;\s>]+)',
        ]
        for pattern in patterns:
            match = re.search(pattern, headers, re.IGNORECASE | re.MULTILINE)
            if match:
                sender_email = self._normalize_email(match.group(1))
                if sender_email:
                    sender_domain = self._domain_from_email(sender_email)
                    self.logger.info(f"Extracted sender: {sender_email}, domain: {sender_domain}")
                    return sender_email, sender_domain

        # Fallback: domain from DKIM header d= attribute
        dkim_domain = re.search(r'^DKIM-Signature:[^\n]*\bd=([^;\s]+)', headers, re.IGNORECASE | re.MULTILINE)
        if dkim_domain:
            sender_domain = dkim_domain.group(1).lower()
            self.logger.info(f"Inferred domain from DKIM-Signature: {sender_domain}")
            return None, sender_domain

        self.logger.warning("Could not extract sender from headers")
        return None, None
    
    def extract_original_sender_ip(self, headers: str) -> Tuple[Optional[str], Optional[str]]:
        """Extract the original sender's public IP and hostname."""
        self.logger.info("Extracting original sender IP and hostname")

        headers_flat = self.unfold_headers(headers)

        # Pattern for: Received: from <hostname> ([<ip>])
        pattern_bracket = r'^Received:\s*from\s+([^\s]+)\s*\(\[([\d\.:a-fA-F]+)\]\)'  # e.g., from M1 ([51.179.200.191])
        match_bracket = re.findall(pattern_bracket, headers_flat, re.IGNORECASE | re.MULTILINE)
        for hostname, ip in reversed(match_bracket):
            self.logger.debug(f"Processing (bracket): Hostname={hostname}, IP={ip}")
            if IPValidator.is_valid_public(ip):
                self.logger.info(f"Found public IP: {ip} (Hostname: {hostname})")
                self.logger.debug(f"DEBUG: Found public IP: {ip} (Hostname: {hostname})")
                return ip, hostname

        # Existing patterns
        pattern = (r'^Received:\s*from\s+([^\s\(]+)\s*'
                  r'(?:\([^)]*\))?\s*'
                  r'(?:by\s+[^\s]+\s*)?'
                  r'(?:\([^)]*\[([^\]]+)\][^)]*\))?')
        simple_pattern = r'^Received:\s*from\s+([^\s]+)\s+\(([^\[]+)\[([^\]]+)\]\)'

        simple_headers = re.findall(simple_pattern, headers_flat, re.IGNORECASE | re.MULTILINE)
        for hostname, hostname2, ip in reversed(simple_headers):
            self.logger.debug(f"Processing (simple): Hostname={hostname}, IP={ip}")
            if IPValidator.is_valid_public(ip):
                self.logger.info(f"Found public IP: {ip} (Hostname: {hostname})")
                self.logger.debug(f"DEBUG: Found public IP: {ip} (Hostname: {hostname})")
                return ip, hostname

        received_headers = re.findall(pattern, headers_flat, re.IGNORECASE | re.MULTILINE)
        for hostname, ip in reversed(received_headers):
            if ip:
                self.logger.debug(f"Processing: Hostname={hostname}, IP={ip}")
                if IPValidator.is_valid_public(ip):
                    self.logger.info(f"Found public IP: {ip} (Hostname: {hostname})")
                    self.logger.debug(f"DEBUG: Found public IP: {ip} (Hostname: {hostname})")
                    return ip, hostname

        # Fallback: client-ip from Received-SPF
        spf_client_ip = re.search(r'^Received-SPF:[^\n]*\bclient-ip=([\d\.:a-fA-F]+)', headers_flat, re.IGNORECASE | re.MULTILINE)
        if spf_client_ip:
            ip = spf_client_ip.group(1)
            if IPValidator.is_valid_public(ip):
                self.logger.info(f"Found public IP in Received-SPF client-ip: {ip}")
                return ip, None

        # Fallback: Try to extract sender IP from Authentication-Results header
        auth_results_pattern = r'^Authentication-Results:.*?(sender IP is ([\d\.]+))'
        auth_results_match = re.search(auth_results_pattern, headers_flat, re.IGNORECASE | re.MULTILINE)
        if auth_results_match:
            ip = auth_results_match.group(2)
            if IPValidator.is_valid_public(ip):
                self.logger.info(f"Found public IP in Authentication-Results: {ip}")
                self.logger.debug(f"DEBUG: Found public IP in Authentication-Results: {ip}")
                return ip, None

        self.logger.warning("No valid public sender IP found")
        self.logger.debug("DEBUG: No valid public sender IP found")
        return None, None

class AuthenticationParser:
    """Parses authentication-related headers."""
    
    def __init__(self):
        self.logger = logging.getLogger(__name__)
    
    def extract_authentication_info(self, headers: str) -> AuthenticationInfo:
        """Extract authentication information from headers."""
        auth_info = AuthenticationInfo()

        auth_pattern = r'^Authentication-Results:\s*([^\n]+(?:\n\s+[^\n]+)*)'
        for auth_match in re.finditer(auth_pattern, headers, re.IGNORECASE | re.MULTILINE):
            auth_content = auth_match.group(1).replace('\n', ' ').replace('\r', ' ').strip()
            auth_info.auth_lines.append(auth_match.group(0).strip())
            self.logger.info(f"Found Authentication-Results: {auth_content}")
            self._merge_auth_content(auth_content, auth_info)

        if not auth_info.spf:
            spf_line = re.search(
                r'^Received-SPF:\s*(Pass|Fail|Softfail|Neutral|None|Permerror|Temperror)',
                headers,
                re.IGNORECASE | re.MULTILINE,
            )
            if spf_line:
                auth_info.spf = spf_line.group(1).lower()
                self.logger.debug(f"DEBUG: SPF from Received-SPF: {auth_info.spf}")

        self._parse_arc_headers(headers, auth_info)
        if not auth_info.arc:
            auth_info.arc = {"result": "none"}

        self._determine_auth_type(auth_info)

        if not auth_info.compauth:
            auth_info.compauth = {"result": "missing", "reason_code": "not_present"}

        return auth_info

    def _merge_auth_content(self, auth_content: str, auth_info: AuthenticationInfo) -> None:
        """Merge SPF/DKIM/DMARC/CompAuth from one Authentication-Results block."""

        def prefer(existing: Optional[str], new_val: str) -> bool:
            if not new_val:
                return False
            if not existing:
                return True
            if new_val.lower() == "pass":
                return True
            if existing.lower() == "pass":
                return False
            return existing.lower() in {"none", "missing", "neutral", "temperror", "permerror"}

        spf_match = re.search(r'spf=(\w+)', auth_content, re.IGNORECASE)
        if spf_match and prefer(auth_info.spf, spf_match.group(1).lower()):
            auth_info.spf = spf_match.group(1).lower()

        dkim_match = re.search(r'dkim=(\w+)', auth_content, re.IGNORECASE)
        if dkim_match and prefer(auth_info.dkim, dkim_match.group(1).lower()):
            auth_info.dkim = dkim_match.group(1).lower()

        dmarc_match = re.search(r'dmarc=(\w+)', auth_content, re.IGNORECASE)
        if dmarc_match:
            dmarc_val = dmarc_match.group(1).lower()
            if dmarc_val == "bestguesspass":
                dmarc_val = "pass"
            if prefer(auth_info.dmarc, dmarc_val):
                auth_info.dmarc = dmarc_val

        compauth_match = re.search(r'compauth=(\w+)(?:\s+reason=(\d+))?', auth_content, re.IGNORECASE)
        if compauth_match:
            comp_result = compauth_match.group(1).lower()
            comp_reason = compauth_match.group(2) if compauth_match.group(2) else "unknown"
            if not auth_info.compauth or comp_result == "pass" or auth_info.compauth.get("result") != "pass":
                auth_info.compauth = {"result": comp_result, "reason_code": comp_reason}
    
    def _parse_arc_headers(self, headers: str, auth_info: AuthenticationInfo):
        """Parse ARC-related headers."""
        arc_match = re.search(
            r'^ARC-Authentication-Results:[^\n]*\b(?:arc=(\w+)|spf=(\w+))',
            headers,
            re.IGNORECASE | re.MULTILINE,
        )
        if arc_match:
            arc_result = arc_match.group(1) or arc_match.group(2)
            if arc_result:
                auth_info.arc = {"result": arc_result.lower()}
    
    def _determine_auth_type(self, auth_info: AuthenticationInfo):
        """Determine the overall authentication type."""
        if auth_info.compauth and auth_info.compauth.get("result") != "missing":
            auth_info.auth_type = "Standard"
        else:
            auth_info.auth_type = "Basic"

class GeoLocationService:
    """Handles IP geolocation and security checks."""
    
    def __init__(self):
        self.logger = logging.getLogger(__name__)
    
    def geolocate_ip(self, ip: str) -> GeoLocationResult:
        """Get geolocation and security information for an IP."""
        result = GeoLocationResult(ip)
        
        if not IPValidator.is_valid_public(ip):
            result.city = "Invalid or Private IP"
            result.blocklist = "N/A"
            return result
        
        self._try_abuseipdb(ip, result)
        self._try_ipinfo(ip, result)
        # Always enrich privacy flags from ip-api + proxycheck (not only on failure).
        self._try_ip_api(ip, result, geolocate=not result.success)
        self._try_proxycheck(ip, result)
        
        result.blocklist = self._check_blocklists(ip)
        
        return result
    
    def _try_abuseipdb(self, ip: str, result: GeoLocationResult) -> bool:
        """Try AbuseIPDB for security flags, usage type, and abuse confidence."""
        try:
            api_key = _resolve_abuseipdb_api_key()
            if not api_key:
                self.logger.warning("AbuseIPDB skipped: no API key configured (AES Settings).")
                return False
            headers = {"Key": api_key, "Accept": "application/json"}
            _rate_limit_abuseipdb()
            response, _insecure = _http_get_with_retry(
                abuseipdb_check_url(),
                headers=headers,
                params={"ipAddress": ip, "maxAgeInDays": 90},
                timeout=15,
                retries=3,
            )
            data = response.json()
            self.logger.debug(f"AbuseIPDB response: {data}")
            
            abuse_data = data.get("data", {})
            if abuse_data:
                _set_security_flag(result, "Tor", abuse_data.get("isTor"))
                _set_security_flag(result, "Proxy", abuse_data.get("isProxy"))
                _set_security_flag(result, "VPN", abuse_data.get("isVpn"))
                _set_security_flag(result, "Anonymous", abuse_data.get("isAnonymous"))
                _set_security_flag(result, "Whitelisted", abuse_data.get("isWhitelisted"))

                score = abuse_data.get("abuseConfidenceScore")
                if score is not None:
                    try:
                        result.abuse_confidence = int(score)
                    except (TypeError, ValueError):
                        pass

                usage = str(abuse_data.get("usageType") or "").strip()
                if usage:
                    result.usage_type = usage
                    org_hint = str(abuse_data.get("isp") or abuse_data.get("domain") or "")
                    _infer_usage_flags(result, usage, result.org or org_hint)
                return True
        except Exception as e:
            self.logger.error(f"AbuseIPDB error for {ip}: {e}")
        return False
    
    def _try_ipinfo(self, ip: str, result: GeoLocationResult) -> bool:
        """Try ipinfo.io for geolocation and security flags."""
        try:
            response, _insecure = _http_get_with_retry(
                f"https://ipinfo.io/{urllib.parse.quote(ip, safe='')}/json",
                params={"token": CONFIG["api_keys"]["ipinfo"]},
                timeout=12,
                retries=2,
            )
            data = response.json()
            self.logger.debug(f"IPInfo response: {data}")
            # Rate-limit / bad-token replies are JSON too; treating them as a
            # hit would suppress the ip-api geolocation fallback.
            if not isinstance(data, dict) or data.get("error") or not (
                data.get("country") or data.get("city")
            ):
                self.logger.warning(f"ipinfo.io returned no location for {ip}: {str(data)[:200]}")
                return False

            result.success = True
            result.geo_source = "ipinfo.io"
            result.city = data.get("city", "Unknown") or "Unknown"
            result.country = data.get("country", "") or ""
            result.org = data.get("org", "Unknown") or "Unknown"
            if not result.timezone and data.get("timezone"):
                result.timezone = str(data["timezone"])

            loc = data.get("loc") or ""
            if isinstance(loc, str) and "," in loc:
                try:
                    lat_s, lon_s = loc.split(",", 1)
                    result.lat = float(lat_s.strip())
                    result.lon = float(lon_s.strip())
                except ValueError:
                    pass
            
            # ASN extraction
            if "asn" in data and isinstance(data["asn"], dict):
                result.asn = data["asn"].get("asn", "Unknown") or "Unknown"
            else:
                # Try to extract ASN from org string if present
                org_str = result.org or ""
                match = re.search(r"AS\d+", org_str)
                if match:
                    result.asn = match.group(0)
                else:
                    result.asn = "Unknown"
            
            privacy = data.get("privacy") or {}
            if isinstance(privacy, dict):
                for flag in ("VPN", "Proxy", "Tor", "Anonymous"):
                    if result.security_data.get(flag) is None:
                        _set_security_flag(result, flag, privacy.get(flag.lower()))
                if "hosting" in privacy:
                    _set_security_flag(result, "Hosting", privacy.get("hosting"))
                if "relay" in privacy:
                    _set_security_flag(result, "Relay", privacy.get("relay"))
                service = str(privacy.get("service") or "").strip()
                if service:
                    result.vpn_service = service
                    _set_security_flag(result, "VPN", True)

            _infer_usage_flags(result, result.usage_type or "", result.org or "")
            return True
        except Exception as e:
            self.logger.error(f"ipinfo.io error for {ip}: {e}")
        return False
    
    def _try_ip_api(self, ip: str, result: GeoLocationResult, *, geolocate: bool = True) -> bool:
        """ip-api.com — geolocation fallback and/or proxy/hosting/mobile enrichment."""
        try:
            quoted = urllib.parse.quote(ip, safe="")
            url = f"http://ip-api.com/json/{quoted}"
            response, _insecure = _http_get_with_retry(
                url,
                params={"fields": "66777215"},
                timeout=12,
                retries=2,
            )
            data = response.json()
            if data.get("status") != "success":
                return False

            if not geolocate and result.city in ("", "Unknown") and data.get("city"):
                cc = data.get("countryCode", "") or ""
                if not result.country or result.country == cc:
                    result.city = data["city"]
                    result.country = result.country or cc
                    result.geo_source = f"{result.geo_source or 'ipinfo.io'} + ip-api.com"
            if geolocate:
                result.success = True
                result.geo_source = "ip-api.com"
                result.city = data.get("city", "Unknown") or "Unknown"
                result.country = data.get("countryCode", "") or data.get("country", "") or ""
                result.org = data.get("org", "Unknown") or data.get("isp", "Unknown") or "Unknown"
                if data.get("lat") is not None and data.get("lon") is not None:
                    try:
                        result.lat = float(data["lat"])
                        result.lon = float(data["lon"])
                    except (TypeError, ValueError):
                        pass
                asn_val = data.get("as", "Unknown") or "Unknown"
                if asn_val != "Unknown":
                    match = re.search(r"AS\d+", asn_val)
                    result.asn = match.group(0) if match else asn_val
                else:
                    org_str = result.org or ""
                    match = re.search(r"AS\d+", org_str)
                    result.asn = match.group(0) if match else "Unknown"

            if not result.timezone and data.get("timezone"):
                result.timezone = str(data["timezone"])
            if "proxy" in data and result.security_data.get("Proxy") is None:
                _set_security_flag(result, "Proxy", data.get("proxy"))
            if "hosting" in data:
                _set_security_flag(result, "Hosting", data.get("hosting"))
            if "mobile" in data:
                _set_security_flag(result, "Mobile", data.get("mobile"))
            _infer_usage_flags(result, result.usage_type or "", result.org or "")
            return True
        except Exception as e:
            self.logger.error(f"ip-api.com error for {ip}: {e}")
        if geolocate:
            result.city = "Geolocation Unavailable"
            result.blocklist = "N/A (Geolocation Failed)"
        return False

    def _try_proxycheck(self, ip: str, result: GeoLocationResult) -> bool:
        """proxycheck.io free tier — VPN/proxy/TOR when other sources leave gaps."""
        try:
            quoted = urllib.parse.quote(ip, safe="")
            response, _insecure = _http_get_with_retry(
                f"https://proxycheck.io/v2/{quoted}",
                params={"vpn": "1", "asn": "1"},
                timeout=12,
                retries=2,
            )
            payload = response.json()
            self.logger.debug(f"proxycheck.io response: {payload}")
            if payload.get("status") != "ok":
                return False
            entry = _proxycheck_entry(payload, ip)
            if not entry:
                return False
            _apply_proxycheck(result, entry)
            return True
        except Exception as e:
            self.logger.error(f"proxycheck.io error for {ip}: {e}")
        return False
    
    def _check_blocklists(self, ip: str) -> str:
        """Check IP against multiple blocklists (IPv4 and IPv6)."""
        try:
            ip_obj = ipaddress.ip_address(ip)
            is_ipv6 = ip_obj.version == 6
            
            blocklists = CONFIG["blocklists"]["ipv6" if is_ipv6 else "ipv4"]
            listed_in = []
            unavailable = []

            for bl in blocklists:
                query = _blocklist_dns_query(ip, bl)
                if not query:
                    continue
                if threat_intel is not None:
                    # Local resolvers often NXDOMAIN every blocklist query; ask the
                    # list's own servers and verify with its test entry first.
                    answers = threat_intel.dnsbl_lookup(query[: -len(bl) - 1], bl)
                    if answers is None:
                        unavailable.append(bl)
                    # 127.255.255.x / 127.0.0.1 are "query refused" codes, not listings.
                    elif any(a.startswith("127.") and not a.startswith("127.255.") and a != "127.0.0.1"
                             for a in answers):
                        listed_in.append(bl)
                elif self._check_single_blocklist(query):
                    listed_in.append(bl)

            if listed_in:
                return f"Listed: {', '.join(listed_in)}"
            if unavailable and len(unavailable) == len(blocklists):
                return "Unavailable (blocklist DNS unreachable)"
            return "Not Listed"
            
        except ValueError:
            return "N/A (Invalid IP)"
    
    def _check_single_blocklist(self, query: str) -> bool:
        """Check a single blocklist query."""
        try:
            dns.resolver.resolve(query, "A")
            return True
        except (dns.resolver.NXDOMAIN, dns.resolver.NoAnswer):
            return False
        except Exception as e:
            self.logger.error(f"Blocklist check error for {query}: {e}")
            return False

class WhoisService:
    """Handles WHOIS lookups."""
    
    def __init__(self):
        self.logger = logging.getLogger(__name__)
    
    def get_whois_info(self, domain: str) -> Dict[str, Any]:
        """Get WHOIS information for a domain (python-whois, then RDAP fallback)."""
        if not domain or domain.lower() in {"unknown", "n/a", "none"}:
            return {"error": "Lookup not performed"}

        domain = domain.strip().lower().rstrip(".")
        # Prefer registrable domain for lookups (e.g. mail.example.co.uk -> example.co.uk).
        lookup_domain = self._registrable_domain(domain) or domain

        try:
            w = whois.whois(lookup_domain)
            registrar = self._get_value(w.registrar, "Unknown")
            creation_date = self._normalize_date(w.creation_date)
            expiration_date = self._normalize_date(w.expiration_date)
            status = self._get_value(w.status, "Unknown")

            result = {
                "registrar": registrar,
                "creation_date": creation_date,
                "expiration_date": expiration_date,
                "status": status,
            }
            if registrar != "Unknown" or creation_date != "Unknown":
                self.logger.info(f"WHOIS for {lookup_domain}: {result}")
                return result
            self.logger.warning(
                "WHOIS returned empty registrar/creation for %s; trying RDAP", lookup_domain
            )
        except Exception as e:
            self.logger.error(f"WHOIS lookup failed for {lookup_domain}: {e}")

        rdap = self._rdap_whois(lookup_domain)
        if rdap:
            self.logger.info(f"RDAP WHOIS for {lookup_domain}: {rdap}")
            return rdap
        return {"error": f"WHOIS/RDAP unavailable for {lookup_domain}"}

    def _registrable_domain(self, domain: str) -> str:
        """Best-effort eTLD+1 for common UK / multi-part suffixes."""
        parts = [p for p in domain.split(".") if p]
        if len(parts) < 2:
            return domain
        multi = {
            "co.uk", "org.uk", "me.uk", "net.uk", "ac.uk", "gov.uk",
            "com.au", "co.nz", "co.jp", "com.br",
        }
        last2 = ".".join(parts[-2:])
        if last2 in multi and len(parts) >= 3:
            return ".".join(parts[-3:])
        return ".".join(parts[-2:])

    def _rdap_whois(self, domain: str) -> Optional[Dict[str, Any]]:
        """RDAP fallback when classic WHOIS socket lookup fails (common for .uk)."""
        urls: List[str] = []
        if domain.endswith(".uk"):
            urls.append(f"https://rdap.nominet.uk/uk/domain/{domain}")
        urls.append(f"https://rdap.org/domain/{domain}")

        for url in urls:
            try:
                resp, _ = _http_get_with_retry(url, timeout=12, retries=1)
                if resp is None or getattr(resp, "status_code", 0) != 200:
                    continue
                data = resp.json()
                registrar = self._rdap_registrar(data) or "Unknown"
                creation = self._rdap_event_date(data, "registration") or "Unknown"
                expiration = self._rdap_event_date(data, "expiration") or "Unknown"
                status = data.get("status") or "Unknown"
                if isinstance(status, list):
                    status = ", ".join(str(s) for s in status) if status else "Unknown"
                return {
                    "registrar": registrar,
                    "creation_date": self._normalize_date(creation),
                    "expiration_date": self._normalize_date(expiration),
                    "status": str(status),
                    "source": "rdap",
                }
            except Exception as exc:
                self.logger.debug("RDAP lookup failed via %s: %s", url, exc)
        return None

    @staticmethod
    def _rdap_event_date(data: Dict[str, Any], action: str) -> Optional[str]:
        for event in data.get("events") or []:
            if str(event.get("eventAction") or "").lower() == action.lower():
                return event.get("eventDate")
        return None

    @staticmethod
    def _rdap_registrar(data: Dict[str, Any]) -> Optional[str]:
        for ent in data.get("entities") or []:
            roles = [str(r).lower() for r in (ent.get("roles") or [])]
            if "registrar" not in roles:
                continue
            vcard = ent.get("vcardArray")
            if isinstance(vcard, list) and len(vcard) > 1:
                for item in vcard[1]:
                    if not item or not isinstance(item, list):
                        continue
                    if item[0] in ("fn", "org") and len(item) > 3 and item[3]:
                        return str(item[3])
            name = ent.get("handle") or ent.get("fn")
            if name:
                return str(name)
        return None

    def get_cert_history(self, domain: str) -> Dict[str, Any]:
        """Certificate Transparency history via crt.sh (free, no API key).

        A domain whose first TLS certificate is only days old is a classic
        phishing-infrastructure tell, independent of WHOIS privacy shields.
        """
        if not domain or domain.lower() in {"unknown", "n/a", "none"}:
            return {"error": "Lookup not performed"}
        lookup = self._registrable_domain(domain.strip().lower().rstrip(".")) or domain

        cached = self._crtsh_cache_get(lookup)
        if cached is not None:
            return cached

        try:
            entries: Optional[List[Any]] = None
            active_only = False
            try:
                resp, _ = _http_get_with_retry(
                    "https://crt.sh/",
                    params={"q": lookup, "output": "json"},
                    timeout=12,
                    retries=0,
                )
                if resp is not None and getattr(resp, "status_code", 0) == 200:
                    entries = resp.json()
            except Exception:
                entries = None

            if entries is None:
                # Big domains (thousands of historical certs) time out on the
                # full query; unexpired-only with dedup is far cheaper.
                resp, _ = _http_get_with_retry(
                    "https://crt.sh/",
                    params={
                        "q": lookup,
                        "output": "json",
                        "exclude": "expired",
                        "deduplicate": "Y",
                    },
                    timeout=12,
                    retries=0,
                )
                if resp is None or getattr(resp, "status_code", 0) != 200:
                    return {"error": "crt.sh unavailable"}
                entries = resp.json()
                active_only = True

            if not isinstance(entries, list) or not entries:
                return self._crtsh_cache_put(lookup, {"error": "No certificates logged"})

            first_seen: Optional[datetime] = None
            newest: Optional[datetime] = None
            recent_30d = 0
            cutoff = datetime.utcnow() - timedelta(days=30)
            for entry in entries:
                raw = str(entry.get("not_before") or "")[:19]
                try:
                    seen = datetime.strptime(raw, "%Y-%m-%dT%H:%M:%S")
                except ValueError:
                    continue
                if first_seen is None or seen < first_seen:
                    first_seen = seen
                if newest is None or seen > newest:
                    newest = seen
                if seen >= cutoff:
                    recent_30d += 1

            if first_seen is None:
                return {"error": "No parseable certificates"}
            return self._crtsh_cache_put(lookup, {
                "domain": lookup,
                "first_seen": first_seen.strftime("%Y-%m-%d"),
                "age_days": (datetime.utcnow() - first_seen).days,
                "newest": newest.strftime("%Y-%m-%d") if newest else "Unknown",
                "total": len(entries),
                "recent_30d": recent_30d,
                "active_only": active_only,
            })
        except Exception as exc:
            self.logger.error(f"crt.sh lookup failed for {lookup}: {exc}")
            return {"error": "crt.sh lookup failed"}

    _CRTSH_CACHE_TTL_DAYS = 7

    def _crtsh_cache_path(self) -> Path:
        return Path(os.environ.get("LOCALAPPDATA", CONFIG["base_path"])) / "GeoFooter" / "crtsh_cache.json"

    def _crtsh_cache_get(self, domain: str) -> Optional[Dict[str, Any]]:
        try:
            path = self._crtsh_cache_path()
            if not path.is_file():
                return None
            cache = json.loads(path.read_text(encoding="utf-8"))
            entry = cache.get(domain)
            if not isinstance(entry, dict):
                return None
            cached_at = datetime.strptime(str(entry.get("cached_at", "")), "%Y-%m-%d %H:%M:%S")
            if datetime.utcnow() - cached_at > timedelta(days=self._CRTSH_CACHE_TTL_DAYS):
                return None
            result = dict(entry)
            result.pop("cached_at", None)
            # Age keeps moving even while the lookup itself is cached.
            if "first_seen" in result:
                try:
                    first = datetime.strptime(str(result["first_seen"]), "%Y-%m-%d")
                    result["age_days"] = (datetime.utcnow() - first).days
                except ValueError:
                    pass
            return result
        except Exception:
            return None

    def _crtsh_cache_put(self, domain: str, result: Dict[str, Any]) -> Dict[str, Any]:
        try:
            path = self._crtsh_cache_path()
            cache: Dict[str, Any] = {}
            if path.is_file():
                try:
                    cache = json.loads(path.read_text(encoding="utf-8"))
                except Exception:
                    cache = {}
            if not isinstance(cache, dict):
                cache = {}
            entry = dict(result)
            entry["cached_at"] = datetime.utcnow().strftime("%Y-%m-%d %H:%M:%S")
            cache[domain] = entry
            if len(cache) > 500:
                cache = dict(list(cache.items())[-400:])
            path.parent.mkdir(parents=True, exist_ok=True)
            path.write_text(json.dumps(cache, indent=1), encoding="utf-8")
        except Exception:
            pass
        return result

    def _get_value(self, value: Any, default: str) -> str:
        """Get value or return default."""
        if not value or (isinstance(value, str) and value.lower() == "unknown"):
            return default
        if isinstance(value, list):
            value = value[0] if value else default
        return str(value) if value else default

    def _normalize_date(self, date_value: Any) -> str:
        """Normalize date to string format."""
        if isinstance(date_value, list):
            date_value = date_value[0] if date_value else None

        if isinstance(date_value, datetime):
            return date_value.strftime("%Y-%m-%d")

        if not date_value or (isinstance(date_value, str) and date_value.lower() == "unknown"):
            return "Unknown"

        # RDAP ISO timestamps: 2001-08-06T08:31:21Z
        text = str(date_value).strip()
        if "T" in text:
            try:
                parsed = datetime.fromisoformat(text.replace("Z", "+00:00"))
                return parsed.strftime("%Y-%m-%d")
            except Exception:
                return text.split("T")[0]
        return text

class SecurityAssessor:
    """Assesses security risk based on various factors."""
    
    def __init__(self):
        self.logger = logging.getLogger(__name__)
    
    @staticmethod
    def _record_factor(
        factors: List[str],
        components: List[RiskScoreComponent],
        category: str,
        label: str,
        points: int,
        note: str = "",
    ) -> None:
        factors.append(label)
        components.append(
            RiskScoreComponent(category=category, label=label, points=points, note=note)
        )

    def assess_security_risk(
        self,
        geo_results: List[GeoLocationResult],
        auth_info: AuthenticationInfo,
        sender_domain_whois: Dict[str, Any]
    ) -> SecurityAssessment:
        """Calculate security risk score and assessment."""
        score = 0
        factors: List[str] = []
        components: List[RiskScoreComponent] = []
        domain_age_flag = None
        
        score = self._assess_authentication(auth_info, score, factors, components)
        score = self._assess_geolocation(geo_results, score, factors, components)
        score, domain_age_flag = self._assess_domain_age(
            sender_domain_whois, score, factors, components
        )
        
        score = max(0, min(100, int(score)))
        risk_level = self._calculate_risk_level(score, domain_age_flag)
        risk_color = risk_color_for_level(risk_level)
        
        return SecurityAssessment(
            score=score,
            risk_level=risk_level,
            risk_color=risk_color,
            factors=factors,
            domain_age_flag=domain_age_flag,
            components=components,
        )

    def apply_route_path_risk(
        self,
        assessment: SecurityAssessment,
        per_hop_analysis: Optional[List[Dict[str, Any]]],
        route_config: Optional[Dict[str, Any]] = None,
    ) -> SecurityAssessment:
        """Add +15 once when any hop is high-risk (AbuseIPDB threshold or listed ASN)."""
        if assessment is None:
            return assessment

        cfg = route_config or load_route_risk_config()
        if not cfg.get("enabled", True):
            return assessment

        provider = str(cfg.get("provider") or "abuseipdb").lower()
        # Only AbuseIPDB is implemented; other providers are ignored until added.
        if provider and provider != "abuseipdb":
            self.logger.info(
                "Route risk provider %s not implemented; using abuseipdb hop scores only.",
                provider,
            )

        threshold = int(cfg.get("abuse_threshold", 25) or 25)
        extra_asns = {
            _normalize_asn(a)
            for a in (cfg.get("extra_high_risk_asns") or [])
            if _normalize_asn(a)
        }

        # Graded band instead of a hard cliff: hops just over the threshold
        # add +8; only well past it (or a listed ASN) adds the full +15.
        full_band = threshold + 15
        hits: List[str] = []
        route_pts = 0
        for hop in per_hop_analysis or []:
            asn = _normalize_asn(hop.get("asn"))
            abuse_score = None
            rep = str(hop.get("reputation") or "")
            m = re.search(r"Score:\s*(\d+)", rep, re.IGNORECASE)
            if m:
                try:
                    abuse_score = int(m.group(1))
                except ValueError:
                    abuse_score = None

            reasons = []
            hop_pts = 0
            if abuse_score is not None and abuse_score >= threshold:
                reasons.append(f"AbuseIPDB {abuse_score}")
                hop_pts = 15 if abuse_score >= full_band else 8
            if asn and asn in extra_asns:
                reasons.append("listed ASN")
                hop_pts = 15
            if not reasons:
                continue
            route_pts = max(route_pts, hop_pts)
            label = asn or (hop.get("ip") or "hop")
            hits.append(f"{label} ({'; '.join(reasons)})")

        if not hits:
            return assessment

        factors = list(assessment.factors or [])
        components = list(assessment.components or [])
        hit_summary = "; ".join(hits[:4])
        if len(hits) > 4:
            hit_summary += f"; +{len(hits) - 4} more"
        self._record_factor(
            factors,
            components,
            "Origin / network",
            f"High-risk ASN/route hop ({hit_summary})",
            route_pts,
            note=(
                f"+8 for AbuseIPDB {threshold}-{full_band - 1}, "
                f"+15 at {full_band}+ or listed ASN; applied once"
            ),
        )
        score = max(0, min(100, int(assessment.score) + route_pts))
        risk_level = self._calculate_risk_level(score, assessment.domain_age_flag)
        risk_color = risk_color_for_level(risk_level)
        return SecurityAssessment(
            score=score,
            risk_level=risk_level,
            risk_color=risk_color,
            factors=factors,
            domain_age_flag=assessment.domain_age_flag,
            components=components,
        )

    def apply_extra_factors(
        self,
        assessment: SecurityAssessment,
        category: str,
        findings: Optional[List[Dict[str, Any]]],
    ) -> SecurityAssessment:
        """Fold ``[{label, points, note}]`` findings into the score under one category."""
        if assessment is None or not findings:
            return assessment
        factors = list(assessment.factors or [])
        components = list(assessment.components or [])
        added = 0
        for f in findings:
            pts = int(f.get("points") or 0)
            self._record_factor(
                factors, components, category, str(f.get("label") or ""), pts,
                note=str(f.get("note") or ""),
            )
            added += pts
        score = max(0, min(100, int(assessment.score) + added))
        risk_level = self._calculate_risk_level(score, assessment.domain_age_flag)
        return SecurityAssessment(
            score=score,
            risk_level=risk_level,
            risk_color=risk_color_for_level(risk_level),
            factors=factors,
            domain_age_flag=assessment.domain_age_flag,
            components=components,
        )

    def apply_deep_findings(
        self,
        assessment: SecurityAssessment,
        *,
        link_points: int = 0,
        link_factors: Optional[List[str]] = None,
        body_points: int = 0,
        body_factors: Optional[List[str]] = None,
        attachment_results: Optional[List[AttachmentVerdict]] = None,
        beacon_count: int = 0,
    ) -> SecurityAssessment:
        """Fold Deep Scan link/body/attachment/beacon findings into the risk score."""
        factors = list(assessment.factors or [])
        components = list(assessment.components or [])
        score = int(assessment.score)

        if link_points > 0:
            note = "; ".join(link_factors or []) or "Suspicious URLs in message body"
            self._record_factor(
                factors, components, "Deep / links", note, int(link_points),
                note="Deep Scan link heuristics",
            )
            score += int(link_points)

        if body_points > 0:
            note = "; ".join(body_factors or []) or "Body phishing heuristics"
            self._record_factor(
                factors, components, "Deep / body", note, int(body_points),
                note="Deep Scan body heuristics",
            )
            score += int(body_points)

        if attachment_results:
            highs = sum(1 for a in attachment_results if not a.ok and a.risk_level == "high")
            meds = sum(1 for a in attachment_results if not a.ok and a.risk_level == "medium")
            if highs:
                add = min(20, 8 * highs)
                self._record_factor(
                    factors, components, "Deep / attachments",
                    f"High-risk attachment(s) ({highs})", add,
                    note="Deep Scan attachment findings folded into score",
                )
                score += add
            elif meds:
                add = min(10, 4 * meds)
                self._record_factor(
                    factors, components, "Deep / attachments",
                    f"Medium-risk attachment(s) ({meds})", add,
                    note="Deep Scan attachment findings folded into score",
                )
                score += add

        if beacon_count and beacon_count > 0:
            add = min(10, 2 * int(beacon_count))
            self._record_factor(
                factors, components, "Deep / beacons",
                f"Tracking beacon(s) detected ({beacon_count})", add,
                note="Deep Scan beacon findings folded into score",
            )
            score += add

        score = max(0, min(100, score))
        risk_level = self._calculate_risk_level(score, assessment.domain_age_flag)
        risk_color = risk_color_for_level(risk_level)
        return SecurityAssessment(
            score=score,
            risk_level=risk_level,
            risk_color=risk_color,
            factors=factors,
            domain_age_flag=assessment.domain_age_flag,
            components=components,
        )

    def apply_high_risk_elements(
        self,
        assessment: SecurityAssessment,
        *,
        link_findings: Optional[List[Any]] = None,
        beacon_count: int = 0,
        attachment_not_ok: int = 0,
    ) -> SecurityAssessment:
        """Fold Basic-scan link / beacon / attachment findings into the score.

        Links: +2 per suspicious (medium), +10 per high-risk.
        Beacons and bad attachments: +2 each.
        """
        if assessment is None:
            return assessment

        def level(item: Any) -> str:
            value = (
                getattr(item, "risk_level", None)
                if not isinstance(item, dict)
                else item.get("risk_level")
            )
            return str(value or "").lower()

        suspicious = sum(1 for f in (link_findings or []) if level(f) == "medium")
        high_links = sum(1 for f in (link_findings or []) if level(f) == "high")
        beacons = int(beacon_count or 0)
        bad_atts = int(attachment_not_ok or 0)

        add = (
            SUSPICIOUS_LINK_POINTS * suspicious
            + HIGH_RISK_LINK_POINTS * high_links
            + HIGH_RISK_ELEMENT_POINTS * (beacons + bad_atts)
        )
        if add <= 0:
            return assessment

        factors = list(assessment.factors or [])
        components = list(assessment.components or [])
        parts: List[str] = []
        if suspicious:
            parts.append(
                f"{suspicious} suspicious link(s) (+{SUSPICIOUS_LINK_POINTS} each)"
            )
        if high_links:
            parts.append(
                f"{high_links} high-risk link(s) (+{HIGH_RISK_LINK_POINTS} each)"
            )
        if beacons:
            parts.append(f"{beacons} beacon(s) (+{HIGH_RISK_ELEMENT_POINTS} each)")
        if bad_atts:
            parts.append(
                f"{bad_atts} bad attachment(s) (+{HIGH_RISK_ELEMENT_POINTS} each)"
            )
        self._record_factor(
            factors,
            components,
            "High-risk elements",
            "; ".join(parts),
            add,
            note="Suspicious links +2; high-risk links +10; beacons/attachments +2",
        )
        score = max(0, min(100, int(assessment.score) + add))
        risk_level = self._calculate_risk_level(score, assessment.domain_age_flag)
        risk_color = risk_color_for_level(risk_level)
        return SecurityAssessment(
            score=score,
            risk_level=risk_level,
            risk_color=risk_color,
            factors=factors,
            domain_age_flag=assessment.domain_age_flag,
            components=components,
        )

    def apply_trusted_sender_discount(
        self, assessment: SecurityAssessment
    ) -> SecurityAssessment:
        """Cut the score for a sender the user has explicitly trusted.

        Runs last so the reduction lands on the fully accumulated score and
        the risk level is derived from the reduced figure.
        """
        if assessment is None:
            return assessment
        original = int(assessment.score)
        if original <= 0:
            return assessment

        score = max(0, min(100, int(round(original * (1.0 - TRUSTED_SENDER_DISCOUNT)))))
        factors = list(assessment.factors or [])
        components = list(assessment.components or [])
        self._record_factor(
            factors,
            components,
            "Sender trust",
            "Trusted sender",
            score - original,
            note=f"{int(round(TRUSTED_SENDER_DISCOUNT * 100))}% reduction for a trusted sender",
        )
        # A very new sending domain still forces HIGH; trust does not clear it.
        risk_level = self._calculate_risk_level(score, assessment.domain_age_flag)
        risk_color = risk_color_for_level(risk_level)
        return SecurityAssessment(
            score=score,
            risk_level=risk_level,
            risk_color=risk_color,
            factors=factors,
            domain_age_flag=assessment.domain_age_flag,
            components=components,
        )

    def apply_score_smoothing(
        self, assessment: SecurityAssessment, sender_email: Optional[str]
    ) -> SecurityAssessment:
        """Damp scan-to-scan jitter by blending with the sender's history.

        Runs last, after every other adjustment, so the smoothed figure is
        what renders. A very new domain still forces HIGH, and raw scores at
        or above the bypass threshold are never softened.
        """
        if assessment is None:
            return assessment
        # Forced-high signals (blocklist, brand-new domain) must not be damped.
        if assessment.domain_age_flag:
            update_sender_score_ema(sender_email, int(assessment.score))
            return assessment

        raw = int(assessment.score)
        smoothed = update_sender_score_ema(sender_email, raw)
        if smoothed is None or smoothed == raw:
            return assessment

        factors = list(assessment.factors or [])
        components = list(assessment.components or [])
        self._record_factor(
            factors,
            components,
            "Score smoothing",
            "Blended with sender history",
            smoothed - raw,
            note=f"Raw {raw} smoothed to {smoothed} to damp scan-to-scan jitter",
        )
        risk_level = self._calculate_risk_level(smoothed, assessment.domain_age_flag)
        risk_color = risk_color_for_level(risk_level)
        return SecurityAssessment(
            score=smoothed,
            risk_level=risk_level,
            risk_color=risk_color,
            factors=factors,
            domain_age_flag=assessment.domain_age_flag,
            components=components,
        )

    def _assess_authentication(
        self,
        auth_info: AuthenticationInfo,
        score: int,
        factors: List[str],
        components: List[RiskScoreComponent],
    ) -> int:
        """Assess authentication methods (reversed scoring: higher is worse)."""
        cat = "Authentication"
        if auth_info.spf:
            if "pass" in auth_info.spf.lower():
                self._record_factor(factors, components, cat, "SPF Validated", 0)
                score += 0
            elif "fail" in auth_info.spf.lower():
                self._record_factor(
                    factors, components, cat, f"SPF {auth_info.spf.upper()}", 15
                )
                score += 15
            else:
                self._record_factor(
                    factors, components, cat, f"SPF {auth_info.spf.upper()}", 7
                )
                score += 7
        else:
            self._record_factor(factors, components, cat, "SPF Missing/Problem", 7)
            score += 7
        
        if auth_info.dkim:
            if "pass" in auth_info.dkim.lower():
                self._record_factor(factors, components, cat, "DKIM Validated", 0)
                score += 0
            elif "fail" in auth_info.dkim.lower():
                self._record_factor(
                    factors, components, cat, f"DKIM {auth_info.dkim.upper()}", 15
                )
                score += 15
            else:
                self._record_factor(
                    factors, components, cat, f"DKIM {auth_info.dkim.upper()}", 7
                )
                score += 7
        else:
            self._record_factor(factors, components, cat, "DKIM Missing/Problem", 7)
            score += 7
        
        if auth_info.dmarc:
            if "pass" in auth_info.dmarc.lower():
                self._record_factor(factors, components, cat, "DMARC Compliant", 0)
                score += 0
            elif "fail" in auth_info.dmarc.lower():
                self._record_factor(
                    factors, components, cat, f"DMARC {auth_info.dmarc.upper()}", 20
                )
                score += 20
            else:
                self._record_factor(
                    factors, components, cat, f"DMARC {auth_info.dmarc.upper()}", 10
                )
                score += 10
        else:
            self._record_factor(factors, components, cat, "DMARC Missing/Problem", 10)
            score += 10

        # CompAuth / ARC — related header signals (display + light scoring).
        if auth_info.compauth and auth_info.compauth.get("result") not in {
            None,
            "",
            "missing",
            "not_present",
        }:
            comp_res = str(auth_info.compauth.get("result") or "").lower()
            reason = auth_info.compauth.get("reason_code", "")
            label = f"CompAuth {comp_res.upper()}"
            if reason:
                label = f"{label} (reason {reason})"
            if "pass" in comp_res:
                self._record_factor(factors, components, cat, label, 0)
            elif "fail" in comp_res:
                self._record_factor(factors, components, cat, label, 10)
                score += 10
            else:
                self._record_factor(factors, components, cat, label, 5)
                score += 5
        else:
            self._record_factor(
                factors, components, cat, "CompAuth Not Applicable / Missing", 0
            )

        arc_res = str((auth_info.arc or {}).get("result") or "none").lower()
        if arc_res in {"", "none", "missing", "not_present"}:
            self._record_factor(factors, components, cat, "ARC Not Present", 0)
        elif "pass" in arc_res:
            self._record_factor(factors, components, cat, "ARC Pass", 0)
        elif "fail" in arc_res:
            self._record_factor(factors, components, cat, f"ARC {arc_res.upper()}", 8)
            score += 8
        else:
            self._record_factor(factors, components, cat, f"ARC {arc_res.upper()}", 3)
            score += 3
        
        return score
    
    def _assess_geolocation(
        self,
        geo_results: List[GeoLocationResult],
        score: int,
        factors: List[str],
        components: List[RiskScoreComponent],
    ) -> int:
        """Assess geolocation and security flags (reversed scoring: higher is worse)."""
        if not geo_results:
            return score
        
        origin_geo = geo_results[0]
        cat = "Origin / network"
        whitelisted = bool(origin_geo.security_data.get("Whitelisted") is True)
        geo_penalty = 0
        
        if origin_geo.success and origin_geo.blocklist and "Listed: " in origin_geo.blocklist:
            blocklists = origin_geo.blocklist.replace('Listed: ', '')
            self._record_factor(
                factors,
                components,
                cat,
                f"Origin IP Blocklisted ({blocklists})",
                100,
                note="Forces score to 100",
            )
            score = 100
            return score
        elif origin_geo.success:
            self._record_factor(
                factors, components, cat, "Origin IP Not Blocklisted", 0
            )
        
        if origin_geo.success:
            threat_active = False
            no_data_count = 0

            for flag_name in SECURITY_FLAG_ORDER:
                is_active = origin_geo.security_data.get(flag_name)
                if flag_name in SECURITY_NO_DATA_FLAGS and is_active is None:
                    no_data_count += 1

                if not is_active:
                    continue

                label = f"Security Flag: {flag_name} Active"
                if flag_name == "VPN" and origin_geo.vpn_service:
                    label = f"Security Flag: VPN ({origin_geo.vpn_service})"
                elif flag_name == "Hosting":
                    label = "Security Flag: Hosting / Datacenter Active"
                elif flag_name == "Whitelisted":
                    label = "Security Flag: Whitelisted (dampens soft signals)"

                note = ""
                if flag_name in SECURITY_THREAT_FLAGS:
                    note = "Counted in threat-flag penalty below"
                    threat_active = True
                elif flag_name == "Hosting":
                    note = "+5 hosting / datacenter"
                elif flag_name == "CDN":
                    note = "+3 CDN heuristic"
                elif flag_name == "Mobile":
                    note = "+2 mobile ISP (weaker signal)"
                elif flag_name == "Whitelisted":
                    note = "Reduces Hosting/CDN/Mobile/abuse penalties"

                self._record_factor(
                    factors, components, cat, label, 0, note=note
                )

            if threat_active:
                pts = 10
                self._record_factor(
                    factors,
                    components,
                    cat,
                    "One or more threat flags active (Tor/Proxy/Anon/VPN/Relay)",
                    pts,
                )
                score += pts
                geo_penalty += pts

            if origin_geo.security_data.get("Hosting") is True:
                pts = 0 if whitelisted else 5
                self._record_factor(
                    factors,
                    components,
                    cat,
                    "Hosting / Datacenter origin",
                    pts,
                    note="Damped by whitelist" if whitelisted else "",
                )
                score += pts
                geo_penalty += pts

            if origin_geo.security_data.get("CDN") is True:
                pts = 0 if whitelisted else 3
                self._record_factor(
                    factors,
                    components,
                    cat,
                    "CDN / edge network origin",
                    pts,
                    note="Damped by whitelist" if whitelisted else "",
                )
                score += pts
                geo_penalty += pts

            if origin_geo.security_data.get("Mobile") is True:
                pts = 0 if whitelisted else 2
                self._record_factor(
                    factors,
                    components,
                    cat,
                    "Mobile ISP origin",
                    pts,
                    note="Damped by whitelist" if whitelisted else "",
                )
                score += pts
                geo_penalty += pts

            # Origin AbuseIPDB confidence (same threshold as hop/route risk).
            route_cfg = load_route_risk_config()
            threshold = int(route_cfg.get("abuse_threshold", 25) or 25)
            abuse_score = origin_geo.abuse_confidence
            if abuse_score is not None:
                if abuse_score >= threshold:
                    # Graded band instead of a hard cliff: confidence values
                    # hovering around the threshold no longer flip +15 on/off.
                    full_band = threshold + 15
                    base_pts = 15 if abuse_score >= full_band else 8
                    pts = 0 if whitelisted else base_pts
                    self._record_factor(
                        factors,
                        components,
                        cat,
                        f"Origin AbuseIPDB confidence {abuse_score} >= {threshold}",
                        pts,
                        note=(
                            "Damped by whitelist"
                            if whitelisted
                            else f"+8 in {threshold}-{full_band - 1}, +15 at {full_band}+"
                        ),
                    )
                    score += pts
                    geo_penalty += pts
                else:
                    self._record_factor(
                        factors,
                        components,
                        cat,
                        f"Origin AbuseIPDB confidence {abuse_score} < {threshold}",
                        0,
                    )

            if whitelisted and geo_penalty > 0:
                # Extra dampening: cut half of remaining soft geo penalties (already
                # zeroed Hosting/CDN/Mobile/abuse above; this trims threat +10 too).
                damp = min(10, geo_penalty // 2)
                if damp > 0:
                    self._record_factor(
                        factors,
                        components,
                        cat,
                        "Whitelisted dampening",
                        -damp,
                        note="Reduces false-positive soft risk",
                    )
                    score = max(0, score - damp)

            if no_data_count > 0:
                # Missing reputation data is usually an API timeout, not a
                # threat signal; cap the penalty so lookup hiccups cannot
                # swing the score between scans.
                pts = min(2, no_data_count)
                self._record_factor(
                    factors,
                    components,
                    cat,
                    f"Security flag data unavailable ({no_data_count})",
                    pts,
                    note="Capped at +2 total; transient lookup gaps stay quiet",
                )
                score += pts
        
        return score
    
    def _assess_domain_age(
        self,
        whois_info: Dict[str, Any],
        score: int,
        factors: List[str],
        components: List[RiskScoreComponent],
    ) -> Tuple[int, Optional[str]]:
        """Assess domain age (reversed scoring: higher is worse)."""
        domain_age_flag = None
        cat = "Domain age"
        
        if not whois_info or "error" in whois_info:
            if whois_info and whois_info.get("error"):
                # WHOIS failures are almost always rate limits or timeouts;
                # recording them as risk made the score flap between scans.
                self._record_factor(
                    factors,
                    components,
                    cat,
                    "Sender Domain WHOIS Unavailable",
                    0,
                    note="Informational; transient lookup failures add no risk",
                )
            return score, domain_age_flag
        
        creation_date = whois_info.get("creation_date", "Unknown")
        if creation_date == "Unknown":
            return score, domain_age_flag
        
        try:
            parsed_date = self._parse_date(creation_date)
            if not parsed_date:
                self._record_factor(
                    factors, components, cat, "Domain Creation Date Invalid", 10
                )
                score += 10
                return score, domain_age_flag
            
            parsed_date = pytz.utc.localize(parsed_date)
            now = datetime.now(pytz.utc)
            age_days = (now - parsed_date).days

            if age_days < 0:
                self._record_factor(
                    factors, components, cat, "Domain Creation Date Invalid", 10
                )
                score += 10
            elif age_days < 10:
                self._record_factor(
                    factors,
                    components,
                    cat,
                    f"Domain Extremely New ({age_days} days)",
                    100,
                    note="Forces score to 100 and risk HIGH",
                )
                domain_age_flag = "Domain Less Than 10 Days Old - CAUTION"
                score = 100
            elif age_days < 30:
                self._record_factor(
                    factors,
                    components,
                    cat,
                    f"Domain Very New ({age_days} days)",
                    50,
                )
                score += 50
            elif age_days < 100:
                self._record_factor(
                    factors, components, cat, f"Domain New ({age_days} days)", 30
                )
                score += 30
            elif age_days < 365:
                self._record_factor(
                    factors,
                    components,
                    cat,
                    f"Domain Moderately New ({age_days} days)",
                    10,
                )
                score += 10
            else:
                self._record_factor(
                    factors,
                    components,
                    cat,
                    f"Domain Established ({age_days} days)",
                    0,
                )
                score += 0
                
        except Exception as e:
            self.logger.error(f"Error processing domain age: {e}")
            self._record_factor(
                factors, components, cat, "Domain Age Processing Error", 5
            )
            score += 5
        
        return score, domain_age_flag
    
    def _parse_date(self, date_str: str) -> Optional[datetime]:
        """Parse date string to datetime."""
        date_formats = ["%Y-%m-%d", "%Y-%m-%dT%H:%M:%SZ", "%Y-%m-%d %H:%M:%S"]
        date_str = date_str.split(' ')[0]
        
        for fmt in date_formats:
            try:
                return datetime.strptime(date_str, fmt)
            except ValueError:
                continue
        
        return None
    
    def _calculate_risk_level(self, score: int, domain_age_flag: Optional[str]) -> str:
        """Calculate risk level based on reversed score (higher is worse)."""
        return risk_level_for_score(score, domain_age_flag)

_RDAP_CACHE: Dict[str, Dict[str, str]] = {}


class HTMLReportGenerator:
    """Generates HTML security reports."""
    
    def __init__(self):
        self.logger = logging.getLogger(__name__)
    
    def generate_report(
        self,
        sender_email: Optional[str],
        sender_domain: Optional[str],
        sender_ip: Optional[str],
        sender_hostname: Optional[str],
        geo_result: GeoLocationResult,
        whois_info: Dict[str, Any],
        auth_info: AuthenticationInfo,
        security_assessment: SecurityAssessment,
        headers: Optional[str] = None,
        guri_db: Optional['GURIDatabase'] = None,
        metadata: Optional[Dict[str, str]] = None,
        footer_mode: str = "compact",
    ) -> str:
        """Generate HTML security report. deep mode returns report_url=... (not footer HTML)."""
        footer_mode = (footer_mode or "compact").lower()
        deep_mode = footer_mode == "deep"
        display_data = self._prepare_display_data(
            sender_email, sender_domain, sender_ip,
            sender_hostname, geo_result, whois_info,
            auth_info, security_assessment,
            headers, guri_db, metadata or {},
            deep_mode=deep_mode,
            footer_mode=footer_mode,
        )
        if deep_mode:
            report_url = self._write_deep_scan_report_file(display_data)
            return f"report_url={report_url}\n" if report_url else "report_url=\n"
        return self._build_html(display_data, footer_mode)
    
    def _prepare_display_data(
        self,
        sender_email: Optional[str],
        sender_domain: Optional[str],
        sender_ip: Optional[str],
        sender_hostname: Optional[str],
        geo_result: GeoLocationResult,
        whois_info: Dict[str, Any],
        auth_info: AuthenticationInfo,
        security_assessment: SecurityAssessment,
        headers: Optional[str] = None,
        guri_db: Optional['GURIDatabase'] = None,
        metadata: Optional[Dict[str, str]] = None,
        deep_mode: bool = False,
        footer_mode: str = "compact",
    ) -> Dict[str, Any]:
        """Prepare all data for display."""
        metadata = metadata or {}
        footer_mode = (footer_mode or "compact").lower()
        # Compact auto-scans skip the expensive hop/CRT/AV work so Outlook's
        # machine stays responsive when mail arrives; full/deep keep full fidelity.
        lite = footer_mode == "compact" and not deep_mode
        tz = pytz.utc
        analysis_time = datetime.now(tz).strftime('%Y-%m-%d %H:%M BST')
        ip_to_display = sender_ip if sender_ip else geo_result.ip
        hostname_to_display = sender_hostname if sender_ip else None
        ip_display = self._format_ip_display(ip_to_display, hostname_to_display)
        location_display = self._format_location(geo_result, sender_hostname)
        auth_display = self._format_authentication(auth_info)
        security_flags = self._format_security_flags(geo_result)
        org_display = geo_result.org
        if org_display and isinstance(org_display, str):
            org_display = re.sub(r"AS\d+\s*", "", org_display).strip()
        if not geo_result.success:
            org_display = "Unknown"
        def _auth_type_string() -> str:
            passed = []
            for meth, val in [("SPF", auth_info.spf), ("DKIM", auth_info.dkim), ("DMARC", auth_info.dmarc)]:
                if val and val.lower() == "pass":
                    passed.append(meth)
            if len(passed) == 3:
                return "SPF + DKIM + DMARC (All Passed)"
            if passed:
                return " + ".join(passed) + " (Partial)"
            return "None"
        auth_type_str = _auth_type_string()
        # Calculate hops and ToT
        hops, tot = self._calculate_hops_and_tot(headers)
        # Routing info: hostnames and ASNs traversed
        routing_info = self._extract_routing_info(headers)
        per_hop_analysis = self._extract_hop_details(headers, lite=lite)

        # Route/ASN risk needs hop AbuseIPDB scores; apply after hop enrichment.
        security_assessment = SecurityAssessor().apply_route_path_risk(
            security_assessment, per_hop_analysis
        )
        if time_checks is not None and headers:
            m_date = re.search(r'^Date:\s*(.+)$', headers, re.IGNORECASE | re.MULTILINE)
            try:
                time_findings = time_checks.check_time_consistency(
                    per_hop_analysis,
                    date_header=m_date.group(1) if m_date else None,
                    origin_country=geo_result.country if geo_result.success else None,
                )
            except Exception as exc:
                self.logger.warning("Timestamp consistency check failed: %s", exc)
                time_findings = []
            security_assessment = SecurityAssessor().apply_extra_factors(
                security_assessment, "Timestamps", time_findings
            )
        if route_checks is not None and per_hop_analysis:
            try:
                route_findings = route_checks.check_routing(per_hop_analysis, lite=lite)
            except Exception as exc:
                self.logger.warning("Routing check failed: %s", exc)
                route_findings = []
            security_assessment = SecurityAssessor().apply_extra_factors(
                security_assessment, "Routing / BGP", route_findings
            )
            # BGP network-info often has a better prefix than RDAP for M365 IPv6.
            for hop in per_hop_analysis:
                bgp = hop.get("bgp") or {}
                prefix = str(bgp.get("prefix") or "").strip()
                if prefix and "/" in prefix and self._blank_reg(hop.get("prefix")):
                    hop["prefix"] = prefix
                if self._blank_reg(hop.get("asn")) and bgp.get("origins"):
                    hop["asn"] = f"AS{bgp['origins'][0]}"

        broker_hit: Optional[Dict[str, Any]] = None
        if broker_match_sender is not None:
            try:
                broker_hit = broker_match_sender(sender_email, sender_domain)
            except Exception as exc:
                self.logger.debug("Broker match failed: %s", exc)
                broker_hit = None
        if broker_hit:
            security_assessment = SecurityAssessor().apply_extra_factors(
                security_assessment,
                "Data broker",
                [
                    {
                        "label": f"Sender matches data broker: {broker_hit.get('name')}",
                        "points": 8,
                        "note": "People-search / data-broker domain — consider a removal "
                        "request (GURI → Data Brokers, or Queue data removal below)",
                    }
                ],
            )
            self._auto_queue_broker(broker_hit, sender_email, sender_domain, headers)

        if not sender_domain or sender_domain == "Unknown":
            sender_domain = self._infer_sender_domain(sender_email, sender_domain, per_hop_analysis, headers)

        if sender_domain and sender_domain != "Unknown":
            needs_whois = (
                not whois_info
                or whois_info.get("error")
                or whois_info.get("registrar") in (None, "Unknown")
            )
            if needs_whois:
                whois_info = WhoisService().get_whois_info(sender_domain)

        whois_display = self._format_whois(whois_info)

        cert_info: Dict[str, Any] = {}
        if not lite and sender_domain and sender_domain != "Unknown":
            cert_info = WhoisService().get_cert_history(sender_domain)
        cert_display = self._format_cert_history(cert_info)
        sender_local_html, sender_local_plain = self._format_sender_local_time(
            headers, geo_result.timezone
        )

        guri = self._generate_or_retrieve_guri(
            guri_db, sender_email, sender_domain, security_assessment, headers, metadata
        )
        attachment_scan_results = self._scan_exported_attachments(
            metadata, deep_mode=deep_mode, skip_av=lite
        )
        attachment_count, attachment_ok, attachment_not_ok = self._summarize_attachment_results(
            attachment_scan_results, metadata
        )
        beacon_count, beacon_blocked, beacon_urls = self._parse_beacon_metadata(metadata)

        link_findings: List[Any] = []
        body_scan: Any = None
        body_available = False
        # Compact exports a capped body sidecar; full/deep export the full body.
        # Link heuristics run whenever a body is present (deep risk fusion stays deep-only).
        html_body, text_body = self._load_deep_body_payload(metadata)
        body_available = bool(html_body or text_body)
        if LinkScanner is not None and body_available:
            link_findings = LinkScanner().scan(
                html_body=html_body,
                text_body=text_body,
                sender_domain=sender_domain,
            )

        intel_reports: List[Any] = []
        risktable_checked = False
        if threat_intel is not None:
            try:
                intel_reports = self._run_threat_intel(
                    "deep" if deep_mode else ("compact" if lite else "full"),
                    per_hop_analysis=per_hop_analysis,
                    sender_ip=sender_ip or geo_result.ip,
                    sender_domain=sender_domain,
                    headers=headers or "",
                    link_findings=link_findings,
                    attachment_results=attachment_scan_results,
                )
                risktable_checked = True
                security_assessment = SecurityAssessor().apply_extra_factors(
                    security_assessment,
                    "Threat intel",
                    threat_intel.score_reports(intel_reports),
                )
            except Exception as exc:
                self.logger.warning("Threat intel lookups failed: %s", exc)

        if deep_mode:
            if BodyScanner is not None:
                body_scan = BodyScanner().scan(
                    html_body=html_body,
                    text_body=text_body,
                    headers=headers or "",
                    sender_email=sender_email,
                    sender_domain=sender_domain,
                )
            link_points, link_factors = (0, [])
            body_points, body_factors = (0, [])
            if LinkScanner is not None and link_findings:
                link_points, link_factors = LinkScanner().risk_points(link_findings)
            if BodyScanner is not None and body_scan is not None:
                body_points, body_factors = BodyScanner().risk_points(body_scan)
            security_assessment = SecurityAssessor().apply_deep_findings(
                security_assessment,
                link_points=link_points,
                link_factors=link_factors,
                body_points=body_points,
                body_factors=body_factors,
                attachment_results=attachment_scan_results,
                beacon_count=beacon_count,
            )

        # Fold Basic-scan findings into the score (Deep mode uses its own weights).
        if not deep_mode:
            security_assessment = SecurityAssessor().apply_high_risk_elements(
                security_assessment,
                link_findings=link_findings,
                beacon_count=beacon_count,
                attachment_not_ok=attachment_not_ok,
            )

        # Sender status and its risk discount must settle before anything
        # renders the score, so every consumer below sees the same figure.
        block_state = self._sender_block_state(sender_email, sender_domain)
        sender_seen_count = record_sender_email(
            sender_email, self._extract_message_id(headers)
        )
        if block_state.get("trusted"):
            security_assessment = SecurityAssessor().apply_trusted_sender_discount(
                security_assessment
            )

        # Final step: damp scan-to-scan jitter against this sender's history
        # so the rendered score, level and colour all use the smoothed figure.
        security_assessment = SecurityAssessor().apply_score_smoothing(
            security_assessment, sender_email
        )

        attachments_html, attachment_report_url = self._create_attachment_assets(
            attachment_count,
            attachment_ok,
            attachment_not_ok,
            attachment_scan_results,
            metadata,
            guri,
            sender_email,
            sender_domain,
            security_assessment,
        )
        beacon_button_html, beacon_launch_url = self._create_beacon_assets(
            beacon_count, beacon_urls, metadata, guri, sender_email, sender_domain,
            beacon_blocked,
        )
        links_summary_line, links_detail_html = self._format_link_summary(
            link_findings, body_available
        )
        links_report_url = self._write_links_report_file(
            link_findings,
            body_available,
            metadata,
            guri,
            sender_email,
            sender_domain,
        )
        links_summary_html = links_summary_line
        if links_summary_line and links_report_url:
            links_summary_html = self._metric_report_link(
                links_report_url, links_summary_line
            )
        action_buttons_html = self._build_action_buttons_html(
            sender_email,
            sender_domain,
            guri,
            block_state,
            links_report_url,
            broker_hit=broker_hit,
            subject=str(
                (metadata or {}).get("original subject")
                or (metadata or {}).get("\ufefforiginal subject")
                or ""
            ),
        )
        att_quarantined = self._parse_quarantined_count(metadata)
        top_mail_banners = self._build_top_of_mail_banners(
            block_state=block_state,
            attachment_count=attachment_count,
            beacon_count=beacon_count,
            beacon_blocked=beacon_blocked,
            attachment_not_ok=attachment_not_ok,
            attachments_quarantined=att_quarantined,
            link_findings=link_findings,
            risk_score=int(security_assessment.score or 0),
            risk_level=str(security_assessment.risk_level or ""),
            sender_seen_count=sender_seen_count,
            body_available=body_available,
            risktable_checked=risktable_checked,
        )
        # Keep a short strip above the footer summary as well.
        sender_blocks_banner = top_mail_banners
        # Military DTG: YYYYMMDDHHmmz (UTC)
        created_zulu = datetime.now(pytz.utc).strftime('%Y%m%d%H%M') + 'z'

        sender_ip_plain = html.escape(sender_ip or "Unknown")
        sender_host_plain = html.escape(sender_hostname.strip()) if sender_hostname and sender_hostname.strip() else ""
        if sender_host_plain and sender_host_plain.lower() == sender_ip_plain.lower():
            sender_host_plain = ""

        # Basic Scan footers always use Basic label (Deep is report-only, never emailed).
        summary_line1_html = (
            f"Aliniant AES Scan Result - Basic Scan | "
            f"{{{{AES_RISK}}}} | "
            f"{attachments_html} | {beacon_button_html}"
        )
        if links_summary_html:
            summary_line1_html = f"{summary_line1_html} | {links_summary_html}"
        sender_ip_part = f"Sender IP {sender_ip_plain}"
        if sender_host_plain:
            sender_ip_part = f"{sender_ip_part} {sender_host_plain}"
        summary_line2_prefix = (
            f"{sender_ip_part} | {location_display} | Hops:{hops} | "
            f"Sent: {sender_local_plain} | "
            f"Created: {created_zulu} | "
        )
        # Plain-text fallback for logs / non-HTML consumers.
        summary_line2 = summary_line2_prefix + "Show Full Scan"
        # Placeholder replaced in _build_html with a link to the already-built full report.
        summary_line2_html = html.escape(summary_line2_prefix) + "{{AES_FULLSCAN}}"
        
        return {
            "analysis_time": analysis_time,
            "risk_level": security_assessment.risk_level,
            "risk_score": security_assessment.score,
            "risk_color": security_assessment.risk_color,
            "domain_age_flag": security_assessment.domain_age_flag,
            "sender_email": self._format_email(sender_email),
            "sender_domain": sender_domain or "Unknown",
            "ip_display": ip_display,
            "hops_and_tot": f"{hops} hops; ToT: {tot}",
            "routing_info": routing_info,
            "location": location_display,
            "organization": org_display,
            "asn": self._format_asn(geo_result.asn),
            "whois": whois_display,
            "cert_history": cert_display,
            "sender_local_time": sender_local_html,
            "action_buttons_html": action_buttons_html,
            "block_state": block_state,
            "sender_blocks_banner": sender_blocks_banner,
            "top_mail_banners": top_mail_banners,
            "links_detail_html": links_detail_html,
            "links_report_url": links_report_url,
            "auth": auth_display,
            "security_flags": security_flags,
            "auth_type": auth_type_str,
            "factors": security_assessment.factors,
            "risk_components": security_assessment.components,
            "per_hop_analysis": per_hop_analysis,
            "guri": guri,
            "attachment_count": attachment_count,
            "attachment_ok": attachment_ok,
            "attachment_not_ok": attachment_not_ok,
            "attachment_scan_results": attachment_scan_results,
            "attachments_html": attachments_html,
            "attachment_report_url": attachment_report_url,
            "beacon_count": beacon_count,
            "beacon_button_html": beacon_button_html,
            "beacon_launch_url": beacon_launch_url,
            "created_zulu": created_zulu,
            "summary_line1_html": summary_line1_html,
            "summary_line2": summary_line2,
            "summary_line2_html": summary_line2_html,
            "deep_mode": deep_mode,
            "link_findings": [
                f.to_dict() if hasattr(f, "to_dict") else f for f in (link_findings or [])
            ],
            "body_scan": body_scan.to_dict() if body_scan is not None and hasattr(body_scan, "to_dict") else {},
            "threat_intel": [r.to_dict() for r in intel_reports],
            "broker_hit": broker_hit,
            "risktable_disable_urls": self._risktable_disable_urls(link_findings),
            "origin_lat": geo_result.lat,
            "origin_lon": geo_result.lon,
        }
    
    def _calculate_hops_and_tot(self, headers: Optional[str]) -> tuple:
        if not headers:
            return ("N/A", "N/A")
        received_lines = HeaderParser.extract_received_blocks(headers)
        hops = len(received_lines)
        times = []
        for line in received_lines:
            m = re.search(r';\s*(.+)$', line)
            if m:
                try:
                    dt = email.utils.parsedate_to_datetime(m.group(1))
                    # Headers with "-0000" parse as naive; normalize so naive and
                    # aware values can be compared without a TypeError.
                    if dt.tzinfo is None or dt.tzinfo.utcoffset(dt) is None:
                        dt = pytz.utc.localize(dt)
                    times.append(dt)
                except Exception:
                    continue
        if len(times) >= 2:
            tot_seconds = abs((max(times) - min(times)).total_seconds())
            tot_str = f"{tot_seconds:.2f}s"
        else:
            tot_str = "N/A"
        return (hops, tot_str)
    
    def _format_ip_display(self, ip: Optional[str], hostname: Optional[str]) -> str:
        """Format IP address display, always appending terminal/hostname if available and not identical to IP."""
        if not ip:
            return "Unknown"
        ip_str = html.escape(ip)
        if hostname and hostname.strip() and hostname.lower() != ip.lower():
            return f"{ip_str} ({html.escape(hostname.strip())})"
        return ip_str
    
    def _format_location(self, geo_result: GeoLocationResult, hostname: Optional[str] = None) -> str:
        """Format location display; note when geo is a cloud outbound relay, not the author."""
        if not geo_result.success and geo_result.city:
            base = html.escape(geo_result.city)
        else:
            city = geo_result.city if geo_result.success else "Unknown"
            country = geo_result.country if geo_result.success else ""
            base = f"{city}, {country}".strip(", ")
            if not base or base == ",":
                base = "Unknown"
            else:
                base = html.escape(base)

        host = (hostname or "").lower()
        if any(
            token in host
            for token in (
                "outbound.protection.outlook.com",
                "mail.protection.outlook.com",
                ".google.com",
                "outbound-mail.sendgrid.net",
                ".amazonaws.com",
            )
        ):
            return (
                f"{base} "
                f"<span style='color:#888888;'>(outbound relay geo — not sender office)</span>"
            )
        return base

    def _format_whois(self, whois_info: Dict[str, Any]) -> Dict[str, str]:
        """Format WHOIS information."""
        if "error" in whois_info and whois_info["error"] != "Lookup not performed":
            return {
                "registrar": f"<span style='color: #888888;'>{html.escape(whois_info['error'])}</span>",
                "created": "N/A",
                "age": "N/A"
            }
        registrar = html.escape(str(whois_info.get("registrar", "Unknown")))
        created = html.escape(str(whois_info.get("creation_date", "Unknown")))
        # Ensure we always pass a string, never None
        creation_date_str = str(whois_info.get("creation_date", "Unknown") or "Unknown")
        age_days = self._calculate_domain_age(creation_date_str)
        age = f"{age_days} days" if age_days is not None else "N/A"
        return {"registrar": registrar, "created": created, "age": age}

    def _format_cert_history(self, cert_info: Dict[str, Any]) -> str:
        """Render crt.sh certificate-transparency history as a footer value."""
        if not cert_info or cert_info.get("error"):
            err = str((cert_info or {}).get("error") or "No data")
            return f"<span style='color: #888888;'>{html.escape(err)} (crt.sh)</span>"

        first = str(cert_info.get("first_seen") or "Unknown")
        age_days = int(cert_info.get("age_days") or 0)
        total = int(cert_info.get("total") or 0)
        recent = int(cert_info.get("recent_30d") or 0)
        active_only = bool(cert_info.get("active_only"))

        if active_only:
            # Fallback query only sees unexpired certs (max ~13 months back),
            # so a short history proves nothing about domain age here.
            colour = "#008000"
            label = f"Oldest active cert {html.escape(first)}"
            note = "large cert estate — full history timed out, unexpired certs only"
        elif age_days <= 30:
            colour = "#FF0000"
            label = f"First cert {html.escape(first)}"
            note = f"only {age_days} days of TLS history — brand-new infrastructure"
        elif age_days <= 180:
            colour = "#CC7A00"
            label = f"First cert {html.escape(first)}"
            note = f"{age_days} days of TLS history — young infrastructure"
        else:
            years = age_days / 365.25
            colour = "#008000"
            label = f"First cert {html.escape(first)}"
            note = f"{years:.1f} yrs of TLS history"

        parts = [
            f"<span style='color: {colour};'>{label} ({html.escape(note)})</span>",
            f"<span style='color: #888888;'>{total} certs logged</span>",
        ]
        if recent >= 25:
            parts.append(
                f"<span style='color: #CC7A00;'>{recent} new certs in last 30d — churn spike</span>"
            )
        return " | ".join(parts)

    def _format_sender_local_time(
        self, headers: Optional[str], geo_timezone: Optional[str]
    ) -> Tuple[str, str]:
        """(html, plain) sender wall-clock time at send, from the Date header.

        The Date header carries the sender's own clock and UTC offset, so the
        hour shown is what their watch said. 00:00-06:00 local sends are a
        classic compromised-account tell. If the header offset disagrees with
        the origin IP's timezone, that mismatch is flagged too.
        """
        unknown = ("<span style='color: #888888;'>Unknown</span>", "Unknown")
        if not headers:
            return unknown
        m = re.search(r'^Date:\s*(.+)$', headers, re.IGNORECASE | re.MULTILINE)
        if not m:
            return ("<span style='color: #888888;'>No Date header</span>", "Unknown")
        try:
            dt = email.utils.parsedate_to_datetime(m.group(1).strip())
        except Exception:
            dt = None
        if dt is None:
            return ("<span style='color: #888888;'>Unparseable Date header</span>", "Unknown")

        no_offset = dt.tzinfo is None or dt.tzinfo.utcoffset(dt) is None
        if no_offset:
            dt = pytz.utc.localize(dt)

        offset_txt = dt.strftime('%z')
        offset_disp = f"UTC{offset_txt[:3]}:{offset_txt[3:]}" if offset_txt else "UTC"
        clock = dt.strftime('%H:%M %a')
        # Military DTG is always Zulu (UTC); keep the sender's wall-clock for the details row.
        military = dt.astimezone(pytz.utc).strftime('%Y%m%d%H%M') + 'z'
        plain = military  # compact footer line 2

        notes: List[str] = []
        colour = "#008000"
        if no_offset:
            colour = "#888888"
            notes.append("no UTC offset in header — shown as UTC")
        elif 0 <= dt.hour < 6:
            colour = "#CC7A00"
            notes.append("sent 00:00–06:00 sender-local — unusual send time, classic compromise tell")

        if geo_timezone and not no_offset:
            try:
                tzobj = pytz.timezone(geo_timezone)
                expected = tzobj.utcoffset(dt.replace(tzinfo=None))
                actual = dt.utcoffset()
                if expected is not None and actual is not None:
                    diff_hours = abs((expected - actual).total_seconds()) / 3600.0
                    if diff_hours >= 2:
                        colour = "#CC7A00"
                        notes.append(
                            f"clock offset {offset_disp} disagrees with origin IP timezone "
                            f"{geo_timezone} by {diff_hours:.0f}h"
                        )
                    else:
                        notes.append(f"matches origin IP timezone {geo_timezone}")
            except Exception:
                pass

        html_out = f"<span style='color: {colour};'>{html.escape(military)}</span>"
        if notes or clock:
            detail = "; ".join([f"{clock} {offset_disp} sender-local"] + notes)
            html_out += (
                f" <span style='color: #888888;'>({html.escape(detail)})</span>"
            )
        return (html_out, plain)

    def _load_sender_rules(self) -> Dict[str, List[str]]:
        """Per-sender rules written by the aes:// footer buttons."""
        path = (
            Path(os.environ.get("LOCALAPPDATA", CONFIG["base_path"]))
            / "GeoFooter" / "aes_sender_rules.json"
        )
        try:
            if path.is_file():
                data = json.loads(path.read_text(encoding="utf-8"))
                if isinstance(data, dict):
                    return {
                        key: [str(v).strip().lower() for v in (data.get(key) or [])]
                        for key in ("block_attachments", "block_beacons", "trusted", "untrusted")
                    }
        except Exception as exc:
            self.logger.warning("Failed to read sender rules: %s", exc)
        return {"block_attachments": [], "block_beacons": [], "trusted": [], "untrusted": []}

    def _sender_block_state(
        self, sender_email: Optional[str], sender_domain: Optional[str]
    ) -> Dict[str, bool]:
        rules = self._load_sender_rules()
        sender = (sender_email or "").strip().lower()
        domain = (sender_domain or "").strip().lower()

        def on_list(key: str) -> bool:
            entries = rules.get(key) or []
            if sender and sender in entries:
                return True
            return bool(domain and domain != "unknown" and domain in entries)

        trusted = on_list("trusted")
        # Trusted always wins for the footer buttons: a sender on the trust list
        # must not still show Beacons/Attachments BLOCKED after Trust was clicked.
        return {
            "attachments": False if trusted else on_list("block_attachments"),
            "beacons": False if trusted else on_list("block_beacons"),
            "trusted": trusted,
            "untrusted": on_list("untrusted") and not trusted,
        }

    @staticmethod
    def _extract_message_id(headers: Optional[str]) -> str:
        """RFC 5322 Message-ID, used to keep rescans from double-counting."""
        match = re.search(
            r"^Message-ID:\s*(.+)$", headers or "", re.IGNORECASE | re.MULTILINE
        )
        return match.group(1).strip() if match else ""

    @staticmethod
    def _count_high_risk_elements(
        link_findings: Optional[List[Any]],
        beacon_count: int,
        attachment_not_ok: int,
    ) -> int:
        """The HRE tally shown on the status strip and charged to the score."""
        def level(item: Any) -> str:
            value = (
                getattr(item, "risk_level", None)
                if not isinstance(item, dict)
                else item.get("risk_level")
            )
            return str(value or "")

        risky_links = sum(
            1 for f in (link_findings or []) if level(f) in ("high", "medium")
        )
        return risky_links + int(beacon_count or 0) + int(attachment_not_ok or 0)

    @staticmethod
    def _link_on_risktable(item: Any) -> bool:
        if isinstance(item, dict):
            return bool(item.get("on_risktable"))
        return bool(getattr(item, "on_risktable", False))

    @classmethod
    def _risktable_disable_urls(cls, link_findings: Optional[List[Any]]) -> List[str]:
        """URLs to defang in the message body (found on RiskTable)."""
        out: List[str] = []
        seen: Set[str] = set()
        for f in link_findings or []:
            if not cls._link_on_risktable(f):
                continue
            url = str(
                f.get("url") if isinstance(f, dict) else getattr(f, "url", "") or ""
            ).strip()
            if not url or url.lower() in seen:
                continue
            seen.add(url.lower())
            out.append(url)
        return out

    @classmethod
    def _risktable_banner_note(
        cls,
        link_findings: Optional[List[Any]],
        *,
        body_available: bool,
        checked: bool,
    ) -> str:
        """Status text after HRE: authentic vs RiskTable hits."""
        if not body_available or not checked:
            return ""
        findings = list(link_findings or [])
        if not findings:
            return ""
        disabled = sum(1 for f in findings if cls._link_on_risktable(f))
        if disabled > 0:
            return f"{disabled} found on RiskTable"
        return "Links appear authentic - not found on RiskTable"

    def _sender_status_label(
        self,
        block_state: Dict[str, bool],
        sender_seen_count: int,
        att_blocked: bool,
        bcn_blocked: bool,
    ) -> str:
        """Unknown / Known/Untrust / Known/Trust, with any blocking appended.

        Known normally comes from how many emails we have seen from the
        sender, but an explicit trust or untrust decision also makes them
        Known: the user cannot have judged a sender they do not know.
        """
        threshold = int(load_sender_status_config().get("known_threshold", 5))
        trusted = bool(block_state.get("trusted"))
        untrusted = bool(block_state.get("untrusted"))
        known = trusted or untrusted or int(sender_seen_count or 0) > threshold

        if not known:
            status = "Unknown"
        elif trusted:
            status = "Known/Trust"
        else:
            # Known but never trusted defaults to Untrust.
            status = "Known/Untrust"

        if att_blocked and bcn_blocked:
            status += " (A&B Blocked)"
        elif att_blocked:
            status += " (A Blocked)"
        elif bcn_blocked:
            status += " (B Blocked)"
        return status

    def _parse_quarantined_count(self, metadata: Dict[str, str]) -> int:
        att = self._metadata_value(metadata, "aes-attachments", "gefooter-attachments")
        m = re.search(r"quarantined=(\d+)", att, re.IGNORECASE)
        return int(m.group(1)) if m else 0

    def _beacon_blocking_enabled(self) -> bool:
        path = (
            Path(os.environ.get("LOCALAPPDATA", CONFIG["base_path"]))
            / "GeoFooter"
            / "aes_beacon_blocking.json"
        )
        try:
            if path.is_file():
                data = json.loads(path.read_text(encoding="utf-8"))
                return bool(isinstance(data, dict) and data.get("enabled"))
        except Exception:
            pass
        return False

    def _render_status_banner_png(
        self, line: str, *, warn: bool, risk_score: int = 0
    ) -> Tuple[str, int, int]:
        """Rasterise the AES status strip to a PNG (Outlook preview-safe).

        Left edge is a colour square for the risk band that fills the full
        height of the strip (staying square): 0–25 green, 25–50 amber, >50 red.
        Status text is vertically centred to the right of that square.
        Returns (local path, width, height), or ('', 0, 0) on failure.
        """
        try:
            from PIL import Image, ImageDraw, ImageFont
        except ImportError:
            self.logger.warning("Pillow not installed — status banner stays HTML text")
            return "", 0, 0

        try:
            score = max(0, min(100, int(risk_score or 0)))
            if score <= 25:
                square_rgb = (76, 175, 80)    # green
            elif score <= 50:
                square_rgb = (255, 167, 38)   # amber
            else:
                square_rgb = (229, 57, 53)    # red

            bg = (255, 248, 230) if warn else (244, 247, 249)       # #fff8e6 / #f4f7f9
            fg = (93, 64, 55) if warn else (55, 71, 79)             # #5d4037 / #37474f
            border_rgb = (230, 194, 0) if warn else (207, 216, 220)  # #e6c200 / #cfd8dc

            # Tahoma 9pt. Pillow sizes in pixels, and 9pt at 96 DPI is 12px.
            fonts_dir = Path(os.environ.get("WINDIR", r"C:\Windows")) / "Fonts"
            font = None
            score_font = None
            for face in ("tahoma.ttf", "arial.ttf"):
                try:
                    font = ImageFont.truetype(str(fonts_dir / face), 12)
                    score_font = ImageFont.truetype(str(fonts_dir / face), 10)
                    break
                except Exception:
                    continue
            if font is None:
                font = ImageFont.load_default()
                score_font = font

            # Extra top padding inside the strip (user: +3px above text).
            pad_top = 7
            pad_bottom = 4
            pad_right = 10
            gap = 8
            # Outlook's Word engine often crops the top few pixels of the first
            # body image. Leave sacrificial white rows ABOVE the bordered strip
            # so the gold frame and text survive that crop.
            clip_guard_top = 6

            probe = Image.new("RGB", (10, 10), bg)
            probe_draw = ImageDraw.Draw(probe)
            bbox = probe_draw.textbbox((0, 0), line, font=font)
            text_w = max(1, bbox[2] - bbox[0])
            text_h = max(1, bbox[3] - bbox[1])

            content_h = max(22, text_h + pad_top + pad_bottom)
            square = content_h  # risk square = full height of bordered strip
            height = clip_guard_top + content_h
            text_left = square + gap
            width = max(720, text_left + text_w + pad_right)

            # White above the strip blends with the reading pane if uncropped.
            img = Image.new("RGB", (width, height), (255, 255, 255))
            draw = ImageDraw.Draw(img)
            y0 = clip_guard_top
            draw.rectangle([0, y0, width - 1, height - 1], fill=bg)

            # Risk square flush left of the bordered strip, full content height.
            draw.rectangle(
                [0, y0, square - 1, height - 1],
                fill=square_rgb,
                outline=square_rgb,
            )
            score_txt = str(score)
            sb = draw.textbbox((0, 0), score_txt, font=score_font)
            sw, sh = sb[2] - sb[0], sb[3] - sb[1]
            draw.text(
                (
                    (square - sw) // 2 - sb[0],
                    y0 + (content_h - sh) // 2 - sb[1],
                ),
                score_txt,
                fill=(255, 255, 255),
                font=score_font,
            )

            # Top-weighted padding: pad_top above ink, pad_bottom below.
            text_y = y0 + pad_top - bbox[1]
            draw.text((text_left, text_y), line, fill=fg, font=font)

            # Border around the content strip only (not the white clip guard).
            draw.rectangle(
                [0, y0, width - 1, height - 1], outline=border_rgb, width=1
            )

            out_dir = Path(CONFIG["base_path"]) / "output" / "banners"
            out_dir.mkdir(parents=True, exist_ok=True)
            digest = hashlib.sha1(f"{score}|{line}".encode("utf-8")).hexdigest()[:12]
            name = f"aes_status_{digest}.png"
            path = out_dir / name
            img.save(path, format="PNG", optimize=True)
            return str(path), width, height
        except Exception as exc:
            self.logger.error("Failed to render status banner PNG: %s", exc)
            return "", 0, 0

    def _build_top_of_mail_banners(
        self,
        *,
        block_state: Dict[str, bool],
        attachment_count: int,
        beacon_count: int,
        beacon_blocked: Optional[int],
        attachment_not_ok: int,
        attachments_quarantined: int,
        link_findings: List[Any],
        risk_score: int = 0,
        risk_level: str = "",
        sender_seen_count: int = 0,
        body_available: bool = False,
        risktable_checked: bool = False,
    ) -> str:
        """AES status strip as a PNG image for the top of the body (preview-safe).

        Outlook's message-list preview scrapes early body text; an <img> with
        empty alt keeps the status out of the preview column while still
        showing at the top of the message when it is opened.
        """
        att_blocked = bool(
            (not block_state.get("trusted"))
            and (block_state.get("attachments") or attachments_quarantined > 0)
        )
        bcn_blocked = bool(
            (not block_state.get("trusted"))
            and (
                block_state.get("beacons")
                or (beacon_blocked or 0) > 0
                or (beacon_count > 0 and self._beacon_blocking_enabled())
            )
        )

        sender_status = self._sender_status_label(
            block_state, sender_seen_count, att_blocked, bcn_blocked
        )

        link_total = len(link_findings or [])
        hre = self._count_high_risk_elements(
            link_findings, beacon_count, attachment_not_ok
        )
        score = max(0, min(100, int(risk_score or 0)))
        level = (risk_level or "").strip().upper() or (
            "LOW" if score <= 25 else ("RAISED" if score <= 50 else "HIGH")
        )
        risktable_note = self._risktable_banner_note(
            link_findings, body_available=body_available, checked=risktable_checked
        )

        # Risk rating sits immediately after Sender status.
        line = (
            f"Sender: {sender_status}  |  "
            f"Risk: {level} ({score}/100)  |  "
            f"Attach: {int(attachment_count or 0)}  |  "
            f"Beacon: {int(beacon_count or 0)}  |  "
            f"Links: {link_total}  |  "
            f"HRE: {hre}"
        )
        if risktable_note:
            line = f"{line}  |  {risktable_note}"
        if score > 25:
            line += "  |  Recommend Full Scan"
        if risk_requires_mitigation(score) and not block_state.get("trusted"):
            line += "  |  TEXT-ONLY MITIGATION"

        warn = hre > 0 or score > 25 or (
            "found on RiskTable" in (risktable_note or "")
        )
        img_path, img_w, img_h = self._render_status_banner_png(
            line, warn=warn, risk_score=score
        )
        if img_path:
            # VBA embeds this PNG as a hidden inline part and rewrites the src to
            # cid:. The file:/// src is only a fallback if embedding fails.
            img_url = "file:///" + str(img_path).replace("\\", "/")
            # The PNG carries its own border, clip-guard, and full-height risk
            # square. Do not wrap it in a bordered/bgcolor table cell — Outlook
            # expands that cell taller than the image and leaves empty cream
            # under the strip.
            return (
                "<!-- AES Top Banners Start -->\n"
                f"<!-- AES-Banner-Img: {img_path} -->\n"
                f"<table border='0' cellpadding='0' cellspacing='0' "
                f"style='border-collapse:collapse; margin:0 0 4px 0;'>"
                f"<tr><td style='padding:0; font-size:0; "
                f"line-height:{img_h}px; mso-line-height-rule:exactly;'>"
                f"<img src='{html.escape(img_url, quote=True)}' "
                f"width='{img_w}' height='{img_h}' "
                f"style='display:block; border:0; outline:none; "
                f"width:{img_w}px; height:{img_h}px;' alt='' />"
                f"</td></tr></table>\n"
                "<!-- AES Top Banners End -->\n"
            )

        bg_hex = "#fff8e6" if warn else "#f4f7f9"
        border_hex = "#e6c200" if warn else "#cfd8dc"
        fg = "#5d4037" if warn else "#37474f"
        sq = (
            "#4CAF50" if score <= 25 else ("#FFA726" if score <= 50 else "#E53935")
        )
        return (
            "<!-- AES Top Banners Start -->\n"
            f"<table border='0' cellpadding='0' cellspacing='0' "
            f"style='border-collapse:collapse; margin:0 0 4px 0;'>"
            f"<tr><td bgcolor='{bg_hex}' "
            f"style='background:{bg_hex}; border:1px solid {border_hex}; padding:0;'>"
            f"<table border='0' cellpadding='0' cellspacing='0' "
            f"style='border-collapse:collapse;'><tr>"
            f"<td bgcolor='{sq}' width='22' height='22' "
            f"style='background:{sq}; width:22px; height:22px; color:#fff; "
            f"font-family:Tahoma,Arial,sans-serif; font-size:9pt; font-weight:bold; "
            f"text-align:center; vertical-align:middle; line-height:22px;'>{score}</td>"
            f"<td style='padding:0 8px; vertical-align:middle; color:{fg}; "
            f"font-family:Tahoma,Arial,sans-serif; font-size:9pt; "
            f"line-height:22px;'>{html.escape(line)}</td>"
            f"</tr></table></td></tr></table>\n"
            "<!-- AES Top Banners End -->\n"
        )

    def _build_sender_blocks_banner(
        self, block_state: Dict[str, bool], who: str
    ) -> str:
        """Deprecated alias kept for older call sites."""
        return self._build_top_of_mail_banners(
            block_state=block_state,
            attachment_count=0,
            beacon_count=0,
            beacon_blocked=None,
            attachment_not_ok=0,
            attachments_quarantined=0,
            link_findings=[],
            risk_score=0,
            risk_level="",
        )

    def _format_link_summary(
        self, link_findings: List[Any], body_available: bool
    ) -> Tuple[str, str]:
        """(summary-line fragment, full-details row html) for link safety checks."""
        if not body_available:
            return ("", "<span style='color: #888888;'>Body not exported — links not checked</span>")

        def field(item: Any, key: str) -> Any:
            return getattr(item, key, None) if not isinstance(item, dict) else item.get(key)

        total = len(link_findings or [])
        if total == 0:
            return ("Links: none found", "<span style='color: #888888;'>No links in message body</span>")

        highs = [f for f in link_findings if str(field(f, "risk_level")) == "high"]
        meds = [f for f in link_findings if str(field(f, "risk_level")) == "medium"]

        if not highs and not meds:
            line = f"Links: {total} OK"
            detail = (
                f"<span style='color: #008000;'>{total} link{'s' if total != 1 else ''} checked "
                f"&#10003; no heuristic flags</span>"
            )
            return (line, detail)

        bad = len(highs) + len(meds)
        line = f"Links: {bad} of {total} suspicious &#9888;"
        parts = [
            f"<span style='color: {'#FF0000' if highs else '#CC7A00'};'>"
            f"{total} checked &mdash; {len(highs)} high risk, {len(meds)} suspicious</span>"
        ]
        for item in (highs + meds)[:4]:
            host = html.escape(str(field(item, "host") or field(item, "url") or "?"))
            reasons = field(item, "reasons") or []
            reason_txt = html.escape("; ".join(str(r) for r in reasons)[:120])
            colour = "#FF0000" if str(field(item, "risk_level")) == "high" else "#CC7A00"
            parts.append(f"<span style='color: {colour};'>{host}</span> <span style='color: #888888;'>({reason_txt})</span>")
        return (line, "<br>".join(parts))

    def _write_links_report_file(
        self,
        link_findings: List[Any],
        body_available: bool,
        metadata: Dict[str, str],
        guri: str,
        sender_email: Optional[str],
        sender_domain: Optional[str],
    ) -> str:
        """Write link-safety HTML report; return file:// URL (empty on failure)."""
        try:
            base_path = Path(CONFIG["base_path"])
            links_dir = base_path / "output" / "links"
            links_dir.mkdir(parents=True, exist_ok=True)

            subject = metadata.get("original subject", "Unknown")
            report_id = self._make_report_id(guri, subject, "links")
            report_name = f"links_report_{report_id}.html"
            report_path = links_dir / report_name

            def field(item: Any, key: str) -> Any:
                return getattr(item, key, None) if not isinstance(item, dict) else item.get(key)

            findings = list(link_findings or [])
            highs = [f for f in findings if str(field(f, "risk_level")) == "high"]
            meds = [f for f in findings if str(field(f, "risk_level")) == "medium"]
            lows = [f for f in findings if str(field(f, "risk_level")) not in {"high", "medium"}]
            ordered = highs + meds + lows

            rows = ""
            for idx, item in enumerate(ordered[:200], 1):
                risk = str(field(item, "risk_level") or "low")
                risk_color = {"high": "#FF4444", "medium": "#CC8800", "low": "#008000"}.get(
                    risk.lower(), "#333"
                )
                reasons = "; ".join(str(r) for r in (field(item, "reasons") or [])) or "—"
                url = str(field(item, "url") or "")
                rows += (
                    f"<tr>"
                    f"<td style='padding:6px;border-bottom:1px solid #eee;'>{idx}</td>"
                    f"<td style='padding:6px;border-bottom:1px solid #eee;color:{risk_color};"
                    f"font-weight:bold;'>{html.escape(risk.upper())}</td>"
                    f"<td style='padding:6px;border-bottom:1px solid #eee;'>"
                    f"{html.escape(str(field(item, 'source') or ''))}</td>"
                    f"<td style='padding:6px;border-bottom:1px solid #eee;'>"
                    f"{html.escape(str(field(item, 'host') or ''))}</td>"
                    f"<td style='padding:6px;border-bottom:1px solid #eee;word-break:break-all;'>"
                    f"{html.escape(url)}</td>"
                    f"<td style='padding:6px;border-bottom:1px solid #eee;font-size:11px;'>"
                    f"{html.escape(reasons)}</td>"
                    f"</tr>"
                )
            if not rows:
                if body_available:
                    rows = (
                        "<tr><td colspan='6' style='padding:8px;'>"
                        "No links found in the message body.</td></tr>"
                    )
                else:
                    rows = (
                        "<tr><td colspan='6' style='padding:8px;'>"
                        "Body was not exported — links were not checked for this scan.</td></tr>"
                    )

            scanned_at = datetime.now(pytz.utc).strftime("%Y%m%d%H%M") + "z"
            summary_colour = (
                "#FF4444" if highs else ("#CC8800" if meds else ("#008000" if findings else "#666"))
            )
            report_html = f"""<!DOCTYPE html>
<html lang="en"><head><meta charset="utf-8"><title>AES Link Safety</title>
<style>body{{font-family:Arial,sans-serif;margin:1.5em;color:#333;background:#fff}}
h1{{font-size:18px;margin-bottom:0.2em}}.meta{{color:#666;font-size:12px;margin-bottom:1em}}
.summary{{background:#f5f5f5;border:1px solid #ddd;border-radius:4px;padding:10px;margin-bottom:1em;font-size:13px}}
table{{width:100%;border-collapse:collapse;font-size:12px}}th{{text-align:left;background:#4c7a2c;color:#fff;padding:6px}}
.footer{{margin-top:1.5em;font-size:11px;color:#666;border-top:1px solid #eee;padding-top:0.8em}}
.note{{font-size:11px;color:#666;margin-top:0.8em;line-height:1.45}}
</style></head><body>
<h1>AES Link Safety</h1>
<div class="meta">Subject: {html.escape(subject)}<br>
Sender: {html.escape(sender_email or 'Unknown')} ({html.escape(sender_domain or 'Unknown')})<br>
Scanned: {scanned_at} | GURI: {html.escape(guri or 'N/A')}</div>
<div class="summary" style="color:{summary_colour};">
<strong>{len(findings)}</strong> link(s) checked —
<span style="color:#FF4444;font-weight:bold;">{len(highs)} high</span>,
<span style="color:#CC8800;font-weight:bold;">{len(meds)} suspicious</span>,
<span style="color:#008000;font-weight:bold;">{len(lows)} low</span>
</div>
<table>
<tr><th>#</th><th>Risk</th><th>Source</th><th>Host</th><th>URL</th><th>Notes</th></tr>
{rows}
</table>
<p class="note">Heuristic checks only (IP hosts, shorteners, lookalikes, suspicious TLDs).
Live Safe Browsing lookups are optional and separate.</p>
<div class="footer">(C) Aliniant Labs | Aliniant Email Scanner (AES)</div>
</body></html>"""
            report_path.write_text(report_html, encoding="utf-8")
            return _suite_file_url("links", report_name)
        except Exception as exc:
            self.logger.error("Failed to write links report: %s", exc)
            return ""

    def _action_link_specs(
        self,
        sender_email: Optional[str],
        sender_domain: Optional[str],
        guri: str,
        block_state: Optional[Dict[str, bool]] = None,
        links_report_url: str = "",
        *,
        include_restore_html: bool = False,
        broker_hit: Optional[Dict[str, Any]] = None,
        subject: str = "",
    ) -> List[Tuple[str, str]]:
        """Ordered (label, url) pairs for AES footer actions (HTML buttons or plain links)."""
        sender = (sender_email or "").strip()
        domain = (sender_domain or "").strip()
        if not sender and (not domain or domain == "Unknown") and not links_report_url and not include_restore_html:
            return []
        block_state = block_state or {}
        query = urllib.parse.urlencode(
            {"sender": sender, "domain": domain, "guri": guri or ""}
        )
        att_on = bool(block_state.get("attachments"))
        bcn_on = bool(block_state.get("beacons"))
        trusted = bool(block_state.get("trusted"))
        untrusted = bool(block_state.get("untrusted"))

        specs: List[Tuple[str, str]] = []
        if links_report_url:
            specs.append(("Show links", self._local_report_href(links_report_url) or links_report_url))
        if sender or (domain and domain != "Unknown"):
            specs.append((
                "Attachments BLOCKED" if att_on else "Block attachments from sender",
                f"aes://block-attachments?{query}",
            ))
            specs.append((
                "Beacons BLOCKED" if bcn_on else "Block beacons from sender",
                f"aes://block-beacons?{query}",
            ))
            trust_label = "Sender Trusted" if trusted else "Trust sender"
            specs.append((trust_label, f"aes://trust-sender?{query}"))
            specs.append((
                "Sender Not Trusted" if untrusted else "Mark not trusted",
                f"aes://untrust-sender?{query}",
            ))
        if broker_hit:
            broker_q = urllib.parse.urlencode(
                {
                    "sender": sender,
                    "domain": domain,
                    "guri": guri or "",
                    "broker": str(broker_hit.get("id") or ""),
                    "subject": (subject or "")[:120],
                }
            )
            specs.append(
                ("Queue data removal", f"aes://broker-removal?{broker_q}")
            )
        if include_restore_html:
            # VBA substitutes the real backup id when applying text-only mitigation.
            specs.append((
                "Restore original HTML format",
                "aes://restore-html?id={{AES_RESTORE_ID}}",
            ))
        return specs

    def _build_action_buttons_plain(
        self,
        sender_email: Optional[str],
        sender_domain: Optional[str],
        guri: str,
        block_state: Optional[Dict[str, bool]] = None,
        links_report_url: str = "",
        *,
        include_restore_html: bool = True,
        broker_hit: Optional[Dict[str, Any]] = None,
        subject: str = "",
    ) -> str:
        """Plain-text AES action links for text-only (mitigated) messages."""
        specs = self._action_link_specs(
            sender_email,
            sender_domain,
            guri,
            block_state,
            links_report_url,
            include_restore_html=include_restore_html,
            broker_hit=broker_hit,
            subject=subject,
        )
        if not specs:
            return ""
        lines = ["AES actions:"]
        for label, url in specs:
            lines.append(f"  {label}:")
            lines.append(f"    {url}")
        return "\n".join(lines)

    def _build_action_buttons_html(
        self,
        sender_email: Optional[str],
        sender_domain: Optional[str],
        guri: str,
        block_state: Optional[Dict[str, bool]] = None,
        links_report_url: str = "",
        broker_hit: Optional[Dict[str, Any]] = None,
        subject: str = "",
    ) -> str:
        """Footer buttons that call back into the AES engine via the aes:// protocol.

        Buttons turn solid red when their block is already active for this
        sender, and Trust sender turns green when the sender is trusted.
        Show links opens the local link-safety report (same visual style).

        Layout uses nested tables rather than CSS padding on <a> tags —
        Outlook's Word engine strips inline-block/margin on anchors, which
        otherwise concatenates the labels into one unreadable run of text.
        """
        specs = self._action_link_specs(
            sender_email,
            sender_domain,
            guri,
            block_state,
            links_report_url,
            broker_hit=broker_hit,
            subject=subject,
        )
        if not specs:
            return ""
        block_state = block_state or {}
        att_on = bool(block_state.get("attachments"))
        bcn_on = bool(block_state.get("beacons"))
        trusted = bool(block_state.get("trusted"))
        untrusted = bool(block_state.get("untrusted"))

        def button_cell(url: str, label: str, active: bool, active_bg: str) -> str:
            if active:
                bg, border, color = active_bg, active_bg, "#ffffff"
            else:
                bg, border, color = "#f2f8fa", "#0f6b7c", "#0f6b7c"
            safe_label = label
            low_label = label.lower()
            if active and "blocked" in low_label:
                safe_label = label + " &#10003;"
            elif active and "trusted" in low_label and "not" not in low_label:
                safe_label = label + " &#10003;"
            elif active and "not trusted" in low_label:
                safe_label = label + " &#10003;"
            return (
                "<td style='padding:2px 6px 2px 0; vertical-align:middle;'>"
                "<table border='0' cellpadding='0' cellspacing='0' "
                "style='border-collapse:separate;'>"
                f"<tr><td bgcolor='{bg}' "
                f"style='background:{bg}; border:1px solid {border}; "
                f"border-radius:3px; padding:4px 10px;'>"
                f"<a href='{html.escape(url, quote=True)}' "
                f"style='color:{color}; font-family:Arial,sans-serif; "
                f"font-size:10px; font-weight:bold; text-decoration:none;'>"
                f"{safe_label}</a>"
                "</td></tr></table></td>"
            )

        cells: List[str] = []
        for label, url in specs:
            active = False
            active_bg = "#0f6b7c"
            low = label.lower()
            if "attachments blocked" in low or (
                "block attachments" in low and att_on
            ):
                active, active_bg = att_on, "#b71c1c"
            elif "beacons blocked" in low or ("block beacons" in low and bcn_on):
                active, active_bg = bcn_on, "#b71c1c"
            elif "sender trusted" in low or (
                low.startswith("trust sender") and trusted
            ):
                active, active_bg = trusted, "#2e7d32"
            elif "not trusted" in low:
                active, active_bg = untrusted, "#ef6c00"
            elif "queue data removal" in low:
                # Distinct teal outline for Incogni-style queue action.
                active, active_bg = False, "#0f6b7c"
                bg = border = "#e8f5f3"
                color = "#0a5c4a"
                safe_label = label
                cells.append(
                    "<td style='padding:2px 6px 2px 0; vertical-align:middle;'>"
                    "<table border='0' cellpadding='0' cellspacing='0' "
                    "style='border-collapse:separate;'>"
                    f"<tr><td bgcolor='{bg}' "
                    f"style='background:{bg}; border:1px solid #0a5c4a; "
                    f"border-radius:3px; padding:4px 10px;'>"
                    f"<a href='{html.escape(url, quote=True)}' "
                    f"style='color:{color}; font-family:Arial,sans-serif; "
                    f"font-size:10px; font-weight:bold; text-decoration:none;'>"
                    f"{safe_label}</a>"
                    "</td></tr></table></td>"
                )
                continue
            cells.append(button_cell(url, label, active, active_bg))

        return (
            "<div style='margin-top:8px; padding-top:6px; border-top:1px solid #ddd;'>"
            "<div style='font-family:Arial,sans-serif; font-size:10px; color:#666; "
            "margin:0 0 4px 0;'>AES actions:</div>"
            "<table border='0' cellpadding='0' cellspacing='0' "
            "style='border-collapse:separate;'>"
            f"<tr>{''.join(cells)}</tr></table></div>"
        )

    def _build_quick_actions_row_html(
        self,
        sender_email: Optional[str],
        sender_domain: Optional[str],
        guri: str,
        block_state: Optional[Dict[str, bool]] = None,
        links_report_url: str = "",
        banner_bg: str = "",
    ) -> str:
        """Third footer row: short Quick Action chips.

        Default chips are white text on green. Attachments/beacons that are
        already blocked stay <BA>/<BB> but turn red. A trusted sender shows
        <ST> in green instead of <TS>.
        """
        sender = (sender_email or "").strip()
        domain = (sender_domain or "").strip()
        has_sender = bool(sender or (domain and domain != "Unknown"))
        links_href = ""
        if links_report_url:
            links_href = self._local_report_href(links_report_url) or links_report_url
        if not has_sender and not links_href:
            return ""

        block_state = block_state or {}
        att_on = bool(block_state.get("attachments"))
        bcn_on = bool(block_state.get("beacons"))
        trusted = bool(block_state.get("trusted"))
        query = urllib.parse.urlencode(
            {"sender": sender, "domain": domain, "guri": guri or ""}
        )
        green_bg, green_fg = AES_CHIP_BG, AES_CHIP_FG
        red_bg, red_fg = AES_CHIP_BLOCKED_BG, "#ffffff"
        # Trusted sender: the chip fills green so it reads at a glance.
        trust_bg, trust_fg = AES_CHIP_TRUSTED_BG, "#ffffff"

        def chip(url: str, code: str, bg: str, fg: str) -> str:
            border = bg if bg != "#ffffff" else fg
            # ~3x the original two-letter chip. width= is what Outlook honours.
            return (
                "<table border='0' cellpadding='0' cellspacing='0' width='78' align='center' "
                "style='border-collapse:separate;'>"
                f"<tr><td bgcolor='{bg}' align='center' width='78' "
                f"style='background:{bg}; border:1px solid {border}; "
                f"width:78px; text-align:center; border-radius:3px; padding:5px 0;'>"
                f"<a href='{html.escape(url, quote=True)}' "
                f"style='color:{fg}; font-family:Arial,sans-serif; "
                f"font-size:10px; font-weight:bold; letter-spacing:1px; text-decoration:none;'>"
                f"{code}</a>"
                "</td></tr></table>"
            )

        planned: List[Tuple[str, str, str, str]] = []
        if links_href:
            planned.append(("SL", links_href, green_bg, green_fg))
        if has_sender:
            planned.append((
                "BA",
                f"aes://block-attachments?{query}",
                red_bg if att_on else green_bg,
                red_fg if att_on else green_fg,
            ))
            planned.append((
                "BB",
                f"aes://block-beacons?{query}",
                red_bg if bcn_on else green_bg,
                red_fg if bcn_on else green_fg,
            ))
            planned.append((
                "ST" if trusted else "TS",
                f"aes://trust-sender?{query}",
                trust_bg if trusted else green_bg,
                trust_fg if trusted else green_fg,
            ))
            planned.append((
                "NT",
                f"aes://untrust-sender?{query}",
                green_bg,
                green_fg,
            ))
        if not planned:
            return ""

        # Shrink-wrapped row so the chips sit together. A 20px spacer cell is
        # what Outlook actually honours between them.
        spacer = (
            "<td width='20' style='width:20px; font-size:1px; line-height:1px;'>&nbsp;</td>"
        )
        label_color = contrasting_text_color(banner_bg)
        if label_color == "#ffffff":
            label_color = AES_STRIP_MUTED
        parts = [
            "<td style='font-family:Arial,sans-serif; font-size:9px; "
            "font-weight:bold; letter-spacing:1.5px; white-space:nowrap; vertical-align:middle;'>"
            f"<span style='color:{label_color};'>QUICK ACTIONS</span></td>"
        ]
        for code, url, bg, fg in planned:
            parts.append(spacer)
            parts.append(
                "<td style='vertical-align:middle;'>"
                f"{chip(url, code, bg, fg)}</td>"
            )
        return (
            "<div style='font-weight:normal; text-align:center;'>"
            "<table border='0' cellpadding='0' cellspacing='0' align='center' "
            "style='border-collapse:separate; margin:0 auto;'>"
            f"<tr>{''.join(parts)}</tr></table></div>"
        )
    
    def _calculate_domain_age(self, creation_date: str) -> Optional[int]:
        """Calculate domain age in days."""
        if not creation_date or creation_date == "Unknown":
            return None
        try:
            date_formats = ["%Y-%m-%d", "%Y-%m-%dT%H:%M:%SZ", "%Y-%m-%d %H:%M:%S"]
            date_str = creation_date.split(' ')[0]
            parsed_date = None
            for fmt in date_formats:
                try:
                    parsed_date = datetime.strptime(date_str, fmt)
                    break
                except ValueError:
                    continue
            if parsed_date:
                parsed_date = pytz.utc.localize(parsed_date)
                now = datetime.now(pytz.utc)
                return max(0, (now - parsed_date).days)
        except Exception as e:
            self.logger.error(f"Error calculating domain age: {e}")
        return None
    
    def _format_authentication(self, auth_info: AuthenticationInfo) -> Dict[str, Tuple[str, str, str] | str]:
        """Convert raw auth results to tuples of (symbol, colour, text)."""
        def status(val: Optional[str]) -> Tuple[str, str, str]:
            if not val or val.lower() in {"none", "missing"}:
                return ("–", "#888888", "None")
            val_lc = val.lower()
            if val_lc == "pass":
                return ("&#10003;", "#008000", "Pass")
            if val_lc == "fail":
                return ("&#10007;", "#ff0000", "Fail")
            return ("–", "#888888", val.capitalize())

        spf = status(auth_info.spf)
        dkim = status(auth_info.dkim)
        dmarc = status(auth_info.dmarc)

        if auth_info.compauth and auth_info.compauth.get("result") not in {"missing", None}:
            comp_res = auth_info.compauth["result"]
            reason = auth_info.compauth.get("reason_code", "unknown")
            reason_txt = COMPAUTH_REASON_CODES.get(reason, f"Reason {reason}")
            comp_disp = f"{comp_res.capitalize()} (reason {reason}: {reason_txt})"
            compauth = status(comp_res)
        else:
            comp_disp = "Not Applicable (Non-Exchange)"
            compauth = ("–", "#888888", "Not Applicable")

        arc_res = (auth_info.arc or {}).get("result", "none")
        arc = status(arc_res)

        return {
            "spf": spf,
            "dkim": dkim,
            "dmarc": dmarc,
            "compauth": compauth,
            "compauth_display": comp_disp,
            "arc": arc,
        }
    
    def _format_security_flags(self, geo_result: GeoLocationResult) -> str:
        """Format security flags. Threat flags: No=green, Yes=red, No Data=amber.
        Hosting/CDN/Mobile are informational (amber when Yes) but still scored.
        """
        flags = []
        display_names = {
            "Hosting": "Hosting / Datacenter",
            "Relay": "Relay",
            "CDN": "CDN",
            "Mobile": "Mobile ISP",
            "Whitelisted": "Whitelisted",
        }
        # Soft infrastructure signals — shown amber (not threat-red), still scored.
        infra_flags = {"Hosting", "CDN", "Mobile"}
        any_bad = False

        for flag_name in SECURITY_FLAG_ORDER:
            is_active = geo_result.security_data.get(flag_name)
            nice = display_names.get(flag_name, flag_name)

            if flag_name == "VPN" and is_active and geo_result.vpn_service:
                service = html.escape(str(geo_result.vpn_service))
                flags.append(
                    f"<span style='color: #FF0000;' class='symbol'>&#10007;</span> "
                    f"<span style='color: #FF0000;'>Is VPN - Yes ({service})</span>"
                )
                any_bad = True
                continue

            if is_active is None:
                if flag_name in SECURITY_NO_DATA_FLAGS:
                    flags.append(
                        f"<span style='color: #FFA500;' class='symbol'>-</span> "
                        f"<span style='color: #FFA500;'>Is {html.escape(nice)} - No Data</span>"
                    )
                continue

            if flag_name == "Whitelisted":
                if is_active:
                    flags.append(
                        f"<span style='color: #008000;' class='symbol'>&#10003;</span> "
                        f"<span style='color: #008000;'>Whitelisted - Yes</span> "
                        f"<span style='color: #666666;'>"
                        f"(AbuseIPDB trust mark: zeroes Hosting/CDN/Mobile/abuse points "
                        f"and halves Tor/Proxy/VPN threat penalty)</span>"
                    )
                else:
                    flags.append(
                        f"<span style='color: #888888;' class='symbol'>-</span> "
                        f"<span style='color: #555555;'>Whitelisted - No</span> "
                        f"<span style='color: #888888;'>"
                        f"(AbuseIPDB: IP not on their trusted-infra list)</span>"
                    )
                continue

            if flag_name in infra_flags:
                if is_active:
                    pts_hint = {
                        "Hosting": "+5 to score",
                        "CDN": "+3 to score",
                        "Mobile": "+2 to score",
                    }.get(flag_name, "")
                    flags.append(
                        f"<span style='color: #CC7A00;' class='symbol'>!</span> "
                        f"<span style='color: #CC7A00;'>Is {html.escape(nice)} - Yes</span> "
                        f"<span style='color: #888888;'>({pts_hint})</span>"
                    )
                else:
                    flags.append(
                        f"<span style='color: #008000;' class='symbol'>&#10003;</span> "
                        f"<span style='color: #008000;'>Is {html.escape(nice)} - No</span>"
                    )
                continue

            # Threat flags: Tor / Proxy / Anonymous / VPN / Relay
            if is_active:
                any_bad = True
                flags.append(
                    f"<span style='color: #FF0000;' class='symbol'>&#10007;</span> "
                    f"<span style='color: #FF0000;'>Is {html.escape(nice)} - Yes</span>"
                )
            else:
                flags.append(
                    f"<span style='color: #008000;' class='symbol'>&#10003;</span> "
                    f"<span style='color: #008000;'>Is {html.escape(nice)} - No</span>"
                )

        # Origin abuse confidence (thresholded in scoring).
        if geo_result.abuse_confidence is not None:
            threshold = int(load_route_risk_config().get("abuse_threshold", 25) or 25)
            conf = int(geo_result.abuse_confidence)
            if conf >= threshold:
                any_bad = True
                flags.append(
                    f"<span style='color: #FF0000;' class='symbol'>&#10007;</span> "
                    f"<span style='color: #FF0000;'>Abuse confidence - {conf}</span> "
                    f"(threshold {threshold})"
                )
            else:
                flags.append(
                    f"<span style='color: #008000;' class='symbol'>&#10003;</span> "
                    f"<span style='color: #008000;'>Abuse confidence - {conf}</span> "
                    f"(threshold {threshold})"
                )
        elif geo_result.success:
            flags.append(
                f"<span style='color: #FFA500;' class='symbol'>-</span> "
                f"<span style='color: #FFA500;'>Abuse confidence - No Data</span>"
            )

        if geo_result.usage_type:
            flags.append(
                f"<span style='color: #555555;'>Usage type: "
                f"{html.escape(str(geo_result.usage_type))}</span>"
            )

        blocklist_all_clear = False
        if geo_result.blocklist and geo_result.blocklist not in [
            "N/A (IPv6/Invalid)",
            "N/A (Invalid IP)",
        ]:
            if "Listed: " in geo_result.blocklist:
                any_bad = True
                blocklists = geo_result.blocklist.replace("Listed: ", "")
                flags.append(
                    f"<span style='color: #FF0000;' class='symbol' title='Blocklisted'>"
                    f"&#10007;</span> <span style='color: #FF0000;'>Blocklisted "
                    f"({html.escape(blocklists)})</span>"
                )
            else:
                blocklist_all_clear = True
                flags.append(
                    f"<span style='color: #008000;' class='symbol' title='Not Blocklisted'>"
                    f"&#10003;</span> <span style='color: #008000;'>Not Blocklisted</span>"
                )

        if not flags:
            return "<span style='color: #FFA500;'>- None</span>"
        joined = "<br>".join(flags)
        if blocklist_all_clear and not any_bad:
            return f"<div>{joined}</div>"
        return joined
    
    def _format_asn(self, asn: str) -> str:
        """Format ASN display."""
        if not asn or asn == "Unknown":
            return "Unknown"
        asn = html.escape(asn)
        if asn.upper().startswith("AS"):
            return asn
        elif asn.isdigit():
            return f"AS{asn}"
        return asn
    
    def _format_email(self, email: Optional[str]) -> str:
        """Format email address for display."""
        if not email:
            return "Unknown"
        return html.escape(email)
    
    def _metadata_value(self, metadata: Dict[str, str], *keys: str) -> str:
        """Read export metadata, accepting legacy GeoFooter key names."""
        for key in keys:
            if key in metadata:
                return metadata[key]
        return ""

    def _scan_exported_attachments(
        self,
        metadata: Dict[str, str],
        deep_mode: bool = False,
        skip_av: bool = False,
    ) -> List[AttachmentVerdict]:
        """Run content-aware attachment scan on files exported by AES/VBA."""
        scan_dir = self._metadata_value(
            metadata, "aes-attachments-dir", "gefooter-attachments-dir"
        )
        if not scan_dir:
            return []

        scanner = AttachmentScanner(deep_mode=deep_mode, use_defender=not skip_av)
        results = scanner.scan_directory(scan_dir)
        self.logger.info(
            "Attachment scan %s (deep=%s av=%s): %d file(s), %d OK, %d Not OK",
            scan_dir,
            deep_mode,
            not skip_av,
            len(results),
            sum(1 for item in results if item.ok),
            sum(1 for item in results if not item.ok),
        )
        for item in results:
            if not item.ok:
                self.logger.warning(
                    "Attachment NOT OK: %s (%s) — %s",
                    item.filename,
                    item.detected_type,
                    "; ".join(item.reasons),
                )
        return results

    def _run_threat_intel(
        self,
        mode: str,
        *,
        per_hop_analysis: List[Dict[str, Any]],
        sender_ip: Optional[str],
        sender_domain: Optional[str],
        headers: str,
        link_findings: List[Any],
        attachment_results: List[Any],
    ) -> List[Any]:
        """Look up hop IPs, domains, links and attachment hashes; annotate them in place."""
        ti = threat_intel
        big_providers = ("google", "microsoft", "amazon", "mimecast")

        def fget(item: Any, key: str) -> Any:
            return item.get(key) if isinstance(item, dict) else getattr(item, key, None)

        ips: List[Tuple[str, str]] = []
        if sender_ip and IPValidator.is_valid_public(sender_ip):
            ips.append((sender_ip, "origin-ip"))
        # Oldest public hops first — they are closest to the real sender.
        for hop in per_hop_analysis or []:
            if hop.get("hop_kind") != "public":
                continue
            ip = str(hop.get("ip") or "").split()[0]
            if not IPValidator.is_valid_public(ip) or any(ip == i for i, _ in ips):
                continue
            org = f"{hop.get('asn_org') or ''} {(hop.get('ip_decode') or {}).get('org_hint') or ''}".lower()
            if any(p in org for p in big_providers):
                continue
            ips.append((ip, "hop-ip"))
        ips = ips[:4]

        header_urls, header_domains = ti.collect_header_indicators(headers)
        domains: List[Tuple[str, str]] = []
        sd = ti.registrable_domain(sender_domain or "") if sender_domain and sender_domain != "Unknown" else ""
        if sd:
            domains.append((sd, "sender-domain"))
        for d in header_domains:
            if d and all(d != x for x, _ in domains):
                domains.append((d, "header-domain"))

        order = {"high": 0, "medium": 1, "low": 2}
        ranked = sorted(
            [f for f in (link_findings or [])
             if str(fget(f, "url") or "").lower().startswith(("http://", "https://"))],
            key=lambda f: order.get(str(fget(f, "risk_level")), 3),
        )
        urls: List[Tuple[str, str]] = []
        for f in ranked:
            u = str(fget(f, "url"))
            if all(u != x for x, _ in urls):
                urls.append((u, "link-url"))
            d = ti.registrable_domain(str(fget(f, "host") or ""))
            if d and all(d != x for x, _ in domains):
                domains.append((d, "link-domain"))
            dest = str(fget(f, "destination") or "")
            if dest:
                dd = ti.registrable_domain(urllib.parse.urlparse(dest).hostname or "")
                if dd and all(dd != x for x, _ in domains):
                    domains.append((dd, "link-domain"))
        for u in header_urls:
            if all(u != x for x, _ in urls):
                urls.append((u, "header-url"))
        urls = urls[:20]
        domains = domains[:10]
        deep_urls = [str(fget(f, "url")) for f in ranked
                     if str(fget(f, "risk_level")) in {"high", "medium"}][:2]

        hashes = [(str(a.sha256), str(getattr(a, "filename", "") or "attachment"))
                  for a in (attachment_results or []) if getattr(a, "sha256", "")]

        reports = ti.ThreatIntel(mode).run(
            ips=ips, domains=domains, urls=urls, hashes=hashes, deep_urls=deep_urls,
        )
        by_key = {(r.kind, r.indicator): r for r in reports}

        for hop in per_hop_analysis or []:
            ip = str(hop.get("ip") or "").split()[0] if hop.get("ip") else ""
            rep = by_key.get(("ip", ip))
            if rep is None:
                continue
            hop["intel_lines"] = rep.lines()
            hop["intel_worst"] = rep.worst

        for f in link_findings or []:
            rep = by_key.get(("url", str(fget(f, "url") or "")))
            dom_rep = by_key.get(("domain", ti.registrable_domain(str(fget(f, "host") or ""))))
            notes: List[str] = []
            worst = "clean"
            seen_any = False
            for r in (rep, dom_rep):
                if r is None:
                    continue
                seen_any = True
                for pr in r.results:
                    if pr.verdict in {"malicious", "suspicious"}:
                        label = ti.PROVIDERS.get(pr.provider, {}).get("label", pr.provider).split(" (")[0]
                        notes.append(f"{label}: {pr.summary}")
                        if pr.verdict == "malicious" or worst == "clean":
                            worst = pr.verdict
            on_table = worst in {"malicious", "suspicious"}
            if isinstance(f, dict):
                f["on_risktable"] = on_table
            else:
                f.on_risktable = on_table
            # Clean TI verdict means the heuristic must not keep calling it suspicious.
            if seen_any and worst == "clean" and str(fget(f, "risk_level")) in {"medium", "high"}:
                cleared = "Threat intel: OK (cleared heuristic risk)"
                reasons = [r for r in (fget(f, "reasons") or []) if r != "No heuristic flags"]
                reasons.append(cleared)
                if isinstance(f, dict):
                    f["reasons"], f["risk_level"] = reasons, "low"
                else:
                    f.reasons, f.risk_level = reasons, "low"
                continue
            if not notes:
                continue
            reasons = [r for r in (fget(f, "reasons") or []) if r != "No heuristic flags"] + notes
            new_level = "high" if worst == "malicious" else (
                "medium" if str(fget(f, "risk_level")) == "low" else str(fget(f, "risk_level")))
            if isinstance(f, dict):
                f["reasons"], f["risk_level"] = reasons, new_level
            else:
                f.reasons, f.risk_level = reasons, new_level

        # Sender / From domain on RiskTable taints every body link.
        sender_rep = by_key.get(
            ("domain", ti.registrable_domain(sender_domain or ""))
        ) if sender_domain else None
        if sender_rep is not None and sender_rep.worst in {"malicious", "suspicious"}:
            for f in link_findings or []:
                if isinstance(f, dict):
                    f["on_risktable"] = True
                else:
                    f.on_risktable = True

        for a in attachment_results or []:
            rep = by_key.get(("hash", str(getattr(a, "sha256", "")).lower()))
            if rep is not None and rep.worst == "malicious":
                try:
                    a.intel_note = "; ".join(rep.lines())[:200]
                except Exception:
                    pass
        return reports

    def _load_deep_body_payload(self, metadata: Dict[str, str]) -> Tuple[str, str]:
        """Load body.html / body.txt sidecars (compact capped or full/deep)."""
        html_path = self._metadata_value(metadata, "aes-body-html-file")
        text_path = self._metadata_value(metadata, "aes-body-text-file")
        body_dir = self._metadata_value(metadata, "aes-body-dir")
        if not html_path and body_dir:
            html_path = str(Path(body_dir) / "body.html")
        if not text_path and body_dir:
            text_path = str(Path(body_dir) / "body.txt")

        html_body = ""
        text_body = ""
        try:
            if html_path and Path(html_path).is_file():
                html_body = Path(html_path).read_text(encoding="utf-8", errors="replace")
        except Exception as exc:
            self.logger.warning("Failed to read deep body HTML: %s", exc)
        try:
            if text_path and Path(text_path).is_file():
                text_body = Path(text_path).read_text(encoding="utf-8", errors="replace")
        except Exception as exc:
            self.logger.warning("Failed to read deep body text: %s", exc)
        return html_body, text_body

    def _summarize_attachment_results(
        self,
        results: List[AttachmentVerdict],
        metadata: Dict[str, str],
    ) -> Tuple[int, int, int]:
        if results:
            return AttachmentScanner.summarize(results)
        return self._parse_attachment_metadata(metadata)

    def _parse_attachment_metadata(self, metadata: Dict[str, str]) -> Tuple[int, int, int]:
        """Return attachment total, OK count, and Not OK count from VBA export metadata."""
        att = self._metadata_value(metadata, "aes-attachments", "gefooter-attachments")
        count_match = re.search(r"count=(\d+)", att, re.IGNORECASE)
        ok_match = re.search(r"ok=(\d+)", att, re.IGNORECASE)
        not_ok_match = re.search(r"not_ok=(\d+)", att, re.IGNORECASE)
        status_match = re.search(r"status=([^;]+)", att, re.IGNORECASE)

        total = int(count_match.group(1)) if count_match else 0
        ok_count = int(ok_match.group(1)) if ok_match else total
        not_ok_count = int(not_ok_match.group(1)) if not_ok_match else 0

        # Legacy export: count + status=OK|Not OK
        if not ok_match and not not_ok_match and status_match:
            status = status_match.group(1).strip().upper()
            if status == "NOT OK" and total > 0:
                not_ok_count = total
                ok_count = 0
            else:
                ok_count = total
                not_ok_count = 0

        return total, ok_count, not_ok_count

    def _aes_subtle_link(
        self,
        href: str,
        inner_html: str,
        *,
        font_weight: str = "normal",
        font_size: str = "11px",
        underline: bool = False,
    ) -> str:
        """Footer link styled as plain text (real href, no box)."""
        deco = "underline" if underline else "none"
        return (
            f'<a href="{html.escape(href, quote=True)}" target="_blank" '
            f'style="color:#ffffff;text-decoration:{deco};background:transparent;border:none;'
            f'padding:0;margin:0;font-weight:{font_weight};font-size:{font_size};'
            f'font-family:Arial,sans-serif;">'
            f"{inner_html}</a>"
        )

    def _local_report_href(self, file_url: str) -> str:
        """Turn a file:///… report URL into aes://open-report (Outlook-safe)."""
        raw = (file_url or "").strip()
        if not raw:
            return ""
        if raw.lower().startswith("aes://"):
            return raw
        path = raw
        if path.lower().startswith("file:///"):
            path = urllib.parse.unquote(path[8:])
        elif path.lower().startswith("file://"):
            path = urllib.parse.unquote(path[7:])
        path = path.replace("/", "\\")
        if path.startswith("\\") and len(path) > 2 and path[2] == ":":
            path = path[1:]
        return "aes://open-report?" + urllib.parse.urlencode({"path": path})

    def _metric_report_link(self, file_url: str, inner_html: str) -> str:
        """Clickable Attachments / Beacons / Links label in the summary strip."""
        href = self._local_report_href(file_url)
        if not href:
            return inner_html
        return self._aes_subtle_link(href, inner_html, underline=True)

    def _make_report_id(self, guri: str, subject: str, suffix: str = "") -> str:
        seed = f"{guri}|{subject}|{suffix}|{datetime.now(pytz.utc).isoformat()}"
        return hashlib.sha256(seed.encode("utf-8")).hexdigest()[:12]

    def _attachment_summary_inner_html(self, total: int, ok_count: int, not_ok_count: int) -> str:
        """Build attachment count + OK/Not OK labels for the footer."""
        ok_style = "color:#90EE90;font-weight:bold;"
        bad_style = "color:#FF4444;font-weight:bold;"
        parts = [f"Attachments: {total}"]

        if total == 0:
            parts.append(f'<span style="{ok_style}">OK</span>')
        else:
            if ok_count > 0:
                if ok_count == total:
                    parts.append(f'<span style="{ok_style}">OK</span>')
                else:
                    parts.append(f'<span style="{ok_style}">{ok_count} OK</span>')
            if not_ok_count > 0:
                if not_ok_count == total:
                    parts.append(f'<span style="{bad_style}">Not OK</span>')
                else:
                    parts.append(f'<span style="{bad_style}">{not_ok_count} Not OK</span>')

        return " ".join(parts)

    def _create_attachment_assets(
        self,
        total: int,
        ok_count: int,
        not_ok_count: int,
        scan_results: List[AttachmentVerdict],
        metadata: Dict[str, str],
        guri: str,
        sender_email: Optional[str],
        sender_domain: Optional[str],
        security_assessment: SecurityAssessment,
    ) -> Tuple[str, str]:
        """Write attachment report page and return footer HTML + report URL."""
        subject = metadata.get("original subject", "Unknown")
        report_id = self._make_report_id(guri, subject, "attachments")
        report_url = self._write_attachment_report_files(
            report_id,
            total,
            ok_count,
            not_ok_count,
            scan_results,
            metadata,
            subject,
            sender_email,
            sender_domain,
            guri,
            security_assessment,
        )
        inner = self._attachment_summary_inner_html(total, ok_count, not_ok_count)
        return self._metric_report_link(report_url, inner), report_url

    def _write_attachment_report_files(
        self,
        report_id: str,
        total: int,
        ok_count: int,
        not_ok_count: int,
        scan_results: List[AttachmentVerdict],
        metadata: Dict[str, str],
        subject: str,
        sender_email: Optional[str],
        sender_domain: Optional[str],
        guri: str,
        security_assessment: SecurityAssessment,
    ) -> str:
        """Write attachment analysis HTML report; return file:// URL."""
        base_path = Path(CONFIG["base_path"])
        attachments_dir = base_path / "output" / "attachments"
        attachments_dir.mkdir(parents=True, exist_ok=True)

        report_name = f"attachment_report_{report_id}.html"
        report_path = attachments_dir / report_name

        rows = ""
        if scan_results:
            for idx, item in enumerate(scan_results, 1):
                status_color = "#008000" if item.ok else "#FF4444"
                status_text = "OK" if item.ok else "Not OK"
                risk_color = {"low": "#008000", "medium": "#CC8800", "high": "#FF4444"}.get(
                    item.risk_level.lower(), "#333333"
                )
                reasons = html.escape("; ".join(item.reasons) if item.reasons else "Clean")
                defender_cell = self._format_antivirus_cell(item)
                rows += (
                    f"<tr>"
                    f"<td style='padding:6px;border-bottom:1px solid #eee;'>{idx}</td>"
                    f"<td style='padding:6px;border-bottom:1px solid #eee;word-break:break-all;'>"
                    f"{html.escape(item.filename)}</td>"
                    f"<td style='padding:6px;border-bottom:1px solid #eee;color:{status_color};font-weight:bold;'>"
                    f"{status_text}</td>"
                    f"<td style='padding:6px;border-bottom:1px solid #eee;color:{risk_color};font-weight:bold;'>"
                    f"{html.escape(item.risk_level.upper())}</td>"
                    f"<td style='padding:6px;border-bottom:1px solid #eee;'>{defender_cell}</td>"
                    f"<td style='padding:6px;border-bottom:1px solid #eee;'>"
                    f"{html.escape(item.declared_extension)}</td>"
                    f"<td style='padding:6px;border-bottom:1px solid #eee;'>"
                    f"{html.escape(item.detected_type)}</td>"
                    f"<td style='padding:6px;border-bottom:1px solid #eee;font-size:11px;'>{reasons}</td>"
                    f"<td style='padding:6px;border-bottom:1px solid #eee;font-size:10px;color:#666;'>"
                    f"{html.escape(item.sha256[:16])}…</td>"
                    f"</tr>"
                )
        elif total > 0:
            rows = (
                f"<tr><td colspan='9' style='padding:8px;'>"
                f"{total} attachment(s) reported by AES but scan results were not available. "
                f"Summary: {ok_count} OK, {not_ok_count} Not OK.</td></tr>"
            )
        else:
            rows = (
                "<tr><td colspan='9' style='padding:8px;'>No attachments on this message.</td></tr>"
            )

        scan_dir = self._metadata_value(metadata, "aes-attachments-dir", "gefooter-attachments-dir")
        scanned_at = datetime.now(pytz.utc).strftime("%Y%m%d%H%M") + "z"
        report_html = f"""<!DOCTYPE html>
<html lang="en"><head><meta charset="utf-8"><title>AES Attachment Analysis</title>
<style>body{{font-family:Arial,sans-serif;margin:1.5em;color:#333;background:#fff}}
h1{{font-size:18px;margin-bottom:0.2em}}.meta{{color:#666;font-size:12px;margin-bottom:1em}}
.summary{{background:#f5f5f5;border:1px solid #ddd;border-radius:4px;padding:10px;margin-bottom:1em;font-size:13px}}
table{{width:100%;border-collapse:collapse;font-size:12px}}th{{text-align:left;background:#4c7a2c;color:#fff;padding:6px}}
.footer{{margin-top:1.5em;font-size:11px;color:#666;border-top:1px solid #eee;padding-top:0.8em}}
</style></head><body>
<h1>AES Attachment Analysis</h1>
<div class="meta">Subject: {html.escape(subject)}<br>
Sender: {html.escape(sender_email or 'Unknown')} ({html.escape(sender_domain or 'Unknown')})<br>
Mail risk: {html.escape(security_assessment.risk_level)} ({security_assessment.score}/100) |
Scanned: {scanned_at} | GURI: {html.escape(guri)}</div>
<div class="summary">
<strong>{total}</strong> attachment(s) —
<span style="color:#008000;font-weight:bold;">{ok_count} OK</span>,
<span style="color:#FF4444;font-weight:bold;">{not_ok_count} Not OK</span>
{f'<br>Scan folder: {html.escape(scan_dir)}' if scan_dir else ''}
</div>
<table>
<tr><th>#</th><th>Filename</th><th>Status</th><th>Risk</th><th>Antivirus</th>
<th>Ext</th><th>Detected type</th><th>Findings</th><th>SHA256</th></tr>
{rows}
</table>
<div class="footer">(C) Aliniant Labs | Aliniant Email Scanner (AES)</div>
</body></html>"""
        report_path.write_text(report_html, encoding="utf-8")
        return _suite_file_url("attachments", report_name)

    def _parse_beacon_metadata(
        self, metadata: Dict[str, str]
    ) -> Tuple[int, Optional[int], List[str]]:
        """Return beacon count, blocked count (None if not reported), and URL list."""
        raw = self._metadata_value(metadata, "aes-beacons", "gefooter-beacons")
        count_match = re.search(r"count=(\d+)", raw, re.IGNORECASE)
        count = int(count_match.group(1)) if count_match else 0
        blocked_match = re.search(r"blocked=(\d+)", raw, re.IGNORECASE)
        blocked = int(blocked_match.group(1)) if blocked_match else None
        urls_match = re.search(r"urls=(.+)", raw, re.IGNORECASE)
        urls: List[str] = []
        if urls_match:
            urls = [u.strip() for u in urls_match.group(1).split("|") if u.strip()]
        return count, blocked, urls

    def _create_beacon_assets(
        self,
        beacon_count: int,
        beacon_urls: List[str],
        metadata: Dict[str, str],
        guri: str,
        sender_email: Optional[str],
        sender_domain: Optional[str],
        beacon_blocked: Optional[int] = None,
    ) -> Tuple[str, str]:
        """Create beacon report page and return footer HTML + report URL."""
        subject = metadata.get("original subject", "Unknown")
        report_id = self._make_report_id(guri, subject, "beacons")
        report_url = self._write_beacon_report_files(
            report_id, beacon_count, beacon_urls, subject, sender_email, sender_domain, guri
        )
        if beacon_count <= 0:
            label = "Beacons: None"
        else:
            label = f"Beacons: {beacon_count} detected"
            if beacon_blocked is not None:
                label = f"{label} / {beacon_blocked} blocked"
        return self._metric_report_link(report_url, html.escape(label)), report_url

    def _ensure_aes_static_pages(self) -> None:
        """Ensure fallback pages exist for beacon viewing without AES installed."""
        pages_dir = _suite_root_path() / "assets" / "pages"
        pages_dir.mkdir(parents=True, exist_ok=True)

        required_path = pages_dir / "aes_required.html"
        if not required_path.exists():
            required_path.write_text(
                """<!DOCTYPE html>
<html lang="en"><head><meta charset="utf-8"><title>Aliniant Email Scanner</title>
<style>body{font-family:Arial,sans-serif;background:#f7f7f7;color:#333;margin:0}
.wrap{max-width:640px;margin:4em auto;background:#fff;border:1px solid #ddd;border-radius:8px;padding:2em;text-align:center}
h1{font-size:20px;margin:0 0 1em}p{font-size:14px;line-height:1.5}</style></head>
<body><div class="wrap"><h1>Aliniant Email Scanner</h1>
<p>You need Aliniant Proprietary Software to view this. Sorry.</p>
<p style="color:#666;font-size:12px;">Install AES (Aliniant Email Scanner) with geolocate_headers.py to open beacon analysis reports.</p>
</div></body></html>""",
                encoding="utf-8",
            )

    def _write_beacon_report_files(
        self,
        report_id: str,
        beacon_count: int,
        beacon_urls: List[str],
        subject: str,
        sender_email: Optional[str],
        sender_domain: Optional[str],
        guri: str,
    ) -> str:
        """Write beacon analysis HTML report; return direct report file URL."""
        base_path = Path(CONFIG["base_path"])
        beacons_dir = base_path / "output" / "beacons"
        beacons_dir.mkdir(parents=True, exist_ok=True)

        report_name = f"beacon_report_{report_id}.html"
        launch_name = f"beacon_open_{report_id}.html"
        report_path = beacons_dir / report_name
        launch_path = beacons_dir / launch_name
        report_url = _suite_file_url("beacons", report_name)

        rows = ""
        entries = list(beacon_urls[:50])
        while len(entries) < beacon_count:
            entries.append("")

        for idx, url in enumerate(entries, 1):
            source, indicator, risk = self._analyze_beacon_entry(url, idx)
            risk_color = {"HIGH": "#FF4444", "MEDIUM": "#CC8800", "LOW": "#008000"}.get(risk, "#333")
            rows += (
                f"<tr>"
                f"<td style='padding:6px;border-bottom:1px solid #eee;'>{idx}</td>"
                f"<td style='padding:6px;border-bottom:1px solid #eee;word-break:break-all;'>{source}</td>"
                f"<td style='padding:6px;border-bottom:1px solid #eee;color:{risk_color};font-weight:bold;'>{risk}</td>"
                f"<td style='padding:6px;border-bottom:1px solid #eee;font-size:11px;'>{indicator}</td>"
                f"</tr>"
            )

        if not rows:
            rows = (
                "<tr><td colspan='4' style='padding:8px;'>No tracking beacons detected.</td></tr>"
            )

        scanned_at = datetime.now(pytz.utc).strftime("%Y%m%d%H%M") + "z"
        report_html = f"""<!DOCTYPE html>
<html lang="en"><head><meta charset="utf-8"><title>AES Beacon Analysis</title>
<style>body{{font-family:Arial,sans-serif;margin:1.5em;color:#333;background:#fff}}
h1{{font-size:18px;margin-bottom:0.2em}}.meta{{color:#666;font-size:12px;margin-bottom:1em}}
.summary{{background:#fff8e6;border:1px solid #e6c200;border-radius:4px;padding:10px;margin-bottom:1em;font-size:13px}}
table{{width:100%;border-collapse:collapse;font-size:12px}}th{{text-align:left;background:#4c7a2c;color:#fff;padding:6px}}
.footer{{margin-top:1.5em;font-size:11px;color:#666;border-top:1px solid #eee;padding-top:0.8em}}
.note{{font-size:11px;color:#666;margin-top:0.8em;line-height:1.45}}
</style></head><body>
<h1>AES Beacon Analysis</h1>
<div class="meta">Subject: {html.escape(subject)}<br>
Sender: {html.escape(sender_email or 'Unknown')} ({html.escape(sender_domain or 'Unknown')})<br>
Scanned: {scanned_at} | GURI: {html.escape(guri)}</div>
<div class="summary">
<strong>{beacon_count}</strong> tracking beacon(s) detected in the HTML body.
These are typically 1×1 images or hidden pixels used to confirm when an email is opened.
</div>
<table>
<tr><th>#</th><th>Source</th><th>Risk</th><th>Analysis</th></tr>
{rows}
</table>
<p class="note">
<strong>Content-ID (cid:)</strong> entries are embedded images referenced inline in the message HTML —
common in Outlook-generated tracking pixels. They are not remote URLs but still indicate read-receipt behaviour.
</p>
<div class="footer">(C) Aliniant Labs | Aliniant Email Scanner (AES)</div>
</body></html>"""
        report_path.write_text(report_html, encoding="utf-8")

        launch_html = f"""<!DOCTYPE html>
<html lang="en"><head><meta charset="utf-8"><title>AES Beacon Analysis</title>
<meta http-equiv="refresh" content="0;url={report_url}">
<style>body{{font-family:Arial,sans-serif;margin:2em;color:#333;text-align:center}}</style>
</head><body>
<p>Opening AES Beacon Analysis…</p>
<p><a href="{report_url}">Click here if you are not redirected</a></p>
</body></html>"""
        launch_path.write_text(launch_html, encoding="utf-8")

        return report_url

    def _analyze_beacon_entry(self, url: str, index: int) -> Tuple[str, str, str]:
        """Return display source HTML, analysis text, and risk level for one beacon."""
        if not url:
            return (
                f"Beacon #{index}",
                "Tracking pixel detected in HTML but source URL was not captured by the mail client.",
                "MEDIUM",
            )

        u = url.lower()
        if u.startswith("cid:"):
            cid = html.escape(url[4:])
            return (
                f"<strong>cid:</strong>{cid}",
                "Embedded MIME image (Content-ID) — inline tracking pixel referenced in HTML body",
                "HIGH",
            )

        indicator = self._classify_beacon_url(url)
        risk = "HIGH"
        if "1x1" in indicator.lower() or "track" in indicator.lower() or "pixel" in indicator.lower():
            risk = "HIGH"
        elif "marketing" in indicator.lower() or "esp" in indicator.lower():
            risk = "MEDIUM"
        return html.escape(url), html.escape(indicator), risk

    def _classify_beacon_url(self, url: str) -> str:
        """Return a short reason string for a tracking URL."""
        u = url.lower()
        if any(x in u for x in ("mailtrack", "/track", "/open?", "pixel", "beacon")):
            return "Tracking pixel / open tracker"
        if any(x in u for x in ("sendgrid.net/wf", "list-manage.com", "hubspot.com")):
            return "Marketing / ESP tracker"
        if "width=1" in u or "height=1" in u:
            return "1x1 image pattern"
        return "Hidden / suspicious image"

    def _generate_or_retrieve_guri(
        self,
        guri_db: Optional['GURIDatabase'],
        sender_email: Optional[str],
        sender_domain: Optional[str],
        security_assessment: SecurityAssessment,
        headers: Optional[str],
        metadata: Dict[str, str],
    ) -> str:
        """Generate or retrieve a GURI record, syncing to the same database as guri_gui.py."""
        guri_sender = sender_email
        if not guri_sender and sender_domain and sender_domain != "Unknown":
            guri_sender = f"unknown@{sender_domain}"

        if not guri_db or not guri_sender:
            return "N/A"

        try:
            recipients = "Unknown"
            if headers:
                to_match = re.search(
                    r'^To:\s*(.+?)(?:\n|$)', headers, re.MULTILINE | re.IGNORECASE
                )
                if to_match:
                    recipients = to_match.group(1).strip()

            dt_str = datetime.now(pytz.utc).strftime('%Y-%m-%d %H:%M:%S')
            avg_risk = f"{security_assessment.risk_level} ({security_assessment.score}/100)"

            subject = metadata.get("original subject", "Unknown")
            if subject == "Unknown" and headers:
                subject_match = re.search(
                    r'^Subject:\s*(.+?)(?:\n|$)', headers, re.MULTILINE | re.IGNORECASE
                )
                if subject_match:
                    subject = subject_match.group(1).strip()
                elif headers:
                    subject = headers.split('\n')[0].strip()

            guri = guri_db.generate_guri(
                guri_sender, recipients, subject, dt_str, avg_risk, document_type="01"
            )
            self.logger.info("GURI synced to database: %s (sender=%s)", guri, guri_sender)
            return guri
        except Exception as e:
            self.logger.error(f"Error generating GURI: {e}")
            return "N/A"

    def _auto_queue_broker(
        self,
        broker_hit: Dict[str, Any],
        sender_email: Optional[str],
        sender_domain: Optional[str],
        headers: Optional[str],
    ) -> None:
        """Queue broker mail for Aura automatically; aura drops mail older than detect_since."""
        if broker_enqueue is None or not headers:
            return
        try:
            m_date = re.search(r"^Date:\s*(.+)$", headers, re.IGNORECASE | re.MULTILINE)
            if not m_date:
                return
            received = email.utils.parsedate_to_datetime(m_date.group(1).strip())
            m_subj = re.search(r"^Subject:\s*(.+)$", headers, re.IGNORECASE | re.MULTILINE)
            item = broker_enqueue(
                sender=sender_email or "",
                domain=sender_domain or "",
                subject=m_subj.group(1).strip() if m_subj else "",
                broker_id=str(broker_hit.get("id") or ""),
                received_at=received.isoformat(),
            )
            if item:
                self.logger.info("Aura: queued %s for removal draft", broker_hit.get("name"))
        except Exception as exc:
            self.logger.debug("Aura auto-queue skipped: %s", exc)

    def _infer_sender_domain(
        self,
        sender_email: Optional[str],
        sender_domain: Optional[str],
        per_hop_analysis: List[Dict[str, Any]],
        headers: Optional[str],
    ) -> Optional[str]:
        """Infer sender domain when From/Return-Path are missing."""
        if sender_domain and sender_domain != "Unknown":
            return sender_domain
        if sender_email and "@" in sender_email:
            return sender_email.split("@", 1)[1].lower()

        if headers:
            for pattern in (
                r'^Received-SPF:[^\n]*\bmailfrom=[^@;\s]+@([^;\s>]+)',
                r'^Authentication-Results:[^\n]*\bsmtp\.mailfrom=[^@;\s]+@([^;\s>]+)',
            ):
                match = re.search(pattern, headers, re.IGNORECASE | re.MULTILINE)
                if match:
                    return match.group(1).lower()

        for hop in per_hop_analysis or []:
            fqdn = hop.get("fqdn") or ""
            if fqdn and fqdn != "Unknown" and "." in fqdn:
                parts = fqdn.split(".")
                if len(parts) >= 2:
                    return ".".join(parts[-2:]).lower()
        return sender_domain
    
    def _build_html(self, data: Dict[str, Any], footer_mode: str = "compact") -> str:
        """Build AES footer HTML.

        compact = two-line summary (plus a Show Full Scan link when score ≤ 50).
        full = summary + detailed analysis embedded in the mail.
        Scores above 50 always embed the full analysis, even in compact mode.
        """
        footer_mode = (footer_mode or "compact").lower()
        if footer_mode not in {"compact", "full"}:
            footer_mode = "compact"

        risk = data["risk_level"]
        score = data["risk_score"]

        def risk_gradient(score):
            s = max(0, min(100, score))
            r = int(76 + (244-76)*s/100)
            g = int(175 + (67-175)*s/100)
            b = int(80 + (54-80)*s/100)
            return f'rgb({r},{g},{b})'

        risk_bg = risk_gradient(score)
        # Ribbon "Full Scan" still uses Basic analysis (embeds full details).
        # Deep Scan is report-only and never builds a footer here.
        scan_type = "Basic"
        summary_line1 = data.get(
            "summary_line1_html",
            f"Aliniant AES Scan Result - {scan_type} Scan | {{{{AES_RISK}}}}",
        )
        # Normalize legacy / provisional labels (incl. old Exhaustive wording).
        summary_line1 = re.sub(
            r"^(?:Mail Scanned|Aliniant AES Scan Result - (?:Basic|Exhaustive|Deep) Scan)",
            f"Aliniant AES Scan Result - {scan_type} Scan",
            summary_line1,
            count=1,
            flags=re.IGNORECASE,
        )

        risk_label = f"Risk {risk} ({score}/100)"
        risk_report_url = self._write_risk_report_file(data, risk, score)
        if risk_report_url:
            risk_html = self._aes_subtle_link(
                self._local_report_href(risk_report_url) or risk_report_url,
                risk_label,
                font_weight="bold",
                font_size="12px",
                underline=True,
            )
        else:
            risk_html = html.escape(risk_label)
        if "{{AES_RISK}}" in summary_line1:
            summary_line1 = summary_line1.replace("{{AES_RISK}}", risk_html)
        else:
            summary_line1 = re.sub(
                r"Risk\s+[A-Za-z]+\s+\(\d+/100\)",
                risk_html,
                summary_line1,
                count=1,
            )

        summary_line2 = data.get("summary_line2_html")
        if not summary_line2:
            summary_line2 = html.escape(data.get("summary_line2", ""))

        # Full analysis is already computed for compact scans; build it once and reuse.
        full_details = self._build_full_details_html(data, risk, score, risk_bg)

        # Above 50: append the full analysis into the mail itself, even on a
        # Short/compact scan. At or below 50 keep the two-line summary and a
        # link to the standalone full-scan report.
        embed_full = footer_mode == "full" or int(score or 0) > 50
        block_state = data.get("block_state") if isinstance(data.get("block_state"), dict) else {}
        # Trusted senders never get HTML stripped — false-positive link noise
        # was converting newsletters to plain text.
        mitigate = risk_requires_mitigation(int(score or 0)) and not bool(
            block_state.get("trusted")
        )

        # Always keep a standalone report on disk when we mitigate, so the
        # text-only mail can still point at the full analysis.
        report_url = ""
        if mitigate or not embed_full:
            report_url = self._write_full_scan_report_file(data, full_details, risk, score) or ""

        if mitigate:
            summary_line2 = summary_line2.replace(
                "{{AES_FULLSCAN}}",
                "HIGH RISK — attachments stripped; message converted to text-only",
            )
        elif embed_full:
            summary_line2 = summary_line2.replace("{{AES_FULLSCAN}}", "Full analysis shown")
        else:
            if report_url:
                show_href = self._local_report_href(report_url) or report_url
                show_link = self._aes_subtle_link(
                    show_href, "Show Full Scan", underline=True
                )
            else:
                show_link = "<span style='text-decoration:underline;'>Show Full Scan</span>"
            summary_line2 = summary_line2.replace("{{AES_FULLSCAN}}", show_link)

        quick_actions = self._build_quick_actions_row_html(
            str(data.get("sender_email") or ""),
            str(data.get("sender_domain") or ""),
            str(data.get("guri") or ""),
            block_state,
            str(data.get("links_report_url") or ""),
            AES_STRIP_BG,
        )
        # Risk is a small coloured dot beside the label; the strip itself is
        # one flat slate colour.
        risk_dot = (
            f"<span style='color:{risk_color_for_level(risk)}; font-size:11px; "
            f"padding-right:5px;'>&#9679;</span>"
        )
        summary_line1 = strip_separators(summary_line1)
        summary_line2 = strip_separators(summary_line2)
        if "Risk" in summary_line1 and risk_dot not in summary_line1:
            summary_line1 = re.sub(
                r"(<a [^>]*>)?(Risk\s)",
                lambda m: f"{risk_dot}{m.group(1) or ''}{m.group(2)}",
                summary_line1,
                count=1,
            )
        actions_row = ""
        if quick_actions:
            actions_row = (
                f"<tr><td style='padding:0 16px 12px 16px;'>"
                f"<div style='border-top:1px solid {AES_STRIP_RULE}; margin:0 0 10px 0; "
                f"font-size:1px; line-height:1px;'>&nbsp;</div>"
                f"{quick_actions}</td></tr>"
            )
        summary_block = (
            f"<table border='0' cellpadding='0' cellspacing='0' width='100%' "
            f"bgcolor='{AES_STRIP_BG}' style='background:{AES_STRIP_BG}; "
            f"border-collapse:collapse;'>"
            f"<tr><td style='padding:12px 16px 3px 16px; text-align:center; "
            f"color:{AES_STRIP_TEXT}; font-family:Arial,sans-serif; font-size:12px; "
            f"font-weight:bold; letter-spacing:0.3px; line-height:1.4;'>{summary_line1}</td></tr>"
            f"<tr><td style='padding:0 16px 12px 16px; text-align:center; "
            f"color:{AES_STRIP_MUTED}; font-family:Arial,sans-serif; font-size:10px; "
            f"font-weight:normal; line-height:1.5;'>{summary_line2}</td></tr>"
            f"{actions_row}"
            f"</table>"
        )

        # Top-of-mail banners sit outside the AES footer block so VBA can
        # inject them after <body>; the footer itself stays the scan strip.
        top_banners = data.get("top_mail_banners") or data.get("sender_blocks_banner") or ""

        disable_urls = data.get("risktable_disable_urls") or []
        risktable_marker = ""
        if disable_urls:
            # Pipe-separated; VBA defangs matching hrefs when applying the footer.
            safe = "|".join(
                str(u).replace("|", "%7C").replace("-->", "") for u in disable_urls[:80]
            )
            risktable_marker = f"<!-- AES-RiskTable-Disable: {safe} -->\n"

        mitigate_markers = ""
        if mitigate:
            plain_lines = [
                f"Risk: {risk} ({int(score)}/100)",
                f"Sender: {data.get('sender_email') or 'Unknown'}",
                "Mitigation: attachments quarantined; message converted to text-only.",
            ]
            if report_url:
                plain_lines.append(f"Full analysis report: {report_url}")
            actions_plain = self._build_action_buttons_plain(
                html.unescape(str(data.get("sender_email") or "")),
                str(data.get("sender_domain") or ""),
                str(data.get("guri") or ""),
                data.get("block_state") if isinstance(data.get("block_state"), dict) else None,
                str(data.get("links_report_url") or ""),
                include_restore_html=True,
            )
            if actions_plain:
                plain_lines.append("")
                plain_lines.append(actions_plain)
            plain_summary = "\n".join(plain_lines)
            mitigate_markers = (
                "<!-- AES-Mitigate: text-strip -->\n"
                "<!-- AES-Plain-Summary-Start -->\n"
                f"{plain_summary}\n"
                "<!-- AES-Plain-Summary-End -->\n"
            )

        if not embed_full:
            return (
                f"{top_banners}"
                f"{risktable_marker}"
                f"{mitigate_markers}"
                f"<!-- AES Start -->\n"
                f"<a id='aes-footer-start' name='aes-footer-start'></a>\n"
                f"<div style='margin:0; padding:0; background:#fff; font-family:Arial, sans-serif; "
                f"font-size:11px; color:#333;'>{summary_block}</div>\n"
                f"<a id='aes-footer-end' name='aes-footer-end'></a>\n"
                f"<!-- AES End -->"
            )

        return (
            f"{top_banners}"
            f"{risktable_marker}"
            f"{mitigate_markers}"
            f"<!-- AES Start -->\n"
            f"<a id='aes-footer-start' name='aes-footer-start'></a>\n"
            f"<div style='margin:0; padding:0; background:#fff; font-family:Arial, sans-serif; "
            f"font-size:11px; color:#333;'>"
            f"{summary_block}\n"
            f"<!-- AES Full Start -->\n"
            f"{full_details}\n"
            f"<!-- AES Full End -->\n"
            f"</div>\n"
            f"<a id='aes-footer-end' name='aes-footer-end'></a>\n"
            f"<!-- AES End -->"
        )

    @staticmethod
    def _risk_row_fill(points: int) -> Tuple[str, bool]:
        """Row background for a score factor: white at 0, solid red at >= 25.

        Returns (hex colour or "" for no fill, True when text should be white).
        Zero and negative (dampening) contributions get no background.
        """
        if points <= 0:
            return "", False
        t = min(points, 25) / 25.0
        red = (211, 47, 47)
        r, g, b = (round(255 + (c - 255) * t) for c in red)
        return f"#{r:02x}{g:02x}{b:02x}", t >= 0.6

    def _build_risk_breakdown_section_html(
        self,
        data: Dict[str, Any],
        risk: str,
        score: int,
        *,
        include_heading: bool = True,
    ) -> str:
        """HTML section explaining how the AES risk score was compiled."""
        components = data.get("risk_components") or []
        rows = []
        points_sum = 0
        for item in components:
            if isinstance(item, RiskScoreComponent):
                category, label, points, note = (
                    item.category,
                    item.label,
                    item.points,
                    item.note,
                )
            elif isinstance(item, dict):
                category = item.get("category", "")
                label = item.get("label", "")
                points = int(item.get("points", 0) or 0)
                note = item.get("note", "") or ""
            else:
                continue
            points_sum += points
            bg, strong_fill = self._risk_row_fill(points)
            point_style = "color:#008000;" if points <= 0 else "color:#b35c00;font-weight:bold;"
            if points >= 50:
                point_style = "color:#c62828;font-weight:bold;"
            if bg and not strong_fill:
                point_style = "color:#7a1f1f;font-weight:bold;"
            text_style = ""
            note_color = "#666"
            if strong_fill:
                point_style = "color:#ffffff;font-weight:bold;"
                text_style = "color:#ffffff;"
                note_color = "#ffffff"
            # Outlook's Word renderer ignores <tr> backgrounds; colour each cell.
            cell_bg = f" bgcolor='{bg}'" if bg else ""
            bg_css = f"background:{bg};" if bg else ""
            points_text = f"+{points}" if points >= 0 else f"&minus;{abs(points)}"
            rows.append(
                "<tr>"
                f"<td{cell_bg} style='padding:6px 8px;border-bottom:1px solid #eee;{bg_css}{text_style}'>"
                f"{html.escape(str(category))}</td>"
                f"<td{cell_bg} style='padding:6px 8px;border-bottom:1px solid #eee;{bg_css}{text_style}'>"
                f"{html.escape(str(label))}</td>"
                f"<td{cell_bg} style='padding:6px 8px;border-bottom:1px solid #eee;text-align:right;"
                f"{bg_css}{point_style}'>{points_text}</td>"
                f"<td{cell_bg} style='padding:6px 8px;border-bottom:1px solid #eee;{bg_css}color:{note_color};'>"
                f"{html.escape(str(note))}</td>"
                "</tr>"
            )

        if not rows:
            rows.append(
                "<tr><td colspan='4' style='padding:8px;color:#666;'>"
                "No scored factors were recorded for this message.</td></tr>"
            )

        thresholds = CONFIG["risk_thresholds"]
        low_max = int(thresholds.get("low_max", 25))
        raised_max = int(thresholds.get("raised_max", 50))
        high_max = int(thresholds.get("high_max", 70))
        domain_flag = data.get("domain_age_flag") or ""
        domain_note = (
            f"<p style='margin:8px 0;font-size:12px;'><strong>Domain caution:</strong> "
            f"{html.escape(str(domain_flag))} (forces risk level HIGH).</p>"
            if domain_flag
            else ""
        )
        heading = (
            "<div style='font-weight:bold;font-size:12px;margin:10px 0 4px 0;color:#0f6b7c;'>"
            "Risk Score Compilation</div>"
            if include_heading
            else ""
        )
        return (
            f"{heading}"
            f"<div style='background:#e8f2f4;border:1px solid #c5dbe0;padding:10px 12px;"
            f"margin:0 0 10px 0;font-size:12px;'>"
            f"<strong>Final score:</strong> {html.escape(risk)} ({score}/100)"
            f"&nbsp;|&nbsp; <strong>Factor points (pre-cap):</strong> {points_sum}"
            f"<br><span style='font-size:11px;color:#555;'>"
            f"Higher score = higher risk. Final score is clamped to 0–100. "
            f"Some factors force the score to 100.</span></div>"
            f"{domain_note}"
            f"<table cellpadding='0' cellspacing='0' style='width:100%;border-collapse:collapse;"
            f"font-size:11px;margin-bottom:8px;'>"
            f"<thead><tr>"
            f"<th style='text-align:left;padding:6px 8px;border-bottom:2px solid #ccc;"
            f"background:#f3f6f7;'>Category</th>"
            f"<th style='text-align:left;padding:6px 8px;border-bottom:2px solid #ccc;"
            f"background:#f3f6f7;'>Factor</th>"
            f"<th style='text-align:right;padding:6px 8px;border-bottom:2px solid #ccc;"
            f"background:#f3f6f7;'>Points</th>"
            f"<th style='text-align:left;padding:6px 8px;border-bottom:2px solid #ccc;"
            f"background:#f3f6f7;'>Notes</th>"
            f"</tr></thead><tbody>{''.join(rows)}</tbody></table>"
            f"<div style='font-size:11px;color:#444;margin:0 0 10px 0;'>"
            f"<strong>How the level is chosen</strong>"
            f"<ul style='margin:4px 0 8px 18px;padding:0;'>"
            f"<li>LOW: score 0–{low_max}</li>"
            f"<li>RAISED: score {low_max + 1}–{raised_max}</li>"
            f"<li>HIGH: score {raised_max + 1}–100, or a domain-age caution flag is set</li>"
            f"<li>Score above {raised_max}: full analysis is embedded in the mail footer</li>"
            f"<li>Score above {high_max}: attachments quarantined and mail converted to text-only</li>"
            f"</ul>"
            f"<strong>Typical point rules</strong>"
            f"<ul style='margin:4px 0 0 18px;padding:0;'>"
            f"<li>SPF/DKIM: pass +0; soft/other +7; fail +15; missing +7</li>"
            f"<li>DMARC: pass +0; soft/other +10; fail +20; missing +10</li>"
            f"<li>CompAuth: pass +0; soft +5; fail +10 / ARC: pass +0; soft +3; fail +8</li>"
            f"<li>Origin IP blocklisted: score forced to 100</li>"
            f"<li>Threat flags (Tor/Proxy/Anon/VPN/Relay): +10 once; Tor/Proxy/Anon/VPN no-data: +2 each</li>"
            f"<li>Hosting/Datacenter +5; CDN +3; Mobile ISP +2 (damped when Whitelisted)</li>"
            f"<li>Origin AbuseIPDB confidence ≥ threshold: +15 (damped when Whitelisted)</li>"
            f"<li>Whitelisted: zeroes soft Hosting/CDN/Mobile/abuse points and halves threat penalty</li>"
            f"<li>High-risk route hop (AbuseIPDB ≥ threshold or listed ASN): +15 once</li>"
            f"<li>Suspicious (medium) links: +2 each; high-risk links: +10 each</li>"
            f"<li>Beacons / bad attachments: +2 each (Basic scan)</li>"
            f"<li>Deep Scan: body heuristics and attachment risk bands add capped points when that mode runs</li>"
            f"<li>Domain age: &lt;10d force 100; &lt;30d +50; &lt;100d +30; &lt;365d +10; else +0</li>"
            f"</ul></div>"
        )

    def _write_risk_report_file(
        self,
        data: Dict[str, Any],
        risk: str,
        score: int,
    ) -> str:
        """Persist risk-score breakdown HTML; return file:// URL for footer link."""
        try:
            base_path = Path(CONFIG["base_path"])
            reports_dir = base_path / "output" / "risk_reports"
            reports_dir.mkdir(parents=True, exist_ok=True)

            guri = str(data.get("guri") or "unknown")
            report_id = self._make_report_id(guri, str(data.get("sender_email") or ""), "risk")
            report_name = f"risk_{report_id}.html"
            report_path = reports_dir / report_name

            sender = html.escape(str(data.get("sender_email") or "Unknown"))
            domain = html.escape(str(data.get("sender_domain") or "Unknown"))
            created = html.escape(str(data.get("created_zulu") or ""))
            breakdown = self._build_risk_breakdown_section_html(
                data, risk, score, include_heading=False
            )

            page = f"""<!DOCTYPE html>
<html lang="en"><head><meta charset="utf-8">
<title>Aliniant AES Risk Score Breakdown</title>
<style>
body {{ font-family: Arial, sans-serif; margin: 16px; color: #222; background: #f7f9fa; }}
.wrap {{ max-width: 900px; margin: 0 auto; background: #fff; padding: 16px 18px;
         border: 1px solid #d5dee3; border-radius: 6px; }}
h1 {{ font-size: 18px; color: #0f6b7c; margin: 0 0 8px 0; }}
.meta {{ color: #555; font-size: 12px; margin-bottom: 12px; }}
.footer {{ margin-top: 18px; font-size: 11px; color: #777; }}
</style></head>
<body><div class="wrap">
<h1>Aliniant AES Risk Score Breakdown</h1>
<div class="meta">Sender: {sender} | Domain: {domain} | Created: {created}</div>
{breakdown}
<div class="footer">(C) Aliniant Labs | Aliniant Email Scanner (AES) — score compiled at scan time</div>
</div></body></html>"""
            report_path.write_text(page, encoding="utf-8")
            return _suite_file_url("risk_reports", report_name)
        except Exception as exc:
            self.logger.error("Failed to write risk report: %s", exc)
            return ""

    def _write_deep_scan_report_file(self, data: Dict[str, Any]) -> str:
        """Write Deep Scan HTML report (browser only — never linked from email footers)."""
        try:
            base_path = Path(CONFIG["base_path"])
            reports_dir = base_path / "output" / "deepscan_reports"
            reports_dir.mkdir(parents=True, exist_ok=True)

            guri = str(data.get("guri") or "unknown")
            report_id = self._make_report_id(guri, str(data.get("sender_email") or ""), "deepscan")
            report_name = f"deepscan_{report_id}.html"
            report_path = reports_dir / report_name

            risk = html.escape(str(data.get("risk_level") or "UNKNOWN"))
            score = int(data.get("risk_score") or 0)
            sender = html.escape(str(data.get("sender_email") or "Unknown"))
            domain = html.escape(str(data.get("sender_domain") or "Unknown"))
            created = html.escape(str(data.get("created_zulu") or ""))
            location = str(data.get("location") or "Unknown")
            org = html.escape(str(data.get("organization") or "Unknown"))
            asn = html.escape(str(data.get("asn") or "Unknown"))
            hops = html.escape(str(data.get("hops_and_tot") or ""))
            auth_type = html.escape(str(data.get("auth_type") or ""))
            sec_flags = str(data.get("security_flags") or "")

            # Threat intel has its own card in the deep report.
            route_table = self._build_full_details_html(
                data, risk, score, str(data.get("risk_color") or "#888"),
                include_threat_intel=False,
            )
            map_html = self._build_hop_map_html(data.get("per_hop_analysis") or [])
            svg_map_html = self._build_route_svg_map_html(data.get("per_hop_analysis") or [])
            links_html = self._build_link_findings_html(data.get("link_findings") or [])
            intel_html = (
                self._build_threat_intel_html(data.get("threat_intel") or [], heading=False)
                or "<p style='font-size:12px;color:#666;'>No threat-intel lookups ran.</p>"
            )
            body_html = self._build_body_findings_html(data.get("body_scan") or {})
            breakdown = self._build_risk_breakdown_section_html(
                data, risk, score, include_heading=True
            )
            beacon_n = int(data.get("beacon_count") or 0)
            beacon_note = (
                f"<p style='font-size:12px;'><strong>Tracking beacons:</strong> {beacon_n} detected "
                f"(see Basic Scan beacon report if previously generated).</p>"
                if beacon_n
                else "<p style='font-size:12px;color:#666;'>No tracking beacons reported in export metadata.</p>"
            )

            page = f"""<!DOCTYPE html>
<html lang="en"><head><meta charset="utf-8">
<title>Aliniant AES Deep Scan Result</title>
<meta name="viewport" content="width=device-width, initial-scale=1">
<link rel="stylesheet" href="https://unpkg.com/leaflet@1.9.4/dist/leaflet.css"
 integrity="sha256-p4NxAoJBhIIN+hmNHrzRCf9tD/miZyoHS5obTRR9BMY=" crossorigin=""/>
<script src="https://unpkg.com/leaflet@1.9.4/dist/leaflet.js"
 integrity="sha256-20nQCchB9co0qIjJZRGuk2/Z9VM+kNiyxNV1lvTlZBo=" crossorigin=""></script>
<style>
body {{ font-family: Segoe UI, Arial, sans-serif; margin: 0; color: #1c2a32; background: #eef3f6; }}
.wrap {{ max-width: 1100px; margin: 0 auto; padding: 18px 16px 40px; }}
.card {{ background: #fff; border: 1px solid #d5dee3; border-radius: 8px; padding: 14px 16px; margin-bottom: 14px; }}
h1 {{ font-size: 22px; color: #0b4f5c; margin: 0 0 6px 0; }}
h2 {{ font-size: 15px; color: #0f6b7c; margin: 0 0 10px 0; }}
.meta {{ color: #555; font-size: 12px; margin-bottom: 10px; }}
.badge {{ display: inline-block; padding: 4px 10px; border-radius: 4px; color: #fff; font-weight: 700; background: {html.escape(str(data.get("risk_color") or "#888"))}; }}
#hopmap {{ height: 380px; width: 100%; border-radius: 6px; border: 1px solid #cfd8dc; }}
table.aes {{ width: 100%; border-collapse: collapse; font-size: 12px; }}
table.aes th, table.aes td {{ border-bottom: 1px solid #e2e8eb; padding: 5px 6px; text-align: left; vertical-align: top; }}
table.aes th {{ color: #334; background: #f4f7f9; }}
.risk-high {{ color: #b71c1c; font-weight: 600; }}
.risk-medium {{ color: #ef6c00; font-weight: 600; }}
.risk-low {{ color: #2e7d32; }}
.footer {{ margin-top: 10px; font-size: 11px; color: #777; }}
.note {{ font-size: 11px; color: #666; margin-top: 8px; }}
</style></head>
<body><div class="wrap">
<div class="card">
<h1>Aliniant AES Deep Scan Result</h1>
<div class="meta">Manual ribbon-only analysis — this report is never linked from the email footer.</div>
<div class="meta">Sender: {sender} | Domain: {domain} | Created: {created} | GURI: {html.escape(guri)}</div>
<div><span class="badge">Risk {risk} ({score}/100)</span>
 &nbsp; Location: {location} &nbsp;|&nbsp; {hops}</div>
</div>

<div class="card">
<h2>Summary</h2>
<p style="margin:4px 0;font-size:13px;"><strong>Organization:</strong> {org} &nbsp;|&nbsp; <strong>ASN:</strong> {asn}</p>
<p style="margin:4px 0;font-size:13px;"><strong>Auth:</strong> {auth_type}</p>
<div style="font-size:12px;">{sec_flags}</div>
</div>

<div class="card">
<h2>Fused risk score</h2>
{breakdown}
</div>

<div class="card">
<h2>Route map</h2>
{map_html}
{svg_map_html}
<p class="note">Public hops with coordinates are plotted. Internal/private hops stay in the route table only.</p>
</div>

<div class="card">
<h2>Route / authentication detail</h2>
{route_table}
</div>

<div class="card">
<h2>Link analysis</h2>
{links_html}
</div>

<div class="card">
<h2>Threat intelligence</h2>
{intel_html}
</div>

<div class="card">
<h2>Body heuristics</h2>
{body_html}
</div>

<div class="card">
<h2>Beacons</h2>
{beacon_note}
</div>

<div class="footer">(C) Aliniant Labs | Aliniant Email Scanner (AES) Deep Scan — ribbon only, never auto-run</div>
</div></body></html>"""
            report_path.write_text(page, encoding="utf-8")
            return _suite_file_url("deepscan_reports", report_name)
        except Exception as exc:
            self.logger.error("Failed to write deep scan report: %s", exc)
            return ""

    # Low-poly world landmass outlines (lon, lat) for the offline SVG route
    # sketch. Deliberately coarse — recognisable shapes, tiny payload.
    _WORLD_OUTLINE: List[List[Tuple[float, float]]] = [
        # North & Central America
        [(-168, 66), (-155, 71), (-130, 70), (-110, 73), (-85, 70), (-80, 60),
         (-65, 60), (-55, 52), (-67, 45), (-70, 42), (-76, 35), (-81, 31),
         (-81, 25), (-84, 30), (-90, 29), (-97, 26), (-97, 22), (-90, 21),
         (-87, 16), (-83, 10), (-79, 9), (-85, 13), (-95, 16), (-105, 20),
         (-110, 23), (-114, 28), (-117, 33), (-122, 37), (-124, 43), (-125, 48),
         (-131, 54), (-140, 59), (-150, 59), (-158, 57), (-165, 54)],
        # Greenland
        [(-45, 60), (-25, 70), (-20, 76), (-30, 83), (-60, 82), (-70, 78),
         (-55, 70)],
        # South America
        [(-77, 8), (-70, 12), (-60, 10), (-52, 5), (-50, 0), (-44, -3),
         (-35, -8), (-39, -15), (-41, -23), (-48, -28), (-53, -34), (-58, -39),
         (-65, -41), (-66, -47), (-69, -52), (-74, -53), (-71, -45), (-73, -40),
         (-71, -32), (-70, -25), (-70, -18), (-76, -14), (-81, -6), (-80, 0)],
        # Eurasia + Arabia
        [(-10, 36), (-9, 43), (-2, 44), (-5, 48), (0, 49), (5, 53), (8, 55),
         (12, 56), (18, 55), (24, 57), (28, 60), (25, 65), (20, 69), (30, 70),
         (42, 68), (55, 68), (70, 72), (85, 74), (100, 77), (115, 74),
         (130, 72), (142, 72), (160, 70), (170, 66), (179, 65), (170, 60),
         (160, 60), (155, 50), (140, 54), (135, 44), (130, 42), (126, 38),
         (122, 37), (121, 30), (115, 22), (108, 18), (105, 10), (103, 1),
         (100, 8), (97, 17), (92, 22), (88, 21), (85, 19), (80, 15), (77, 8),
         (72, 19), (68, 23), (66, 25), (57, 25), (55, 25), (52, 28), (48, 30),
         (60, 25), (58, 22), (55, 17), (45, 12), (43, 15), (39, 21), (35, 28),
         (35, 36), (30, 36), (27, 37), (23, 36), (19, 40), (15, 38), (12, 44),
         (5, 43), (0, 39), (-6, 36)],
        # Africa
        [(-6, 35), (10, 37), (20, 32), (32, 31), (34, 27), (37, 21), (38, 18),
         (43, 11), (51, 12), (45, 2), (40, -3), (40, -10), (35, -20),
         (33, -26), (28, -33), (20, -35), (15, -27), (12, -18), (9, -8),
         (9, 0), (6, 4), (-5, 5), (-8, 4), (-13, 9), (-17, 15), (-17, 21),
         (-13, 27)],
        # Australia
        [(114, -22), (122, -18), (130, -12), (136, -12), (142, -11),
         (145, -15), (149, -20), (153, -27), (150, -37), (144, -38),
         (140, -38), (135, -35), (129, -32), (124, -33), (115, -34),
         (113, -26)],
        # British Isles
        [(-5, 50), (2, 52), (-2, 56), (-4, 58), (-6, 55)],
        # Japan
        [(130, 31), (133, 34), (140, 36), (141, 43), (143, 44), (140, 41),
         (136, 36), (131, 31)],
    ]

    def _route_map_points(self, hops: List[Dict[str, Any]]) -> List[Dict[str, Any]]:
        """Ordered map stops: each hop's sending server, then the server that
        received it (so the final mailbox server is plotted too).

        Consecutive stops within ~25 km merge into one marker listing both.
        """
        raw: List[Dict[str, Any]] = []

        def add(lat: Any, lon: Any, hop_no: int, side: str, name: Any, place: str,
                approx: bool, extra: str = "") -> None:
            try:
                lat_f, lon_f = float(lat), float(lon)
            except (TypeError, ValueError):
                return
            actor = self._journey_actor(name)
            text = f"Hop {hop_no} {side}: {actor}"
            if place:
                text += f" — {place}"
            if extra:
                text += f" ({extra})"
            raw.append({"lat": lat_f, "lon": lon_f, "lines": [text], "approx": approx})

        for idx, hop in enumerate(hops or [], 1):
            if hop.get("hop_kind") != "internal" and hop.get("lat") is not None:
                place = ", ".join(
                    str(p) for p in (hop.get("city"), hop.get("country"))
                    if p and p not in {"Unknown", "—"}
                )
                approx = hop.get("geo_confidence") == "low"
                if approx:
                    place += " (approx.)"
                ip = str(hop.get("ip") or "").split()[0]
                add(hop["lat"], hop["lon"], idx, "from", hop.get("fqdn"), place, approx,
                    ip if ip not in {"", "Unknown"} else "")
            by_loc = hop.get("by_location") or {}
            if by_loc.get("lat") is not None:
                add(by_loc["lat"], by_loc["lon"], idx, "received by", hop.get("by_host"),
                    self._journey_place(by_loc), by_loc.get("confidence") == "low")

        stops: List[Dict[str, Any]] = []
        for p in raw:
            if stops:
                last = stops[-1]
                d_lat = (p["lat"] - last["lat"]) * 111.0
                d_lon = (p["lon"] - last["lon"]) * 111.0 * math.cos(math.radians(p["lat"]))
                if (d_lat * d_lat + d_lon * d_lon) ** 0.5 < 25:
                    for line in p["lines"]:
                        if line not in last["lines"]:
                            last["lines"].append(line)
                    last["approx"] = last["approx"] and p["approx"]
                    continue
            stops.append(dict(p, lines=list(p["lines"])))
        for n, s in enumerate(stops, 1):
            s["n"] = n
            s["label"] = f"Stop {n}: " + " · ".join(s["lines"])
        return stops

    def _build_route_svg_map_html(self, hops: List[Dict[str, Any]]) -> str:
        """Self-contained SVG world map with the mail route plotted hop by hop.

        Unlike the Leaflet map above it, this needs no internet access and
        renders in any browser (and most HTML viewers) as-is.
        """
        points = [
            (s["lon"], s["lat"], s["n"], s["label"]) for s in self._route_map_points(hops)
        ]

        if not points:
            return ""

        def xy(lon: float, lat: float) -> Tuple[float, float]:
            return ((lon + 180.0) * 2.0, (90.0 - lat) * 2.0)

        land = ""
        for poly in self._WORLD_OUTLINE:
            pts = " ".join(
                f"{x:.0f},{y:.0f}" for x, y in (xy(lon, lat) for lon, lat in poly)
            )
            land += (
                f"<polygon points='{pts}' fill='#dde8d8' stroke='#b7c8b0' "
                f"stroke-width='0.6'/>"
            )

        graticule = ""
        for lon_line in range(-150, 180, 30):
            x, _ = xy(lon_line, 0)
            graticule += (
                f"<line x1='{x:.0f}' y1='0' x2='{x:.0f}' y2='360' "
                f"stroke='#cfdde6' stroke-width='0.4'/>"
            )
        for lat_line in range(-60, 90, 30):
            _, y = xy(0, lat_line)
            graticule += (
                f"<line x1='0' y1='{y:.0f}' x2='720' y2='{y:.0f}' "
                f"stroke='#cfdde6' stroke-width='0.4'/>"
            )

        route_pts = [xy(lon, lat) for lon, lat, _, _ in points]
        path = ""
        if len(route_pts) > 1:
            coords = " ".join(f"{x:.1f},{y:.1f}" for x, y in route_pts)
            path = (
                f"<polyline points='{coords}' fill='none' stroke='#c62828' "
                f"stroke-width='1.6' stroke-dasharray='5,3' opacity='0.85'/>"
            )

        dots = ""
        last_i = len(points) - 1
        for i, ((x, y), (_, _, n, _)) in enumerate(zip(route_pts, points)):
            colour = "#2e7d32" if i == 0 else ("#1565c0" if i == last_i else "#c62828")
            dots += (
                f"<circle cx='{x:.1f}' cy='{y:.1f}' r='5' fill='{colour}' "
                f"stroke='#fff' stroke-width='1.2'/>"
                f"<text x='{x:.1f}' y='{y + 2.6:.1f}' text-anchor='middle' "
                f"font-size='6.5' font-family='Arial' fill='#fff' "
                f"font-weight='bold'>{n}</text>"
            )

        legend_items = "".join(
            f"<div style='font-size:11px;color:#445;margin:1px 0;'>{html.escape(lbl)}</div>"
            for _, _, _, lbl in points
        )
        return (
            "<div style='margin-top:10px;'>"
            "<div style='font-size:12px;color:#555;margin-bottom:4px;'>"
            "Offline route sketch (origin <span style='color:#2e7d32;font-weight:bold;'>green</span>, "
            "delivery <span style='color:#1565c0;font-weight:bold;'>blue</span>):</div>"
            "<svg viewBox='0 0 720 360' style='width:100%;max-width:760px;background:#eaf3f8;"
            "border:1px solid #cfd8dc;border-radius:6px;' xmlns='http://www.w3.org/2000/svg'>"
            f"{graticule}{land}{path}{dots}</svg>"
            f"<div style='margin-top:4px;'>{legend_items}</div>"
            "</div>"
        )

    def _build_hop_map_html(self, hops: List[Dict[str, Any]]) -> str:
        stops = self._route_map_points(hops)
        points = [
            {
                "lat": s["lat"], "lon": s["lon"], "n": s["n"], "approx": s["approx"],
                "label": "<br>".join(html.escape(line) for line in s["lines"]),
                "role": "origin" if i == 0 else ("delivery" if i == len(stops) - 1 else "relay"),
            }
            for i, s in enumerate(stops)
        ]

        if not points:
            return "<p style='font-size:12px;color:#666;'>No public hop coordinates available to plot.</p><div id='hopmap' style='display:none'></div>"

        pts_json = json.dumps(points)
        return f"""
<div id="hopmap"></div>
<script>
(function() {{
  var points = {pts_json};
  if (!points.length || typeof L === 'undefined') return;
  var map = L.map('hopmap');
  // tile.openstreetmap.org returns 403 for file:// reports (no Referer), so use
  // CARTO and fall back to Esri if CARTO tiles fail to load.
  var carto = L.tileLayer('https://{{s}}.basemaps.cartocdn.com/rastertiles/voyager/{{z}}/{{x}}/{{y}}{{r}}.png', {{
    maxZoom: 19,
    subdomains: 'abcd',
    attribution: '&copy; OpenStreetMap contributors &copy; CARTO'
  }});
  var tileErrors = 0;
  carto.on('tileerror', function() {{
    if (++tileErrors !== 3) return;
    map.removeLayer(carto);
    L.tileLayer('https://server.arcgisonline.com/ArcGIS/rest/services/World_Street_Map/MapServer/tile/{{z}}/{{y}}/{{x}}', {{
      maxZoom: 19,
      attribution: 'Tiles &copy; Esri'
    }}).addTo(map);
  }});
  carto.addTo(map);
  var colours = {{origin: '#2e7d32', relay: '#c62828', delivery: '#1565c0'}};
  var latlngs = [];
  points.forEach(function(p) {{
    var ll = [p.lat, p.lon];
    latlngs.push(ll);
    var border = p.approx ? '2px dashed #fff' : '2px solid #fff';
    var icon = L.divIcon({{
      className: '',
      iconSize: [24, 24],
      iconAnchor: [12, 12],
      html: "<div style='width:24px;height:24px;border-radius:12px;background:" +
            colours[p.role] + ";border:" + border + ";box-shadow:0 1px 4px rgba(0,0,0,.45);" +
            "color:#fff;font:bold 12px Arial;line-height:24px;text-align:center;'>" + p.n + "</div>"
    }});
    L.marker(ll, {{icon: icon}}).addTo(map).bindPopup(
      "<b>Stop " + p.n + "</b><br>" + p.label + (p.approx ? "<br><i>Approximate location</i>" : "")
    ).bindTooltip("Stop " + p.n);
  }});
  if (latlngs.length > 1) {{
    L.polyline(latlngs, {{color: '#0f6b7c', weight: 3, opacity: 0.75}}).addTo(map);
  }}
  map.fitBounds(latlngs, {{padding: [28, 28], maxZoom: 7}});
}})();
</script>
"""

    def _build_threat_intel_html(self, reports: List[Dict[str, Any]], heading: bool = True) -> str:
        """Indicator × provider verdict table (inline styles; renders in Outlook too).

        Clean/OK rows are collapsed to a one-line summary so the table only
        expands indicators that are actually suspicious or malicious.
        """
        if not reports:
            return ""
        colours = {"malicious": ("#d32f2f", "#fff"), "suspicious": ("#ffb300", "#222"),
                   "clean": ("#e8f5e9", "#1b5e20"), "unknown": ("#f5f5f5", "#555"),
                   "error": ("#f5f5f5", "#999")}
        # "clean" means not suspicious — show OK so it is not read as a threat tier.
        display = {"malicious": "malicious", "suspicious": "suspicious",
                   "clean": "OK", "unknown": "unknown", "error": "error"}
        labels = (threat_intel.PROVIDERS if threat_intel is not None else {})
        kind_order = {"hash": 0, "url": 1, "domain": 2, "ip": 3}
        rank = {"malicious": 0, "suspicious": 1, "clean": 2, "unknown": 3, "error": 4}

        flagged = [r for r in reports if str(r.get("worst") or "") in {"malicious", "suspicious"}]
        ok_count = sum(1 for r in reports if str(r.get("worst") or "") == "clean")
        other = [r for r in reports if str(r.get("worst") or "") not in {"malicious", "suspicious", "clean"}]
        # Prefer flagged (+ unknown/error); hide the OK wall unless nothing else exists.
        show = flagged + other
        if not show and ok_count:
            head = (
                "<div style='font-size:12px; font-weight:600; color:#333; margin:10px 0 4px 0;'>Threat intelligence</div>"
                if heading else ""
            )
            return (
                head
                + f"<div style='font-size:10px; color:#1b5e20; margin:2px 0 6px 0;'>"
                f"All {ok_count} indicator{'s' if ok_count != 1 else ''} OK"
                f" — none flagged by threat intel. "
                f"Run Full or Deep scan to see these in full.</div>"
            )

        rows = []
        for rep in sorted(show, key=lambda r: (rank.get(r.get("worst"), 5), kind_order.get(r.get("kind"), 9))):
            worst = str(rep.get("worst") or "unknown")
            bg, fg = colours.get(worst, colours["unknown"])
            cells = []
            for pr in rep.get("results") or []:
                pbg, pfg = colours.get(pr.get("verdict"), colours["unknown"])
                name = labels.get(pr.get("provider"), {}).get("label", pr.get("provider", "")).split(" (")[0]
                text = f"<b>{html.escape(name)}</b>: {html.escape(str(pr.get('summary') or pr.get('verdict')))}"
                if pr.get("link"):
                    text = f"<a href='{html.escape(pr['link'], quote=True)}' style='color:{pfg};'>{text}</a>"
                cells.append(
                    f"<span style='display:inline-block; margin:1px 3px 1px 0; padding:1px 5px; "
                    f"background:{pbg}; color:{pfg}; border-radius:3px;'>{text}</span>"
                )
            indicator = str(rep.get("indicator") or "")
            if len(indicator) > 70:
                indicator = indicator[:67] + "…"
            role = str(rep.get("role") or "").replace("-", " ")
            rows.append(
                "<tr>"
                f"<td bgcolor='{bg}' style='background:{bg}; color:{fg}; padding:2px 5px; font-weight:600;'>"
                f"{html.escape(display.get(worst, worst))}</td>"
                f"<td style='padding:2px 5px;'>{html.escape(str(rep.get('kind')))}<br>"
                f"<span style='color:#777; font-size:9px;'>{html.escape(role)}</span></td>"
                f"<td style='padding:2px 5px; word-break:break-all;'>{html.escape(indicator)}</td>"
                f"<td style='padding:2px 5px;'>{''.join(cells)}</td>"
                "</tr>"
            )
        head = (
            "<div style='font-size:12px; font-weight:600; color:#333; margin:10px 0 4px 0;'>Threat intelligence</div>"
            if heading else ""
        )
        ok_line = (
            f"<div style='font-size:10px; color:#1b5e20; margin:0 0 4px 0;'>"
            f"{ok_count} other indicator{'s' if ok_count != 1 else ''} OK (hidden).</div>"
            if ok_count else ""
        )
        return (
            head
            + ok_line
            + "<table style='width:100%; border-collapse:collapse; font-size:10px; color:#333;' cellpadding='0' cellspacing='0'>"
            "<tr><th style='text-align:left; border-bottom:1px solid #ccc; padding:2px 5px;'>Verdict</th>"
            "<th style='text-align:left; border-bottom:1px solid #ccc; padding:2px 5px;'>Type</th>"
            "<th style='text-align:left; border-bottom:1px solid #ccc; padding:2px 5px;'>Indicator</th>"
            "<th style='text-align:left; border-bottom:1px solid #ccc; padding:2px 5px;'>Sources</th></tr>"
            + "".join(rows)
            + "</table>"
        )

    def _build_link_findings_html(self, findings: List[Dict[str, Any]]) -> str:
        if not findings:
            return "<p style='font-size:12px;color:#666;'>No links extracted from the message body.</p>"
        rows = []
        for item in findings[:100]:
            risk = str(item.get("risk_level") or "low")
            cls = f"risk-{risk}" if risk in {"low", "medium", "high"} else ""
            reasons = "; ".join(item.get("reasons") or []) or "—"
            rows.append(
                "<tr>"
                f"<td class='{cls}'>{html.escape(risk)}</td>"
                f"<td>{html.escape(str(item.get('source') or ''))}</td>"
                f"<td>{html.escape(str(item.get('host') or ''))}</td>"
                f"<td style='word-break:break-all;'>{html.escape(str(item.get('url') or ''))}</td>"
                f"<td>{html.escape(reasons)}</td>"
                "</tr>"
            )
        return (
            "<table class='aes'><tr><th>Risk</th><th>Source</th><th>Host</th><th>URL</th><th>Notes</th></tr>"
            + "".join(rows)
            + "</table>"
        )

    def _build_body_findings_html(self, body_scan: Dict[str, Any]) -> str:
        findings = (body_scan or {}).get("findings") or []
        display = html.escape(str((body_scan or {}).get("display_name") or ""))
        from_email = html.escape(str((body_scan or {}).get("from_email") or ""))
        from_domain = html.escape(str((body_scan or {}).get("from_domain") or ""))
        header = (
            f"<p style='font-size:12px;margin:0 0 8px 0;'>"
            f"<strong>From display:</strong> {display or '—'} &nbsp;|&nbsp; "
            f"<strong>From:</strong> {from_email or '—'} &nbsp;|&nbsp; "
            f"<strong>Domain:</strong> {from_domain or '—'}</p>"
        )
        if not findings:
            return header + "<p style='font-size:12px;color:#666;'>No body heuristic findings.</p>"
        rows = []
        for item in findings:
            risk = str(item.get("risk_level") or "medium")
            cls = f"risk-{risk}" if risk in {"low", "medium", "high"} else ""
            rows.append(
                "<tr>"
                f"<td class='{cls}'>{html.escape(risk)}</td>"
                f"<td>{html.escape(str(item.get('category') or ''))}</td>"
                f"<td>{html.escape(str(item.get('detail') or ''))}</td>"
                "</tr>"
            )
        return (
            header
            + "<table class='aes'><tr><th>Risk</th><th>Category</th><th>Detail</th></tr>"
            + "".join(rows)
            + "</table>"
        )

    def _write_full_scan_report_file(
        self,
        data: Dict[str, Any],
        full_details: str,
        risk: str,
        score: int,
    ) -> str:
        """Persist already-computed full analysis HTML; return file:// URL for footer link."""
        try:
            base_path = Path(CONFIG["base_path"])
            reports_dir = base_path / "output" / "fullscan_reports"
            reports_dir.mkdir(parents=True, exist_ok=True)

            guri = str(data.get("guri") or "unknown")
            report_id = self._make_report_id(guri, str(data.get("sender_email") or ""), "fullscan")
            report_name = f"fullscan_{report_id}.html"
            report_path = reports_dir / report_name

            sender = html.escape(str(data.get("sender_email") or "Unknown"))
            domain = html.escape(str(data.get("sender_domain") or "Unknown"))
            created = html.escape(str(data.get("created_zulu") or ""))

            page = f"""<!DOCTYPE html>
<html lang="en"><head><meta charset="utf-8">
<title>Aliniant AES Basic Scan Result - Full Analysis</title>
<style>
body {{ font-family: Arial, sans-serif; margin: 16px; color: #222; background: #f7f9fa; }}
.wrap {{ max-width: 1100px; margin: 0 auto; background: #fff; padding: 16px 18px;
         border: 1px solid #d5dee3; border-radius: 6px; }}
h1 {{ font-size: 18px; color: #0f6b7c; margin: 0 0 8px 0; }}
.meta {{ color: #555; font-size: 12px; margin-bottom: 14px; }}
.footer {{ margin-top: 18px; font-size: 11px; color: #777; }}
</style></head>
<body><div class="wrap">
<h1>Aliniant AES Basic Scan Result - Full Analysis</h1>
<div class="meta">Risk: {html.escape(risk)} ({score}/100) |
Sender: {sender} | Domain: {domain} | Created: {created}</div>
{full_details}
<div class="footer">(C) Aliniant Labs | Aliniant Email Scanner (AES) — already calculated at scan time (no re-scan)</div>
</div></body></html>"""
            report_path.write_text(page, encoding="utf-8")
            return _suite_file_url("fullscan_reports", report_name)
        except Exception as exc:
            self.logger.error("Failed to write full scan report: %s", exc)
            return ""

    def _build_full_details_html(
        self,
        data: Dict[str, Any],
        risk: str,
        score: int,
        risk_bg: str,
        *,
        include_threat_intel: bool = True,
    ) -> str:
        """Build the detailed AES analysis table and route map."""
        # Security flags as left-aligned rows
        security_flags_lines = data["security_flags"].split('<br>')
        sec_flags_html = ""
        for flag in security_flags_lines:
            sec_flags_html += f"<div style='margin-bottom:2px;'>{flag}</div>"
        # Main data fields as label-value rows
        fields = [
            ("Sender", data['sender_email']),
            ("Domain", data['sender_domain']),
            ("IP", data['ip_display']),
            ("Hops/ToT", data['hops_and_tot']),
            ("Route", data['routing_info']),
            ("Location", data['location']),
            ("Organization", data['organization']),
            ("A.S.N", data['asn']),
            ("Registrar", data['whois']['registrar']),
            ("Created", f"{data['whois']['created']} ({data['whois']['age']})"),
            ("TLS Certs", data.get('cert_history') or "<span style='color: #888888;'>No data</span>"),
            ("Sent (local)", data.get('sender_local_time') or "<span style='color: #888888;'>Unknown</span>"),
            ("Links", data.get('links_detail_html') or "<span style='color: #888888;'>Not checked</span>"),
            ("Auth Type", data['auth_type']),
            ("SPF", f"<span style='color: {data['auth']['spf'][1]}; font-family:Arial;'>{data['auth']['spf'][0]}</span><span style='color: {data['auth']['spf'][1]};'>{data['auth']['spf'][2]}</span>"),
            ("DKIM", f"<span style='color: {data['auth']['dkim'][1]}; font-family:Arial;'>{data['auth']['dkim'][0]}</span><span style='color: {data['auth']['dkim'][1]};'>{data['auth']['dkim'][2]}</span>"),
            ("DMARC", f"<span style='color: {data['auth']['dmarc'][1]}; font-family:Arial;'>{data['auth']['dmarc'][0]}</span><span style='color: {data['auth']['dmarc'][1]};'>{data['auth']['dmarc'][2]}</span>"),
            ("CompAuth", f"<span style='color: {data['auth']['compauth'][1]}; font-family:Arial;'>{data['auth']['compauth'][0]}</span><span style='color: {data['auth']['compauth'][1]};'>{data['auth']['compauth'][2]}</span> <span style='color: #888888;'>{data['auth']['compauth_display']}</span>"),
            ("ARC", f"<span style='color: {data['auth']['arc'][1]}; font-family:Arial;'>{data['auth']['arc'][0]}</span><span style='color: {data['auth']['arc'][1]};'>{data['auth']['arc'][2]}</span>"),
        ]
        # Use tighter vertical spacing for all rows from Sender to ARC
        row_style = "padding:2px 4px 2px 0; line-height:1.1; vertical-align:middle;"
        rows = ""
        for label, value in fields:
            rows += f"<tr><td style='{row_style} font-weight:bold; font-size:11px; color:#333;'>{label}</td>"
            rows += f"<td style='{row_style} font-size:11px; color:#333;'>{value}</td></tr>"
        # Security flags row (default spacing)
        rows += f"<tr><td style='padding:1px 4px 1px 0; font-weight:bold; font-size:11px; color:#333;'>Security Flags</td>"
        rows += f"<td style='padding:1px 4px 1px 0; font-size:11px; color:#333;'>{sec_flags_html}</td></tr>"
        # Per-Hop Analysis Table
        per_hop_html = ""
        if data.get('per_hop_analysis'):
            per_hop_html += "<div style='font-weight:bold; font-size:11px; margin:8px 0 2px 0;'>Route Table</div>"
            # table-layout:fixed + per-column widths keep long hostnames /
            # reverse-DNS strings wrapping inside the table instead of
            # spilling off the right edge.
            per_hop_html += (
                "<table cellpadding='0' cellspacing='0' style='width:100%; max-width:100%; "
                "table-layout:fixed; border-collapse:collapse; font-size:10px;'>"
            )
            # IP gets more room for decode notes (6to4 / ULA / provider); take a
            # little from Host, Reverse DNS, Prefix and BGP Peers.
            col_widths = [
                "2%", "14%", "16%", "5%", "8%", "5%", "3%", "8%",
                "10%", "8%", "8%", "4%", "5%",
            ]
            per_hop_html += "".join(f"<col style='width:{w};'>" for w in col_widths)
            per_hop_html += "<tr>"
            headers = [
                "#", "Host (Decoded)", "IP", "ASN", "Organisation", "Prefix", "RIR", "Country/City", "Reverse DNS", "Reputation", "Timestamp", "Delay", "BGP Peers"
            ]
            # Very faint outlines so columns stay readable without competing with
            # the row risk colouring. Outlook's Word renderer needs per-cell border.
            cell_edge = "border:1px solid #e4eaee;"
            for h in headers:
                per_hop_html += (
                    f"<th style='{cell_edge} border-bottom:1px solid #c5d0d6; padding:2px 4px; "
                    f"text-align:left; font-size:10px; color:#333; font-weight:600; "
                    f"word-break:break-word; overflow-wrap:anywhere;'>{h}</th>"
                )
            per_hop_html += "</tr>"
            for idx, hop in enumerate(data['per_hop_analysis'], 1):
                # Determine row color
                abuse_score = None
                hop_kind = str(hop.get('hop_kind') or 'unknown')
                if hop['reputation'] and 'Score:' in hop['reputation']:
                    m_score = re.search(r'Score: (\d+)', hop['reputation'])
                    if m_score:
                        try:
                            abuse_score = int(m_score.group(1))
                        except Exception:
                            abuse_score = None
                blocklisted = False
                raw_ip = str(hop.get('ip') or '')
                ip_for_lookup = raw_ip.split()[0] if raw_ip and raw_ip not in {'Unknown', 'Internal', 'localhost / internal'} else ''
                if ip_for_lookup and IPValidator.is_valid_public(ip_for_lookup):
                    geo = GeoLocationService().geolocate_ip(ip_for_lookup)
                    if geo.blocklist and 'Listed:' in geo.blocklist:
                        blocklisted = True
                if (hop.get('address_flag') == 'suspicious' or hop.get('time_flags')
                        or hop.get('intel_worst') == 'malicious'):
                    row_bg = '#F44336'
                    row_color = '#fff'
                    row_weight = '600'
                elif hop_kind == 'internal':
                    row_bg = '#ECEFF1'  # Neutral grey — not a public risk signal
                    row_color = '#37474F'
                    row_weight = '400'
                elif blocklisted or (abuse_score is not None and abuse_score > 10):
                    row_bg = '#F44336'  # Red
                    row_color = '#fff'
                    row_weight = '600'
                elif abuse_score is not None and abuse_score <= 10:
                    row_bg = '#4CAF50'  # Green
                    row_color = '#fff'
                    row_weight = '600'
                else:
                    row_bg = '#FFC107'  # Amber
                    row_color = '#333'
                    row_weight = '400'
                per_hop_html += f"<tr style='background:{row_bg}; color:{row_color}; font-weight:{row_weight};'>"
                per_hop_html += f"<td style='{cell_edge} padding:2px 4px; word-break:break-word; overflow-wrap:anywhere; color:{row_color}; font-weight:{row_weight};'>{idx}</td>"
                host_disp = html.escape(str(hop.get('fqdn') or ''))
                if hop.get('by_host') and str(hop.get('by_host')) != str(hop.get('fqdn') or ''):
                    host_disp += (
                        f"<br><span style='color:{row_color}; font-size:9px; font-weight:{row_weight};'>"
                        f"by {html.escape(str(hop['by_host']))}</span>"
                    )
                loc = hop.get('decoded_location') or hop.get('notes') or ''
                if loc and str(loc) != str(hop.get('fqdn') or ''):
                    host_disp += (
                        f"<br><span style='color:{row_color}; font-size:9px; font-weight:{row_weight};'>"
                        f"({html.escape(str(loc))})</span>"
                    )
                if hop_kind == 'internal':
                    host_disp += (
                        f"<br><span style='color:{row_color}; font-size:9px; font-weight:600;'>"
                        f"[internal hop]</span>"
                    )
                if hop.get('exit_hop'):
                    host_disp += (
                        f"<br><span style='color:{row_color}; font-size:9px; font-weight:600;'>"
                        f"[sending exit]</span>"
                    )
                for rflag in hop.get('route_flags') or []:
                    host_disp += (
                        f"<br><span style='color:{row_color}; font-size:9px; font-weight:700;'>"
                        f"&#9888; {html.escape(str(rflag))}</span>"
                    )
                per_hop_html += f"<td style='{cell_edge} padding:2px 4px; word-break:break-word; overflow-wrap:anywhere; color:{row_color}; font-weight:{row_weight};'>{host_disp}</td>"
                small = f"<br><span style='color:{row_color}; font-size:9px; font-weight:400;'>"
                ip_disp = html.escape(str(hop.get('ip') or 'Unknown'))
                for dline in hop.get('ip_decode_lines') or []:
                    ip_disp += f"{small}{html.escape(str(dline))}</span>"
                per_hop_html += f"<td style='{cell_edge} padding:2px 4px; word-break:break-word; overflow-wrap:anywhere; color:{row_color}; font-weight:{row_weight};'>{ip_disp}</td>"
                per_hop_html += f"<td style='{cell_edge} padding:2px 4px; word-break:break-word; overflow-wrap:anywhere; color:{row_color}; font-weight:{row_weight};'>{html.escape(str(hop.get('asn') or 'Unknown'))}</td>"
                # Force clean the organization name
                asn_org_disp = hop.get('asn_org') or 'Unknown'
                if asn_org_disp and asn_org_disp not in {'Unknown', 'Internal / local hop', 'Insufficient header data'}:
                    cleaned_org = asn_org_disp
                    cleaned_org = re.sub(r'AS\d+\s*', '', cleaned_org, flags=re.IGNORECASE)
                    suffixes = [' Limited', ' Ltd', ' LLC', ' Inc', ' Corporation', ' Corp', ' Company', ' Co']
                    for suffix in suffixes:
                        if cleaned_org.endswith(suffix):
                            cleaned_org = cleaned_org[:-len(suffix)]
                            break
                    cleaned_org = cleaned_org.strip(' ,.-')
                    if not cleaned_org:
                        cleaned_org = 'Unknown'
                    asn_org_disp = (
                        f"<span style='color:{row_color}; font-size:9px; font-weight:{row_weight};'>"
                        f"{html.escape(cleaned_org)}</span>"
                    )
                else:
                    asn_org_disp = html.escape(str(asn_org_disp))
                per_hop_html += f"<td style='{cell_edge} padding:2px 4px; word-break:break-word; overflow-wrap:anywhere; color:{row_color}; font-weight:{row_weight};'>{asn_org_disp}</td>"
                per_hop_html += f"<td style='{cell_edge} padding:2px 4px; word-break:break-word; overflow-wrap:anywhere; color:{row_color}; font-weight:{row_weight};'>{hop['prefix']}</td>"
                per_hop_html += f"<td style='{cell_edge} padding:2px 4px; word-break:break-word; overflow-wrap:anywhere; color:{row_color}; font-weight:{row_weight};'>{hop['rir']}</td>"
                geo_disp = (
                    f"{html.escape(str(hop.get('country') or ''))}"
                    f"<br><span style='color:{row_color}; font-size:9px; font-weight:{row_weight};'>"
                    f"{html.escape(str(hop.get('city') or ''))}</span>"
                )
                if hop.get('location_source'):
                    geo_disp += f"{small}({html.escape(str(hop['location_source']))})</span>"
                else:
                    hl = hop.get('host_location')
                    if hl and addr_decode is not None and hl.get('country') and hl.get('country') != hop.get('country'):
                        geo_disp += f"{small}(host name suggests {html.escape(addr_decode.location_label(hl))})</span>"
                per_hop_html += f"<td style='{cell_edge} padding:2px 4px; word-break:break-word; overflow-wrap:anywhere; color:{row_color}; font-weight:{row_weight};'>{geo_disp}</td>"
                per_hop_html += f"<td style='{cell_edge} padding:2px 4px; word-break:break-word; overflow-wrap:anywhere; color:{row_color}; font-weight:{row_weight};'>{hop['reverse_dns']}</td>"
                rep_disp = str(hop.get('reputation') or '')
                for iline in hop.get('intel_lines') or []:
                    rep_disp += f"{small}{html.escape(str(iline))}</span>"
                per_hop_html += f"<td style='{cell_edge} padding:2px 4px; word-break:break-word; overflow-wrap:anywhere; color:{row_color}; font-weight:{row_weight};'>{rep_disp}</td>"
                ts_disp = html.escape(str(hop.get('timestamp') or ''))
                for tflag in hop.get('time_flags') or []:
                    ts_disp += (
                        f"<br><span style='color:{row_color}; font-size:9px; font-weight:700;'>"
                        f"&#9888; {html.escape(str(tflag))}</span>"
                    )
                per_hop_html += f"<td style='{cell_edge} padding:2px 4px; word-break:break-word; overflow-wrap:anywhere; color:{row_color}; font-weight:{row_weight};'>{ts_disp}</td>"
                per_hop_html += f"<td style='{cell_edge} padding:2px 4px; word-break:break-word; overflow-wrap:anywhere; color:{row_color}; font-weight:{row_weight};'>{hop['delay']}</td>"
                bgp_disp = "<br>".join(html.escape(str(p)) for p in hop.get('bgp_peers') or [])
                for bline in hop.get('bgp_lines') or []:
                    bgp_disp += f"{small}{html.escape(str(bline))}</span>"
                per_hop_html += f"<td style='{cell_edge} padding:2px 4px; word-break:break-word; overflow-wrap:anywhere; color:{row_color}; font-weight:{row_weight};'>{bgp_disp}</td>"
                per_hop_html += "</tr>"
            per_hop_html += "</table>"
            per_hop_html += self._build_route_journey_html(data['per_hop_analysis'])
            if include_threat_intel:
                per_hop_html += self._build_threat_intel_html(data.get('threat_intel') or [])
            # GURI, copyright, datetime at the bottom
            guri = data.get('guri', 'N/A')
            now = datetime.now(pytz.utc)
            created_zulu_str = data.get('created_zulu', now.strftime('%Y%m%d%H%M') + 'z')
            copyright_year = now.strftime('%Y')
            per_hop_html += (
                "<table style='width:100%; margin-top:6px; font-size:11px; color:#333; border-collapse:collapse;' cellpadding='0' cellspacing='0'>"
                "<tr>"
                f"<td style='text-align:left; padding:2px 4px;'>GURI: {guri}</td>"
                f"<td style='text-align:center; padding:2px 4px;'>(C) Aliniant Labs {copyright_year}</td>"
                f"<td style='text-align:right; padding:2px 4px;'>Created: {created_zulu_str}</td>"
                "</tr></table>"
            )
        risk_breakdown = self._build_risk_breakdown_section_html(data, risk, score)
        return (
            f"<table cellpadding='0' cellspacing='0' style='width:100%; border-collapse:collapse; "
            f"font-size:11px; margin-top:4px;'>"
            f"<tr><td colspan='2' style='background:{risk_bg}; color:#fff; font-weight:bold; "
            f"font-size:13px; padding:3px 4px 3px 4px; text-align:center;'>"
            f"EMAIL SECURITY ANALYSIS - RISK {risk} ({score}/100)</td></tr>"
            f"{rows}"
            f"</table>"
            f"{risk_breakdown}"
            f"{self._build_attachment_scan_html(data.get('attachment_scan_results', []))}"
            f"{per_hop_html}"
            f"{data.get('action_buttons_html', '')}"
        )

    def _build_attachment_scan_html(self, results: List[AttachmentVerdict]) -> str:
        if not results:
            return ""

        rows = ""
        for item in results:
            status_color = "#008000" if item.ok else "#FF4444"
            status_text = "OK" if item.ok else "Not OK"
            reasons = html.escape("; ".join(item.reasons) if item.reasons else "Clean")
            defender_cell = self._format_antivirus_cell(item)
            rows += (
                f"<tr>"
                f"<td style='padding:4px;border-bottom:1px solid #eee;'>{html.escape(item.filename)}</td>"
                f"<td style='padding:4px;border-bottom:1px solid #eee;color:{status_color};font-weight:bold;'>{status_text}</td>"
                f"<td style='padding:4px;border-bottom:1px solid #eee;'>{defender_cell}</td>"
                f"<td style='padding:4px;border-bottom:1px solid #eee;'>{html.escape(item.declared_extension)}</td>"
                f"<td style='padding:4px;border-bottom:1px solid #eee;'>{html.escape(item.detected_type)}</td>"
                f"<td style='padding:4px;border-bottom:1px solid #eee;font-size:10px;'>{reasons}</td>"
                f"<td style='padding:4px;border-bottom:1px solid #eee;font-size:9px;'>{html.escape(item.sha256[:16])}…</td>"
                f"</tr>"
            )

        return (
            "<div style='font-weight:bold;font-size:11px;margin:8px 0 2px 0;'>Attachment Scan</div>"
            "<table cellpadding='0' cellspacing='0' style='width:100%;border-collapse:collapse;font-size:10px;'>"
            "<tr>"
            "<th style='border-bottom:1px solid #ccc;padding:4px;text-align:left;'>File</th>"
            "<th style='border-bottom:1px solid #ccc;padding:4px;text-align:left;'>Status</th>"
            "<th style='border-bottom:1px solid #ccc;padding:4px;text-align:left;'>Antivirus</th>"
            "<th style='border-bottom:1px solid #ccc;padding:4px;text-align:left;'>Ext</th>"
            "<th style='border-bottom:1px solid #ccc;padding:4px;text-align:left;'>Detected</th>"
            "<th style='border-bottom:1px solid #ccc;padding:4px;text-align:left;'>Findings</th>"
            "<th style='border-bottom:1px solid #ccc;padding:4px;text-align:left;'>SHA256</th>"
            "</tr>"
            f"{rows}</table>"
        )

    def _format_antivirus_cell(self, item: AttachmentVerdict) -> str:
        status = getattr(item, "defender_status", "") or ""
        threat = getattr(item, "defender_threat", "") or ""
        provider = getattr(item, "av_provider", "") or ""
        prefix = f"{html.escape(provider)}: " if provider else ""
        if status == "clean":
            return f"<span style='color:#008000;'>{prefix}Clean</span>"
        if status == "threat":
            label = threat or "Threat"
            return (
                f"<span style='color:#FF4444;font-weight:bold;'>"
                f"{prefix}{html.escape(label)}</span>"
            )
        if status == "skipped":
            return f"<span style='color:#888888;'>{prefix}Skipped</span>"
        if status == "error":
            return f"<span style='color:#FFA500;'>{prefix}Error</span>"
        return "<span style='color:#888888;'>—</span>"

    # Microsoft 365 / Exchange Online host-label vocabularies.
    # First-label form e.g. AM0PR10CA0036 → Amsterdam / Production / forest 10 / Client Access.
    _M365_DC_CODES: Dict[str, str] = {
        "AM": "Amsterdam",
        "AS": "Asia (Singapore)",
        "BL": "Boydton (VA)",
        "BN": "Boydton (VA)",
        "BY": "Boydton (VA)",
        "CH": "Chicago",
        "CW": "Canada West",
        "DB": "Dublin",
        "DM": "Des Moines",
        "DS": "Dallas",
        "DU": "Dublin",
        "HE": "Helsinki",
        "HK": "Hong Kong",
        "LO": "London",
        "ME": "Melbourne",
        "MN": "Minneapolis",
        "PS": "Pune",
        "SA": "San Antonio",
        "SG": "Singapore",
        "SJ": "San Jose",
        "SN": "San Antonio",
        "SY": "Sydney",
        "TY": "Tokyo",
        "VI": "Vienna",
    }
    _M365_ROLE_CODES: Dict[str, str] = {
        "CA": "Client Access",
        "MB": "Mailbox",
        "WS": "Web Services",
        "FE": "Front End",
        "BE": "Back End",
        "EP": "Edge Proxy",
        "BP": "Backend Proxy",
        "DL": "Directory Lookup",
        "PEPF": "Perimeter / Edge Protection",
    }
    _M365_ENV_CODES: Dict[str, str] = {
        "PR": "Production",
        "P": "Production",
        "SD": "Staging",
        "T": "Test",
    }
    _M365_REGION_CODES: Dict[str, str] = {
        "EUR": "Europe",
        "NAM": "North America",
        "APC": "Asia Pacific",
        "CAN": "Canada",
        "LAM": "Latin America",
        "AFR": "Africa",
        "MEA": "Middle East",
        "JPN": "Japan",
        "KOR": "Korea",
        "IND": "India",
        "AUS": "Australia",
    }

    _JOURNEY_CONTINENTS: Dict[str, str] = {
        **{cc: "North America" for cc in ("US", "CA", "MX")},
        **{cc: "Europe" for cc in (
            "GB", "IE", "NL", "DE", "FR", "AT", "FI", "CH", "NO", "SE", "IT",
            "ES", "PL", "BE", "DK", "PT", "EU",
        )},
        **{cc: "Asia Pacific" for cc in ("SG", "HK", "JP", "KR", "MY", "IN", "AU", "NZ")},
        "BR": "South America", "ZA": "Africa", "AE": "Middle East", "BH": "Middle East",
    }

    @staticmethod
    def _journey_place(loc: Optional[Dict[str, Any]]) -> str:
        if not loc:
            return ""
        place = ", ".join(p for p in (loc.get("city"), loc.get("country")) if p)
        if place and loc.get("confidence") == "low":
            place += " (approx.)"
        elif place and loc.get("confidence") not in (None, "high"):
            place += " (likely)"
        return place

    @staticmethod
    def _journey_actor(name: Optional[str]) -> str:
        name = str(name or "").strip()
        if not name:
            return "an unnamed server"
        if addr_decode is not None:
            d = addr_decode.decode_m365_host(name)
            if d and d.get("role_name") and not d.get("tenant"):
                role = str(d["role_name"]).split(" (")[0]
                return f"Microsoft {role} {name.split('.')[0].upper()}"
        return name

    def _build_route_journey_html(self, hops: List[Dict[str, Any]]) -> str:
        """Plain-English walk through the Received chain (origin → mailbox)."""
        if not hops:
            return ""
        lines: List[str] = []
        prev_cc = ""
        total = 0.0
        blank = {None, "", "Unknown", "—", "NA", "N/A"}
        for idx, hop in enumerate(hops, 1):
            from_loc = hop.get("host_location")
            if not from_loc and hop.get("country") not in blank:
                from_loc = {"city": hop.get("city") if hop.get("city") not in blank else "",
                            "country": hop.get("country"),
                            "confidence": hop.get("geo_confidence") or "high"}
            by_loc = hop.get("by_location")
            from_place = self._journey_place(from_loc)
            by_place = self._journey_place(by_loc)
            try:
                delay = float(str(hop.get("delay") or "").rstrip("s"))
            except ValueError:
                delay = None
            if delay is not None:
                total += max(0.0, delay)

            if hop.get("hop_kind") == "internal":
                ip = str(hop.get("ip") or "").replace(" (internal)", "")
                text = (
                    f"Starts on {hop.get('fqdn') or 'an internal host'} inside a private "
                    f"network ({ip}) — the sending platform's own relay."
                )
            else:
                src = self._journey_actor(hop.get("fqdn"))
                if from_place:
                    src += f" in {from_place}"
                dst = self._journey_actor(hop.get("by_host"))
                if by_place:
                    dst += f" in {by_place}"
                prefix = "Leaves the sending platform: " if hop.get("exit_hop") else ""
                text = f"{prefix}{src} hands the message to {dst}."
                proto = str(hop.get("protocol") or "").upper()
                if re.match(r"^(E?SMTPS?A?|LMTP|HTTPS?)$", proto):
                    text = text[:-1] + f" over {proto}."
            if delay is not None and delay >= 10:
                text += f" {delay:.0f}s spent here (typical of inbound filtering or queuing)."
            elif delay is not None and delay < 0:
                text += " Clock goes backwards here — check the timestamp flags."

            here_cc = (by_loc or {}).get("country") or (from_loc or {}).get("country") or ""
            if prev_cc and here_cc and here_cc != prev_cc:
                a = self._JOURNEY_CONTINENTS.get(prev_cc, prev_cc)
                b = self._JOURNEY_CONTINENTS.get(here_cc, here_cc)
                if a != b:
                    text += f" Crosses from {a} to {b}."
                else:
                    text += f" Moves from {prev_cc} to {here_cc}."
            if here_cc:
                prev_cc = here_cc
            for rflag in hop.get("route_flags") or []:
                text += f" ⚠ Tromboning: {rflag}."
            bgp = hop.get("bgp") or {}
            if any(str(s).startswith("invalid") for s in (bgp.get("rpki") or {}).values()):
                text += " ⚠ Its IP prefix is announced by a network RPKI does not authorise (possible BGP hijack)."
            lines.append(f"<b>Hop {idx}:</b> {html.escape(text)}")

        last = hops[-1]
        dest = self._journey_actor(last.get("by_host"))
        dest_place = self._journey_place(last.get("by_location"))
        summary = f"Delivered by {dest}" + (f" in {dest_place}" if dest_place else "")
        summary += f" — {len(hops)} hop(s), about {total:.0f}s end to end."
        lines.append(f"<b>{html.escape(summary)}</b>")
        return (
            "<div style='font-weight:bold; font-size:11px; margin:8px 0 2px 0;'>Route Journey</div>"
            "<div style='font-size:10px; color:#333; line-height:1.45;'>"
            + "<br>".join(lines)
            + "</div>"
        )

    def _m365_city_name(self, dc: str, region_group: Optional[str]) -> str:
        if addr_decode is not None:
            hit = addr_decode.m365_city(dc, region_group)
            if hit:
                key, conf = hit
                city = addr_decode._CITIES[key][0]
                return city if conf == "high" else f"{city} ({conf} confidence)"
            return f"datacenter {dc}"
        return self._M365_DC_CODES.get(dc, dc)

    def _decode_m365_host_label(self, label: str, region_group: Optional[str] = None) -> Optional[str]:
        """Decode an Exchange Online first-label hostname fragment.

        Examples:
          AM0PR10CA0036  → Amsterdam / Production / forest 10 / Client Access #0036
          DU0P250MB0793  → Dublin / Production / forest 250 / Mailbox #0793
          AMBP250MB1516  → Amsterdam / Backend Proxy / forest 250 / Mailbox #1516
          AM1PEPF000252DE → Amsterdam / Perimeter / Edge Protection #000252
        """
        if not label:
            return None
        raw = label.strip().upper()
        if not raw or len(raw) < 4:
            return None

        # Perimeter / PEPF form: AM1PEPF000252DE
        m = re.match(r'^([A-Z]{2})([0-9A-Z])(PEPF)([0-9A-F]+)$', raw)
        if m:
            dc, ring, role, num = m.groups()
            city = self._m365_city_name(dc, region_group)
            role_name = self._M365_ROLE_CODES.get(role, role)
            return f"{city} / {role_name} #{num} (ring {ring})"

        # Classic form: AM0PR10CA0036 / LO4P123CA0189 / DU0P250MB0793 / CWLP265MB1234
        # (ring can be a digit or letter: AMBP250…, CWLP265…, PAXP193…)
        m = re.match(
            r'^([A-Z]{2})([0-9A-Z])(PR|P|SD|T)(\d+)(CA|MB|WS|FE|BE|EP|BP|DL)(\d+)$',
            raw,
        )
        if m:
            dc, ring, env, forest, role, num = m.groups()
            city = self._m365_city_name(dc, region_group)
            env_name = self._M365_ENV_CODES.get(env, env)
            role_name = self._M365_ROLE_CODES.get(role, role)
            return (
                f"{city} / {env_name} / forest {forest} / "
                f"{role_name} #{num} (ring {ring})"
            )

        return None

    def _decode_pod_name(self, segment: str) -> str:
        """Decode M365 region/pod second-label (EURPRD10, EURP250, GBRP123)."""
        if not segment:
            return segment
        seg = segment.strip().upper()

        # EURPRD10 / NAMPRD05 / APCPRD01
        m = re.match(r'^([A-Z]{3})PRD(\d+)$', seg)
        if m:
            region, num = m.groups()
            region_name = self._M365_REGION_CODES.get(region, region)
            return f"{region_name} Production forest {num}"

        # EURP250 / NAMP123
        m = re.match(r'^([A-Z]{3})P(\d+)$', seg)
        if m:
            region, num = m.groups()
            region_name = self._M365_REGION_CODES.get(region, region)
            return f"{region_name} Pod {num}"

        # Legacy ISO-ish: SWEP280, GBRP123
        m = re.match(r'^([A-Z]{2,3})([A-Z])P(\d+)$', seg)
        if m:
            country_code, _region_letter, pod_number = m.groups()
            country = pycountry.countries.get(alpha_2=country_code[:2])
            country_name = country.name if country else country_code
            return f"{country_name} Pod {pod_number}"

        return segment

    def _decode_m365_hostname(self, hostname: Optional[str]) -> Optional[str]:
        """Full M365 FQDN decode — first label (server) + optional region label."""
        if not hostname:
            return None
        host = hostname.strip().rstrip(".").lower()
        if not host:
            return None
        if addr_decode is not None:
            described = addr_decode.describe_m365_host(host)
            if described:
                return described
        # Only attempt rich decode on known Microsoft mail infrastructure.
        if not any(
            token in host
            for token in (
                "outlook.com",
                "office365.com",
                "protection.outlook.com",
                "mail.protection",
            )
        ):
            # Still try first-label decode — some internal hops omit the suffix
            # in Received lines but keep the AM0PR… form.
            label0 = host.split(".", 1)[0]
            return self._decode_m365_host_label(label0)

        parts = host.split(".")
        bits: List[str] = []
        region_group = (
            addr_decode.m365_region_group(parts[1])
            if addr_decode is not None and len(parts) > 1
            else None
        )
        first = self._decode_m365_host_label(parts[0], region_group)
        if first:
            bits.append(first)
        if len(parts) > 1:
            second = self._decode_pod_name(parts[1])
            # Only keep second if it actually decoded into something richer.
            if second and second.upper() != parts[1].upper():
                bits.append(second)
        return " | ".join(bits) if bits else None

    def _extract_routing_info(self, headers: Optional[str]) -> str:
        if not headers:
            return "N/A"
        received_lines = HeaderParser.extract_received_blocks(headers)
        hops = []
        for line in received_lines:
            # Extract all possible IPs (IPv4/IPv6)
            ips = re.findall(r'\[([\d\.:a-fA-F]+)\]', line)
            host = None
            m_host = re.search(r'from ([^ ]+)', line)
            decoded_location = None
            if m_host:
                host = m_host.group(1)
                decoded_location = self._decode_m365_hostname(host)
            asn = "Unknown"
            found_public_ip = False
            ip_used = None
            for ip in ips:
                if IPValidator.is_valid_public(ip):
                    found_public_ip = True
                    ip_used = ip
                    geo = GeoLocationService().geolocate_ip(ip)
                    if geo.asn and geo.asn != "Unknown":
                        asn = geo.asn
                        break
                    # Try whois lookup for ASN if geolocation fails
                    try:
                        w = whois.whois(ip)
                        if hasattr(w, 'asn') and w.asn:
                            asn = str(w.asn)
                            break
                        for field in ['org', 'netname', 'descr']:
                            val = getattr(w, field, None)
                            if val and isinstance(val, str) and 'AS' in val.upper():
                                m_asn = re.search(r'AS(\d+)', val.upper())
                                if m_asn:
                                    asn = f"AS{m_asn.group(1)}"
                                    break
                    except Exception:
                        pass
            # If no public IP found, try to resolve hostname to IP and use BGPView
            if not found_public_ip and host:
                try:
                    ip_from_host = socket.gethostbyname(host)
                    if IPValidator.is_valid_public(ip_from_host):
                        ip_used = ip_from_host
                        # Query BGPView API
                        try:
                            resp = requests.get(f'https://api.bgpview.io/ip/{ip_from_host}', timeout=5)
                            if resp.status_code == 200:
                                data = resp.json()
                                asn_data = data.get('data', {}).get('prefixes', [])
                                if asn_data:
                                    asn = asn_data[0].get('asn', {}).get('asn', 'Unknown')
                                    if asn != 'Unknown':
                                        asn = f"AS{asn}"
                        except Exception:
                            pass
                except Exception:
                    pass
            # Compose display string
            if not (found_public_ip or (ip_used and asn != "Unknown")):
                if decoded_location:
                    hops.append(f"{host or 'Unknown'} ({decoded_location}, Unknown)")
                else:
                    hops.append(f"{host or 'Unknown'} (Unknown)")
            else:
                if decoded_location:
                    hops.append(f"{host or 'Unknown'} ({decoded_location}, {asn})")
                else:
                    hops.append(f"{host or 'Unknown'} ({asn})")
        return "; ".join(hops) if hops else "N/A"

    def _is_shared_mail_infra(self, hop: Dict[str, Any], ip: Optional[str] = None) -> bool:
        """True for inbound shared MTAs (Microsoft / Google delivery path).

        Compact scans should not spend their geo budget on these — the sending
        mail-server exit (Sailthru, SES, etc.) matters more and is always whois'd.
        """
        inbound = ("microsoft", "google")
        org = f"{hop.get('asn_org') or ''} {(hop.get('ip_decode') or {}).get('org_hint') or ''}".lower()
        if any(p in org for p in inbound):
            return True
        for name in (hop.get("fqdn"), hop.get("by_host"), hop.get("helo")):
            n = str(name or "").lower()
            if any(
                s in n
                for s in (
                    "outlook.com", "office365.com", "protection.outlook.com",
                    "google.com", "gmail.com", "googlemail.com",
                )
            ):
                return True
        if ip and addr_decode is not None:
            dec = addr_decode.decode_ip(ip)
            if dec and dec.get("org_hint"):
                hint = str(dec["org_hint"]).lower()
                if any(p in hint for p in inbound):
                    return True
        return False

    @staticmethod
    def _blank_reg(*values: Any) -> bool:
        """True when a registration field is missing or a placeholder (NA / — / Unknown)."""
        for value in values:
            s = str(value or "").strip()
            if not s:
                continue
            if s.lower() in {"unknown", "n/a", "na", "none", "null", "—", "-", "skipped (compact)"}:
                continue
            return False
        return True

    def _whois_ip_fallback(self, ip: str) -> Dict[str, str]:
        """python-whois fallback when RDAP is thin (bounded so scans cannot hang)."""
        out = {"asn": "", "org": "", "prefix": "", "rir": "", "country": "", "city": ""}
        try:
            from concurrent.futures import ThreadPoolExecutor, TimeoutError as FuturesTimeout

            def _lookup() -> Any:
                return whois.whois(ip)

            with ThreadPoolExecutor(max_workers=1) as pool:
                fut = pool.submit(_lookup)
                try:
                    w = fut.result(timeout=8)
                except FuturesTimeout:
                    fut.cancel()
                    return out
        except Exception:
            return out
        if not w:
            return out
        asn = getattr(w, "asn", None) or getattr(w, "asn_cidr", None)
        if asn and not self._blank_reg(asn):
            asn_s = str(asn).split()[0].strip()
            if asn_s.isdigit():
                asn_s = f"AS{asn_s}"
            elif not asn_s.upper().startswith("AS") and "AS" in asn_s.upper():
                m = re.search(r"AS\d+", asn_s, re.I)
                asn_s = m.group(0).upper() if m else asn_s
            if not self._blank_reg(asn_s):
                out["asn"] = asn_s
        for key in ("org", "organization", "orgname", "descr", "netname"):
            val = getattr(w, key, None)
            if val:
                text = (val[0] if isinstance(val, list) else str(val)).strip()
                if text and not self._blank_reg(text):
                    out["org"] = text
                    break
        for key in ("asn_cidr", "cidr", "network"):
            cidr = getattr(w, key, None)
            if cidr:
                text = (cidr[0] if isinstance(cidr, list) else str(cidr)).strip()
                if text and "/" in text and not self._blank_reg(text):
                    out["prefix"] = text
                    break
        rir = getattr(w, "asn_registry", None) or getattr(w, "source", None)
        if rir and not self._blank_reg(rir):
            out["rir"] = str(rir).strip().upper()
        country = getattr(w, "country", None)
        if country:
            text = (country[0] if isinstance(country, list) else str(country)).strip()
            if text and not self._blank_reg(text):
                out["country"] = text
        city = getattr(w, "city", None)
        if city:
            text = (city[0] if isinstance(city, list) else str(city)).strip()
            if text and not self._blank_reg(text):
                out["city"] = text
        return out

    def _ripestat_prefix(self, ip: str) -> Dict[str, str]:
        """Prefix / origin ASN from RIPEstat when RDAP returns NA (common for M365 IPv6)."""
        out = {"prefix": "", "asn": "", "rir": ""}
        if route_checks is None:
            return out
        try:
            # network-info only — avoid the heavier peer/RPKI fan-out in bgp_origin_info.
            net = route_checks._ripestat("network-info", ip)  # noqa: SLF001
        except Exception:
            return out
        if not net:
            return out
        prefix = str(net.get("prefix") or "").strip()
        if prefix and "/" in prefix and not self._blank_reg(prefix):
            out["prefix"] = prefix
        asns = net.get("asns") or []
        if asns:
            asn0 = str(asns[0]).strip()
            if asn0.isdigit():
                out["asn"] = f"AS{asn0}"
            elif asn0.upper().startswith("AS"):
                out["asn"] = asn0.upper()
        return out

    def _fill_ip_registration(self, hop: Dict[str, Any], public_ip: str, *, force_whois: bool = False) -> None:
        """RDAP → RIPEstat → WHOIS until ASN / prefix / RIR / org are filled.

        Prefers registry allocation (registrant / netname / hostname map) over
        BGP transit ASN so Sailthru-on-Amazon does not show up as Amazon-only.
        """
        self._apply_rdap_fields(hop, public_ip)
        need_prefix = self._blank_reg(hop.get("prefix"))
        need_asn = self._blank_reg(hop.get("asn"))
        if need_prefix or need_asn or self._blank_reg(hop.get("rir")) or self._blank_reg(hop.get("asn_org")):
            ripe = self._ripestat_prefix(public_ip)
            if ripe.get("prefix") and self._blank_reg(hop.get("prefix")):
                hop["prefix"] = ripe["prefix"]
            if ripe.get("asn") and self._blank_reg(hop.get("asn")):
                hop["asn"] = ripe["asn"]
        still_thin = self._blank_reg(hop.get("prefix")) or self._blank_reg(hop.get("asn"))
        if force_whois or still_thin:
            w = self._whois_ip_fallback(public_ip)
            if w.get("asn") and self._blank_reg(hop.get("asn")):
                hop["asn"] = w["asn"]
            if w.get("org") and self._blank_reg(hop.get("asn_org")):
                hop["asn_org"] = w["org"]
            if w.get("prefix") and self._blank_reg(hop.get("prefix")):
                hop["prefix"] = w["prefix"]
            if w.get("rir") and self._blank_reg(hop.get("rir")):
                hop["rir"] = w["rir"]
            if w.get("country") and self._blank_reg(hop.get("country")):
                hop["country"] = w["country"]
                if not hop.get("location_source"):
                    hop["location_source"] = "IP WHOIS registry"
                    hop["geo_confidence"] = "low"
            if w.get("city") and self._blank_reg(hop.get("city")):
                hop["city"] = w["city"]

        # Hostname map (mta*.sailthru.com → AS18877) beats BGP transit ASN.
        if addr_decode is not None and hop.get("hop_kind") == "public":
            for name in (hop.get("fqdn"), hop.get("helo"), hop.get("by_host")):
                asn_h, org_h = addr_decode.provider_asn_for_host(name)
                if not asn_h:
                    continue
                bgp = str(hop.get("asn_bgp") or hop.get("asn") or "").strip().upper()
                hop["asn"] = asn_h
                if self._blank_reg(hop.get("asn_org")) and org_h:
                    hop["asn_org"] = org_h
                if bgp and bgp != asn_h.upper():
                    org_bgp = str(hop.get("org_bgp") or "").strip()
                    note = f"BGP transit {bgp}" + (f" ({org_bgp})" if org_bgp else "")
                    lines = list(hop.get("ip_decode_lines") or [])
                    if note not in lines:
                        lines.append(note)
                    hop["ip_decode_lines"] = lines
                break

        if hop.get("rdap_netname"):
            lines = list(hop.get("ip_decode_lines") or [])
            note = f"Allocated netblock {hop['rdap_netname']}"
            if note not in lines:
                lines.insert(0, note)
            hop["ip_decode_lines"] = lines

        # Normalise leftover placeholders so the table never shows literal "NA".
        if self._blank_reg(hop.get("prefix")):
            hop["prefix"] = "—"
        if self._blank_reg(hop.get("rir")):
            hop["rir"] = "—"

    def _enrich_public_hop(
        self, hop: Dict[str, Any], public_ip: str, *, lite: bool, deep_lookups: bool
    ) -> None:
        """Geo + reputation for a public hop; reverse-DNS when deep_lookups (or exit).

        Must run *after* ``_fill_ip_registration`` so BGP geo ASN/org (Amazon) does
        not overwrite the allocation decode (Sailthru).
        """
        geo = GeoLocationService().geolocate_ip(public_ip)
        blank = {None, "", "Unknown", "—", "Skipped (compact)", "NA", "N/A"}
        if hop.get("asn") in blank or self._blank_reg(hop.get("asn")):
            if geo.asn and not self._blank_reg(geo.asn):
                hop["asn"] = geo.asn
        if hop.get("asn_org") in blank or self._blank_reg(hop.get("asn_org")):
            if geo.org and not self._blank_reg(geo.org):
                hop["asn_org"] = geo.org
        filled_geo = False
        if self._blank_reg(hop.get("country")) and geo.country and not self._blank_reg(geo.country):
            hop["country"] = geo.country
            filled_geo = True
        elif self._blank_reg(hop.get("country")):
            hop["country"] = "Unknown"
        if self._blank_reg(hop.get("city")) and geo.city not in blank | {"Geolocation Unavailable"}:
            hop["city"] = geo.city
            filled_geo = True
        elif self._blank_reg(hop.get("city")):
            hop["city"] = "Unknown"
        if filled_geo and not hop.get("location_source"):
            hop["location_source"] = (
                f"IP geolocation via {geo.geo_source or 'lookup'} — approximate"
            )
            hop["geo_confidence"] = "low"
        if hop.get("lat") is None:
            hop["lat"] = geo.lat
            hop["lon"] = geo.lon
        if geo.abuse_confidence is not None:
            hop["reputation"] = f"Score: {geo.abuse_confidence}"
        elif hop.get("reputation") in blank | {"N/A"}:
            hop["reputation"] = "N/A"
        if deep_lookups or hop.get("exit_hop"):
            ptr = ""
            try:
                ptr = socket.gethostbyaddr(public_ip)[0]
            except Exception:
                ptr = ""
            if ptr and not self._blank_reg(ptr):
                hop["reverse_dns"] = ptr
            else:
                # Fall back to Received-header hostname for this IP (often the PTR).
                for name in (hop.get("fqdn"), hop.get("helo"), hop.get("by_host")):
                    n = str(name or "").strip().rstrip(".")
                    if not n or addr_decode is None:
                        continue
                    if addr_decode.ip_literal(n):
                        continue
                    if "." in n:
                        hop["reverse_dns"] = n
                        break
                else:
                    if hop.get("reverse_dns") in blank or self._blank_reg(hop.get("reverse_dns")):
                        hop["reverse_dns"] = "Unknown"

    def _ensure_exit_hop_whois(self, hops: List[Dict[str, Any]], *, lite: bool) -> None:
        """Fully register the sending exit; on full/deep, fill thin hops too."""
        for hop in hops:
            if hop.get("hop_kind") != "public":
                continue
            raw_ip = str(hop.get("ip") or "").split()[0]
            if not IPValidator.is_valid_public(raw_ip):
                continue
            is_shared = self._is_shared_mail_infra(hop, raw_ip)
            if is_shared and lite:
                # Compact: still repair placeholder RDAP (Prefix=NA) cheaply.
                if self._blank_reg(hop.get("prefix")) or self._blank_reg(hop.get("rir")):
                    self._fill_ip_registration(hop, raw_ip, force_whois=False)
                continue
            if not is_shared:
                hop["exit_hop"] = True
            thin = (
                self._blank_reg(hop.get("country"))
                or self._blank_reg(hop.get("asn"))
                or hop.get("reputation") == "Skipped (compact)"
                or self._blank_reg(hop.get("prefix"))
            )
            if thin or hop.get("exit_hop") or not lite:
                # Registration first so geo BGP (Amazon) cannot mask allocation (Sailthru).
                self._fill_ip_registration(
                    hop, raw_ip, force_whois=bool(hop.get("exit_hop") or not lite)
                )
                self._enrich_public_hop(
                    hop, raw_ip, lite=lite, deep_lookups=(not lite) or bool(hop.get("exit_hop"))
                )
            if hop.get("exit_hop"):
                break
        # Full / deep: WHOIS any remaining public hop still missing prefix/ASN.
        if not lite:
            for hop in hops:
                if hop.get("hop_kind") != "public":
                    continue
                raw_ip = str(hop.get("ip") or "").split()[0]
                if not IPValidator.is_valid_public(raw_ip):
                    continue
                if (
                    self._blank_reg(hop.get("prefix"))
                    or self._blank_reg(hop.get("asn"))
                    or self._blank_reg(hop.get("rir"))
                    or self._blank_reg(hop.get("reverse_dns"))
                    or self._blank_reg(hop.get("country"))
                    or self._blank_reg(hop.get("asn_org"))
                ):
                    self._fill_ip_registration(hop, raw_ip, force_whois=True)
                    self._enrich_public_hop(hop, raw_ip, lite=False, deep_lookups=True)

    def _ipwhois_lookup(self, ip: str) -> dict:
        """RDAP lookup: allocation (registrant/net) preferred over BGP transit ASN.

        ipwhois ``asn`` / ``asn_description`` are often the *announcing* ASN
        (e.g. Amazon AS14618 for a Sailthru block). The ARIN network + registrant
        entity is the proper owner decode (Sailthru, NET-192-64-236-0-1).
        """
        cached = _RDAP_CACHE.get(ip)
        if cached is not None:
            return cached
        empty = {
            "asn": "", "asn_bgp": "", "org": "", "org_bgp": "", "org_alloc": "",
            "prefix": "", "rir": "", "netname": "", "country": "", "city": "",
        }
        try:
            data = IPWhois(ip).lookup_rdap(depth=1)
            asn_bgp = ""
            asn_raw = data.get("asn")
            if not self._blank_reg(asn_raw):
                asn_bgp = str(asn_raw).strip()
                if asn_bgp.isdigit():
                    asn_bgp = f"AS{asn_bgp}"
                elif not asn_bgp.upper().startswith("AS"):
                    asn_bgp = f"AS{asn_bgp}"
                if self._blank_reg(asn_bgp):
                    asn_bgp = ""

            asn_desc = data.get("asn_description")
            org_bgp = ""
            if asn_desc is not None and not self._blank_reg(str(asn_desc)):
                org_bgp = str(asn_desc).split(" | ")[0].strip()

            nir_val = data.get("nir") or data.get("asn_registry") or ""
            if self._blank_reg(nir_val):
                nir_val = ""

            network = data.get("network") if isinstance(data.get("network"), dict) else {}
            netname = ""
            if network:
                nn = network.get("name") or ""
                if not self._blank_reg(nn):
                    netname = str(nn).strip()

            # Prefer allocation CIDR (e.g. 192.64.236.0/22) over BGP more-specific.
            prefix = ""
            if network:
                cidr = network.get("cidr") or ""
                if not self._blank_reg(cidr) and "/" in str(cidr):
                    prefix = str(cidr).split(",")[0].strip()
            if not prefix:
                bgp_cidr = data.get("asn_cidr") or ""
                if not self._blank_reg(bgp_cidr) and "/" in str(bgp_cidr):
                    prefix = str(bgp_cidr).strip()

            org_alloc, country, city = self._rdap_registrant_geo(data)

            # Display ASN stays BGP until hostname/allocation hints override later;
            # organisation prefers registrant / netname over Amazon transit label.
            org = org_alloc or (netname.title() if netname else "") or org_bgp
            result = {
                "asn": asn_bgp,
                "asn_bgp": asn_bgp,
                "org": org,
                "org_bgp": org_bgp,
                "org_alloc": org_alloc,
                "prefix": prefix,
                "rir": str(nir_val).strip().upper() if nir_val else "",
                "netname": netname,
                "country": country,
                "city": city,
            }
            _RDAP_CACHE[ip] = result
            return result
        except (ipwhois_exceptions.IPDefinedError, ipwhois_exceptions.HTTPLookupError, Exception):
            _RDAP_CACHE[ip] = empty
            return empty

    @staticmethod
    def _rdap_registrant_geo(data: Dict[str, Any]) -> Tuple[str, str, str]:
        """Pull registrant name + city/country from RDAP entity contacts."""
        objects = data.get("objects") if isinstance(data.get("objects"), dict) else {}
        entity_ids = list(data.get("entities") or [])
        # Prefer explicit registrant handles, then any object with role registrant.
        ordered: List[str] = []
        for hid in entity_ids:
            if hid and hid not in ordered:
                ordered.append(str(hid))
        for hid, obj in objects.items():
            roles = [str(r).lower() for r in (obj.get("roles") or [])] if isinstance(obj, dict) else []
            if "registrant" in roles and hid not in ordered:
                ordered.insert(0, str(hid))

        def _parse_addr(val: Any) -> Tuple[str, str]:
            text = ""
            if isinstance(val, list) and val:
                first = val[0]
                if isinstance(first, dict):
                    text = str(first.get("value") or "")
                else:
                    text = str(first)
            elif isinstance(val, dict):
                text = str(val.get("value") or "")
            elif val:
                text = str(val)
            lines = [ln.strip() for ln in text.replace("\r", "").split("\n") if ln.strip()]
            if not lines:
                return "", ""
            country = ""
            city = ""
            # Typical ARIN: street / city / ST / zip / Country
            if len(lines) >= 2 and re.search(r"[A-Za-z]", lines[-1]):
                country_line = lines[-1]
                if re.search(r"united states|u\.?s\.?a\.?", country_line, re.I):
                    country = "US"
                elif len(country_line) == 2 and country_line.isalpha():
                    country = country_line.upper()
                else:
                    country = country_line
            for i, ln in enumerate(lines):
                if re.fullmatch(r"[A-Z]{2}", ln) and i > 0:
                    city = lines[i - 1]
                    break
            if not city and len(lines) >= 3:
                # city often second line when street is first
                if not re.search(r"\d", lines[1]):
                    city = lines[1]
            return city, country

        for hid in ordered:
            obj = objects.get(hid)
            if not isinstance(obj, dict):
                continue
            roles = [str(r).lower() for r in (obj.get("roles") or [])]
            if roles and "registrant" not in roles and hid not in entity_ids[:1]:
                continue
            contact = obj.get("contact") if isinstance(obj.get("contact"), dict) else {}
            name = str(contact.get("name") or "").strip()
            city, country = _parse_addr(contact.get("address"))
            if name or country or city:
                return name, country, city
        return "", "", ""

    def _apply_rdap_fields(self, hop: Dict[str, Any], public_ip: str) -> None:
        """Fill ASN / organisation / prefix / RIR / registry geo from RDAP."""
        info = self._ipwhois_lookup(public_ip)
        if info.get("asn") and self._blank_reg(hop.get("asn")):
            hop["asn"] = info["asn"]
        # Allocation registrant always beats BGP transit org (Amazon hosting Sailthru).
        if info.get("org_alloc"):
            hop["asn_org"] = info["org_alloc"]
        elif info.get("org") and self._blank_reg(hop.get("asn_org")):
            hop["asn_org"] = info["org"]
        if info.get("prefix") and (
            self._blank_reg(hop.get("prefix"))
            or (
                info.get("netname")
                and str(hop.get("prefix") or "").count(".") == 3
                and "/" in str(info["prefix"])
            )
        ):
            # Prefer allocation prefix when we have a named netblock.
            if info.get("netname") or self._blank_reg(hop.get("prefix")):
                hop["prefix"] = info["prefix"]
        if info.get("rir") and self._blank_reg(hop.get("rir")):
            hop["rir"] = info["rir"]
        if info.get("country") and self._blank_reg(hop.get("country")):
            hop["country"] = info["country"]
            if not hop.get("location_source"):
                hop["location_source"] = "IP RDAP registrant"
                hop["geo_confidence"] = "medium"
        if info.get("city") and self._blank_reg(hop.get("city")):
            hop["city"] = info["city"]
        if info.get("netname"):
            hop["rdap_netname"] = info["netname"]
        if info.get("asn_bgp"):
            hop["asn_bgp"] = info["asn_bgp"]
        if info.get("org_bgp"):
            hop["org_bgp"] = info["org_bgp"]

    @staticmethod
    def _apply_address_and_host_decoding(
        hop: Dict[str, Any],
        ip_literal: Optional[str],
        provider_ctx: set,
        public_ip: Optional[str],
        geo_via_embedded: Optional[str],
    ) -> None:
        """Attach offline IP decoding and hostname-derived location to a hop."""
        hop.setdefault('ip_decode_lines', [])
        hop.setdefault('host_location', None)
        hop.setdefault('location_source', '')
        if addr_decode is None:
            return
        unknown_vals = {None, '', 'Unknown', '—', 'Skipped (compact)', 'NA', 'N/A',
                        'Internal / local hop', 'Insufficient header data'}

        if ip_literal:
            dec = addr_decode.decode_ip(ip_literal, provider_context=provider_ctx)
            if dec:
                lines = list(dec.get('lines') or [])
                if geo_via_embedded and public_ip:
                    lines.append(f"Geolocated via embedded IPv4 {public_ip}")
                hop['ip_decode'] = dec
                hop['ip_decode_lines'] = lines
                if dec.get('scope') == 'suspicious':
                    hop['address_flag'] = 'suspicious'
                if hop.get('hop_kind') != 'public' and dec.get('scope') == 'private':
                    hop['hop_kind'] = 'internal'
                    if str(hop.get('ip') or '') in {'Unknown', 'Internal', 'localhost / internal'}:
                        hop['ip'] = f"{ip_literal} (internal)"
                    hop['reputation'] = 'N/A (internal)'
                    for key in ('asn', 'prefix', 'rir', 'reverse_dns'):
                        if hop.get(key) in unknown_vals:
                            hop[key] = '—'
                if dec.get('org_hint') and hop.get('asn_org') in unknown_vals:
                    hop['asn_org'] = dec['org_hint']
                elif hop.get('hop_kind') == 'internal' and hop.get('asn_org') in unknown_vals:
                    hop['asn_org'] = f"Private network ({dec.get('kind')})"
                if dec.get('asn_hint') and hop.get('asn') in unknown_vals:
                    hop['asn'] = dec['asn_hint']

        def _host_loc(name: Any) -> Optional[Dict[str, Any]]:
            if name and not addr_decode.ip_literal(str(name)):
                return addr_decode.decode_host_location(str(name))
            return None

        # The bracketed IP belongs to the "from" side; the by-host only tells us
        # where this hop's timestamp was written.
        loc = _host_loc(hop.get('fqdn')) or _host_loc(hop.get('helo'))
        hop['by_location'] = _host_loc(hop.get('by_host'))

        # Fill ASN/org from hostname when geo/RDAP left them blank — public hops only
        # (do not stamp Sailthru ASN onto RFC1918 internal relays).
        if hop.get("hop_kind") == "public":
            for name in (hop.get("fqdn"), hop.get("by_host"), hop.get("helo")):
                asn_h, org_h = addr_decode.provider_asn_for_host(name)
                if asn_h and hop.get("asn") in unknown_vals:
                    hop["asn"] = asn_h
                if org_h and hop.get("asn_org") in unknown_vals:
                    hop["asn_org"] = org_h
                if asn_h or org_h:
                    break

        if not loc:
            return
        hop['host_location'] = loc
        conf = loc.get('confidence', 'low')
        org_text = f"{hop.get('asn_org') or ''} {(hop.get('ip_decode') or {}).get('org_hint') or ''}".lower()
        big_provider = any(p in org_text for p in ('microsoft', 'google', 'amazon', 'mimecast'))
        registry_geo = f"{hop.get('country')}/{hop.get('city')}"
        approx_geo = hop.get('geo_confidence') == 'low'
        override = (
            hop.get('country') not in unknown_vals
            and (big_provider or approx_geo)
            and conf in {'high', 'medium'}
            and loc.get('country') not in {None, '', 'EU'}
            and (loc.get('country') != hop.get('country') or approx_geo)
        )
        if hop.get('country') in unknown_vals or override:
            hop['country'] = loc.get('country') or 'Unknown'
            hop['city'] = loc.get('city') or '—'
            hop['location_source'] = f"from host name: {loc.get('detail', '')} ({conf} confidence)"
            if override:
                hop['location_source'] += f"; IP geolocation said {registry_geo}"
                hop.pop('geo_confidence', None)
            if loc.get('lat') is not None and (override or hop.get('lat') is None):
                hop['lat'] = loc['lat']
                hop['lon'] = loc['lon']

    def _extract_hop_details(self, headers: Optional[str], *, lite: bool = False) -> list:
        import socket
        import email.utils
        if not headers:
            return []

        received_lines = HeaderParser.extract_received_blocks(headers)
        hops = []
        prev_time = None
        public_enriched = 0
        max_public_lite = 3
        joined_received = "\n".join(received_lines).lower()
        provider_ctx = set()
        if "google.com" in joined_received:
            provider_ctx.add("google")
        if "outlook.com" in joined_received or "office365.com" in joined_received:
            provider_ctx.add("microsoft")

        for line in received_lines:
            hop = {
                'fqdn': None,
                'by_host': None,
                'helo': None,
                'protocol': None,
                'decoded_location': None,
                'ip': 'Unknown',
                'asn': 'Unknown',
                'asn_org': 'Unknown',
                'prefix': '—',
                'rir': '—',
                'country': 'Unknown',
                'city': 'Unknown',
                'lat': None,
                'lon': None,
                'reverse_dns': 'Unknown',
                'reputation': 'Unknown',
                'timestamp': 'Unknown',
                'delay': 'N/A',
                'bgp_peers': [],
                'hop_kind': 'unknown',  # public | internal | unknown
                'notes': '',
            }

            # from / by / with / helo (full unfolded Received line)
            m_from = re.search(
                r'\bfrom\s+([^\s\(\];]+)(?:\s+\(([^)]*)\))?',
                line,
                re.IGNORECASE,
            )
            if m_from:
                hop['fqdn'] = m_from.group(1).strip().strip('"')
                paren = (m_from.group(2) or "").strip()
                if paren:
                    # HELO/EHLO often appears inside the parenthetical
                    m_helo = re.search(
                        r'(?:helo|ehlo)\s*=?\s*([^\s\);]+)',
                        paren,
                        re.IGNORECASE,
                    )
                    if m_helo:
                        hop['helo'] = m_helo.group(1).strip()

            m_by = re.search(r'\bby\s+([^\s\(\];]+)', line, re.IGNORECASE)
            if m_by:
                hop['by_host'] = m_by.group(1).strip().strip('"')

            m_with = re.search(r'\bwith\s+([^\s\(\];]+)', line, re.IGNORECASE)
            if m_with:
                hop['protocol'] = m_with.group(1).strip()

            if not hop['helo']:
                m_helo2 = re.search(
                    r'(?:helo|ehlo)\s*=?\s*([^\s\);]+)',
                    line,
                    re.IGNORECASE,
                )
                if m_helo2:
                    hop['helo'] = m_helo2.group(1).strip()

            from_desc = self._decode_m365_hostname(hop['fqdn']) if hop['fqdn'] else None
            by_desc = (
                self._decode_m365_hostname(hop['by_host'])
                if hop['by_host'] and hop['by_host'] != hop['fqdn']
                else None
            )
            if from_desc and by_desc:
                hop['decoded_location'] = f"from: {from_desc} → by: {by_desc}"
            elif by_desc:
                hop['decoded_location'] = f"by: {by_desc}"
            else:
                hop['decoded_location'] = from_desc

            # Collect candidate IPs (bracketed + client-ip=)
            candidates: List[str] = re.findall(r'\[([\d\.:a-fA-F]+)\]', line)
            client_ip = re.search(r'\bclient-ip=([\d\.:a-fA-F]+)', line, re.IGNORECASE)
            if client_ip:
                candidates.append(client_ip.group(1))
            # Exchange Online writes "from HOST (2603:10a6:…) by …" — no brackets.
            for paren_ip in re.findall(r'\(([0-9a-fA-F:.]{7,45})\)', line):
                if IPValidator.is_valid_public(paren_ip) or IPValidator.is_private(paren_ip):
                    if paren_ip not in candidates:
                        candidates.append(paren_ip)
            # Gmail and others write bare IP literals as the from/by host
            # ("by 2002:a05:7300:…") — those are addresses too.
            if addr_decode is not None:
                for tok in (hop['fqdn'], hop['by_host'], hop['helo']):
                    lit = addr_decode.ip_literal(tok)
                    if lit and lit not in candidates:
                        candidates.append(lit)

            public_ip = None
            private_ip = None
            decode_ip_literal = None
            geo_via_embedded = None
            for ip in candidates:
                if IPValidator.is_valid_public(ip):
                    public_ip = ip
                    decode_ip_literal = ip
                    break
                dec = addr_decode.decode_ip(ip, provider_context=provider_ctx) if addr_decode else None
                if dec and dec.get('geo_ip') and IPValidator.is_valid_public(dec['geo_ip']):
                    # 6to4 / Teredo / NAT64 / IPv4-mapped carrying a public IPv4.
                    public_ip = dec['geo_ip']
                    decode_ip_literal = ip
                    geo_via_embedded = ip
                    break
                if IPValidator.is_private(ip) and not private_ip:
                    private_ip = ip
                if not decode_ip_literal:
                    decode_ip_literal = ip

            # Resolve hostname only when it looks like a real host (not localhost/mail@…)
            resolve_name = hop['fqdn'] or hop['by_host']
            if (
                not public_ip
                and resolve_name
                and '@' not in resolve_name
                and resolve_name.lower() not in {'localhost', 'local', 'unknown', 'mail'}
                and '.' in resolve_name
            ):
                try:
                    ip_from_host = socket.gethostbyname(resolve_name)
                    if IPValidator.is_valid_public(ip_from_host):
                        public_ip = ip_from_host
                    elif IPValidator.is_private(ip_from_host) and not private_ip:
                        private_ip = ip_from_host
                except Exception:
                    pass

            local_tokens = {
                'localhost', 'local', '127.0.0.1', '::1', 'mail', 'mail@localhost',
            }
            from_l = (hop['fqdn'] or '').lower()
            by_l = (hop['by_host'] or '').lower()
            proto_l = (hop['protocol'] or '').lower()

            def _looks_internal_host(name: str) -> bool:
                if not name:
                    return False
                if name in local_tokens:
                    return True
                if '@localhost' in name or name.endswith('.localhost'):
                    return True
                if name.endswith('.local') or name.endswith('.internal'):
                    return True
                if name.startswith('mail@'):
                    return True
                return False

            is_local_name = (
                _looks_internal_host(from_l)
                or _looks_internal_host(by_l)
                or 'local' in proto_l
            )

            note_parts: List[str] = []
            if hop['helo']:
                note_parts.append(f"helo {hop['helo']}")
            if hop['protocol']:
                note_parts.append(f"with {hop['protocol']}")

            if public_ip:
                hop['hop_kind'] = 'public'
                hop['ip'] = geo_via_embedded or public_ip
                # Compact: don't burn geo budget on shared delivery MTAs
                # (Microsoft/Google/…). Prefer the sending exit (e.g. Sailthru).
                # RDAP still runs for every public hop; exit hop is force-whois'd
                # after reverse.
                is_shared = self._is_shared_mail_infra(hop, public_ip)
                if not lite:
                    do_enrich = True
                elif is_shared:
                    do_enrich = False
                else:
                    do_enrich = public_enriched < max_public_lite
                # RDAP allocation first; geo enrichment second (must not mask Sailthru with Amazon).
                self._fill_ip_registration(
                    hop, public_ip, force_whois=(not lite and not is_shared)
                )
                if do_enrich:
                    public_enriched += 1
                    self._enrich_public_hop(
                        hop, public_ip, lite=lite, deep_lookups=not lite
                    )
                    if not lite and hop.get("reputation") in (None, "", "Unknown", "N/A"):
                        try:
                            api_key = _resolve_abuseipdb_api_key()
                            if api_key:
                                abuse_headers = {
                                    "Key": api_key,
                                    "Accept": "application/json",
                                }
                                _rate_limit_abuseipdb()
                                resp, _insecure = _http_get_with_retry(
                                    abuseipdb_check_url(),
                                    headers=abuse_headers,
                                    params={"ipAddress": public_ip, "maxAgeInDays": 90},
                                    timeout=12,
                                    retries=2,
                                )
                                if resp.status_code == 200:
                                    data = resp.json().get("data", {})
                                    hop["reputation"] = (
                                        f"Score: {data.get('abuseConfidenceScore', 'N/A')}"
                                    )
                            else:
                                hop["reputation"] = "N/A (no API key)"
                        except Exception:
                            hop["reputation"] = "Unknown"
                else:
                    hop["reputation"] = hop.get("reputation") or "Skipped (compact)"
            elif private_ip or is_local_name:
                hop['hop_kind'] = 'internal'
                if private_ip:
                    hop['ip'] = f"{private_ip} (internal)"
                elif is_local_name:
                    hop['ip'] = 'localhost / internal'
                else:
                    hop['ip'] = 'Internal'
                hop['asn'] = '—'
                hop['asn_org'] = 'Internal / local hop'
                hop['prefix'] = '—'
                hop['rir'] = '—'
                hop['country'] = '—'
                hop['city'] = '—'
                hop['reverse_dns'] = '—'
                hop['reputation'] = 'N/A (internal)'
                if not hop['fqdn'] and hop['by_host']:
                    hop['fqdn'] = hop['by_host']
                    note_parts.insert(0, 'receiving MTA (no from host)')
                elif not hop['fqdn']:
                    hop['fqdn'] = 'Internal hop'
            else:
                hop['hop_kind'] = 'unknown'
                hop['ip'] = 'Unknown'
                if not hop['fqdn'] and hop['by_host']:
                    hop['fqdn'] = hop['by_host']
                    note_parts.insert(0, 'receiving MTA (no from host)')
                elif not hop['fqdn']:
                    hop['fqdn'] = '(no from host in Received)'
                hop['asn_org'] = 'Insufficient header data'
                hop['reputation'] = 'N/A'

            self._apply_address_and_host_decoding(
                hop, decode_ip_literal, provider_ctx, public_ip, geo_via_embedded
            )

            hop['notes'] = '; '.join(note_parts)
            if hop['decoded_location'] is None and hop['notes']:
                hop['decoded_location'] = hop['notes']
            elif hop['notes'] and hop['decoded_location']:
                hop['decoded_location'] = f"{hop['decoded_location']} | {hop['notes']}"
            elif hop['notes']:
                hop['decoded_location'] = hop['notes']

            m_time = re.search(r';\s*(.+)$', line)
            if m_time:
                try:
                    raw_time = m_time.group(1)
                    dt = email.utils.parsedate_to_datetime(raw_time)
                    if dt.tzinfo is None or dt.tzinfo.utcoffset(dt) is None:
                        dt = pytz.utc.localize(dt)
                    # RFC 5322: "-0000" means "offset unknown", distinct from +0000.
                    if re.search(r'[+-]\d{4}', raw_time) and '-0000' not in raw_time:
                        hop['utc_offset_min'] = int(dt.utcoffset().total_seconds() // 60)
                    hop['timestamp_utc'] = dt.astimezone(pytz.utc).isoformat()
                    hop['timestamp'] = dt.strftime('%Y-%m-%d %H:%M:%S')
                    if prev_time:
                        delay = (dt - prev_time).total_seconds()
                        hop['delay'] = f"{delay:.2f}s"
                    prev_time = dt
                except Exception:
                    hop['timestamp'] = 'Unknown'

            hops.append(hop)

        # Received headers are newest-first; reverse so hop 1 is the origin /
        # starting point and the last hop is delivery into the mailbox.
        hops.reverse()
        prev_dt = None
        for hop in hops:
            hop['delay'] = 'N/A'
            t_iso = hop.get('timestamp_utc')
            if not t_iso:
                continue
            try:
                dt = datetime.fromisoformat(str(t_iso))
            except ValueError:
                continue
            if prev_dt is not None:
                hop['delay'] = f"{(dt - prev_dt).total_seconds():.2f}s"
            prev_dt = dt
        # Always WHOIS/RDAP the sending mail-server exit (first non-shared public hop).
        self._ensure_exit_hop_whois(hops, lite=lite)
        return hops

    def _strip_subdomains_until_resolvable(self, hostname: str) -> str:
        """Strip subdomains from left to right until a resolvable domain is found."""
        if not hostname or '.' not in hostname:
            return hostname
        
        parts = hostname.split('.')
        if len(parts) < 2:
            return hostname
        
        # Try from shortest to longest (right to left)
        for i in range(len(parts) - 1, 0, -1):
            test_domain = '.'.join(parts[i:])
            try:
                socket.gethostbyname(test_domain)
                return test_domain
            except (socket.gaierror, socket.herror):
                continue
        
        # If nothing resolves, return the original
        return hostname

    def _clean_organization_name(self, org_name: str) -> str:
        """Clean organization name by removing any AS numbers and common suffixes."""
        if not org_name or org_name == 'Unknown':
            return org_name
        
        original = org_name
        self.logger.debug(f"Cleaning org name: '{original}'")
        
        # Remove 'AS' followed by numbers and optional whitespace at the start
        org_name = re.sub(r'^AS\d+\s*', '', org_name, flags=re.IGNORECASE)
        self.logger.debug(f"After removing start AS: '{org_name}'")
        
        # Remove 'AS' followed by numbers anywhere else in the string
        org_name = re.sub(r'AS\d+', '', org_name, flags=re.IGNORECASE)
        self.logger.debug(f"After removing all AS: '{org_name}'")
        
        # Remove common suffixes
        suffixes_to_remove = [
            ' Limited', ' Ltd', ' LLC', ' Inc', ' Corporation', ' Corp', ' Company', ' Co',
            ' International', ' Intl', ' Technologies', ' Tech', ' Networks', ' Network',
            ' Communications', ' Comm', ' Services', ' Service', ' Solutions', ' Solution', ' LLC.', ' Inc.'
        ]
        for suffix in suffixes_to_remove:
            if org_name.endswith(suffix):
                org_name = org_name[:-len(suffix)]
                self.logger.debug(f"After removing suffix '{suffix}': '{org_name}'")
                break
        
        # Clean up extra whitespace and punctuation
        org_name = org_name.strip(' ,.-')
        self.logger.debug(f"Final cleaned result: '{org_name}'")
        
        return org_name if org_name else 'Unknown'

    def generate_outlook_signature(self, data: Dict[str, Any]) -> str:
        """Generate a simplified, Outlook-compatible HTML signature."""
        score = data['security_assessment']['score']
        risk = data['security_assessment']['risk_level']
        risk_bg = data['security_assessment']['risk_color']
        
        # Simple text-based signature for Outlook
        signature = f"""
<div style="font-family: Arial, sans-serif; font-size: 10px; color: #333; border-top: 1px solid #ccc; padding-top: 5px; margin-top: 10px;">
<div style="background-color: {risk_bg}; color: white; font-weight: bold; font-size: 11px; padding: 2px 4px; text-align: center; margin-bottom: 5px;">
EMAIL SECURITY ANALYSIS - RISK {risk} ({score}/100)
</div>
<div style="margin-bottom: 3px;"><strong>Sender:</strong> {data['sender_email']}</div>
<div style="margin-bottom: 3px;"><strong>Domain:</strong> {data['sender_domain']}</div>
<div style="margin-bottom: 3px;"><strong>IP:</strong> {data['sender_ip']}</div>
<div style="margin-bottom: 3px;"><strong>Location:</strong> {data['location']}</div>
<div style="margin-bottom: 3px;"><strong>Organization:</strong> {data['organization']}</div>
<div style="margin-bottom: 3px;"><strong>Auth:</strong> {data['auth_type']}</div>
<div style="margin-bottom: 3px;"><strong>SPF:</strong> {data['spf'][1] if isinstance(data['spf'], tuple) else data['spf']}</div>
<div style="margin-bottom: 3px;"><strong>DKIM:</strong> {data['dkim'][1] if isinstance(data['dkim'], tuple) else data['dkim']}</div>
<div style="margin-bottom: 3px;"><strong>DMARC:</strong> {data['dmarc'][1] if isinstance(data['dmarc'], tuple) else data['dmarc']}</div>
<div style="margin-bottom: 3px;"><strong>Blocklist:</strong> {data['security_flags']}</div>
<div style="font-size: 9px; color: #666; margin-top: 5px; border-top: 1px solid #eee; padding-top: 3px;">
GURI: {data.get('guri', 'N/A')} | (C) Aliniant Labs 2025 | Created: {data.get('created', 'N/A')}
</div>
</div>
"""
        return signature

def _notify_guri_scan_finished() -> None:
    """Tell a running GURI that AES just finished a scan (Push mode)."""
    try:
        from PySide6.QtCore import QCoreApplication
        from PySide6.QtNetwork import QLocalSocket

        app = QCoreApplication.instance()
        if app is None:
            app = QCoreApplication([])
        sock = QLocalSocket()
        sock.connectToServer("GeoFooter_GURI_GUI_v1")
        if not sock.waitForConnected(300):
            return
        sock.write(b"PUSH\n")
        sock.flush()
        sock.waitForBytesWritten(300)
        sock.disconnectFromServer()
    except Exception:
        return


def main():
    """Main execution function."""
    import sys
    import os
    
    # Setup logging
    LoggingSetup.setup()
    logger = logging.getLogger(__name__)
    
    # Send debug output to stderr so stdout only contains the footer file path
    print("Script starting...", file=sys.stderr)
    print(f"Arguments: {sys.argv}", file=sys.stderr)
    
    if len(sys.argv) < 2 or len(sys.argv) > 4:
        print("Usage: python geolocate_headers.py <header_file> [output_file] [compact|full|deep]", file=sys.stderr)
        sys.exit(1)
    
    header_file = sys.argv[1]
    output_file = None
    footer_mode = "compact"

    if len(sys.argv) == 3:
        if sys.argv[2].lower() in {"compact", "full", "deep"}:
            footer_mode = sys.argv[2].lower()
        else:
            output_file = sys.argv[2]
    elif len(sys.argv) == 4:
        output_file = sys.argv[2]
        footer_mode = sys.argv[3].lower()

    if footer_mode not in {"compact", "full", "deep"}:
        footer_mode = "compact"
    
    print(f"Header file: {header_file}", file=sys.stderr)
    print(f"Output file: {output_file}", file=sys.stderr)
    print(f"Footer mode: {footer_mode}", file=sys.stderr)
    
    if not os.path.exists(header_file):
        print(f"Error: Header file '{header_file}' not found.", file=sys.stderr)
        sys.exit(1)
    
    print("Header file exists, proceeding with processing...", file=sys.stderr)
    
    try:
        # Read header file
        # Read header file
        with open(header_file, 'r', encoding='utf-8') as f:
            raw_headers = f.read()

        headers, export_metadata = HeaderParser.parse_export_file(raw_headers)
        
        logger.info(f"Processing header file: {header_file}")
        if export_metadata:
            logger.info(f"Export metadata: {export_metadata}")
        
        # Initialize components
        logger.info("Initializing components...")
        header_parser = HeaderParser()
        auth_parser = AuthenticationParser()
        geo_service = GeoLocationService()
        whois_service = WhoisService()
        security_assessor = SecurityAssessor()
        html_generator = HTMLReportGenerator()
        logger.info("Initializing GURI database...")
        try:
            guri_db, guri_db_type = connect_guri_database(
                base_path=str(_GEOFOOTER_ROOT), logger=logger
            )
            logger.info("GURI database initialized successfully (%s)", guri_db_type)
        except Exception as e:
            logger.error(f"Failed to initialize GURI database: {e}")
            guri_db = None
        logger.info("All components initialized successfully")
        
        # Extract sender information
        sender_email, sender_domain = header_parser.extract_sender_info(headers)
        sender_ip, sender_hostname = header_parser.extract_original_sender_ip(headers)
        
        logger.info(f"Sender: {sender_email} ({sender_domain})")
        logger.info(f"IP: {sender_ip} ({sender_hostname})")
        
        # Geolocate sender IP
        geo_result = geo_service.geolocate_ip(sender_ip) if sender_ip else GeoLocationResult(ip="Unknown")
        
        # Get WHOIS information
        whois_info = whois_service.get_whois_info(sender_domain) if sender_domain else {}
        
        # Extract authentication information
        auth_info = auth_parser.extract_authentication_info(headers)
        
        # Assess security risk
        security_assessment = security_assessor.assess_security_risk(
            [geo_result], auth_info, whois_info
        )
        
        # Generate HTML report
        html_content = html_generator.generate_report(
            sender_email, sender_domain, sender_ip, sender_hostname,
            geo_result, whois_info, auth_info, security_assessment, headers, guri_db,
            export_metadata,
            footer_mode,
        )
        
        # Ensure output directory exists
        _suite_output().mkdir(parents=True, exist_ok=True)
        
        # Use provided output file path or generate default one
        if output_file is None:
            # Generate timestamp-based filename matching VBA pattern
            from datetime import datetime
            import random
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S_")
            random_suffix = str(random.randint(0, 999)).zfill(3)
            footer_filename = f"footer_{timestamp}{random_suffix}.html"
            output_file = str(_suite_output(footer_filename))
            print(f"Generated output file path: {output_file}", file=sys.stderr)
        else:
            print(f"Using provided output file path: {output_file}", file=sys.stderr)
        
        # Write HTML/report pointer atomically (temp + replace) so Outlook's
        # reconcile sweep never sees a truncated empty output file.
        print(f"Writing HTML report to: {output_file}", file=sys.stderr)
        try:
            out_dir = os.path.dirname(output_file) or "."
            os.makedirs(out_dir, exist_ok=True)
            tmp_path = output_file + ".tmp"
            with open(tmp_path, 'w', encoding='utf-8', newline='\n') as f:
                f.write(html_content)
            os.replace(tmp_path, output_file)
            print(f"✓ HTML report written successfully", file=sys.stderr)
        except Exception as e:
            print(f"✗ Error writing HTML report: {e}", file=sys.stderr)
            try:
                if os.path.exists(output_file + ".tmp"):
                    os.remove(output_file + ".tmp")
            except OSError:
                pass
            raise
        
        logger.info(f"Report generated successfully: {output_file}")
        print(f"Report generated: {output_file}", file=sys.stderr)
        
        # Verify file was created
        if os.path.exists(output_file):
            file_size = os.path.getsize(output_file)
            print(f"✓ File verification: {output_file} exists ({file_size} bytes)", file=sys.stderr)
        else:
            print(f"✗ File verification failed: {output_file} does not exist", file=sys.stderr)
            sys.exit(1)
        
        # CRITICAL: Only output the footer file path to stdout (single line, no extra text)
        # This is what VBA expects to capture
        print(output_file)
        _notify_guri_scan_finished()
        try:
            logging.shutdown()
        except Exception:
            pass
        # Force-exit: leftover HTTP/DNS threads from whois/ipwhois can hang
        # pythonw for minutes after the report is written, which used to jam the
        # Outlook VBS waiter and leave AES PROC stuck.
        os._exit(0)

    except Exception as e:
        logger.error(f"Error processing headers: {e}")
        logger.error(f"Exception details: {type(e).__name__}: {str(e)}")
        import traceback
        logger.error(f"Traceback: {traceback.format_exc()}")
        print(f"Error: {e}", file=sys.stderr)
        if output_file:
            try:
                with open(output_file + ".fail", "w", encoding="utf-8") as fail_f:
                    fail_f.write(f"{type(e).__name__}: {e}\n")
            except OSError:
                pass
        try:
            logging.shutdown()
        except Exception:
            pass
        os._exit(1)

if __name__ == "__main__":
    main()