"""Read tracking for outgoing Aliniant mail via AES's own pixel.

VBA (``MSCANPixel``) adds ``<img src="{base}/p/{id}.gif">`` to each HTML mail
sent from the Aliniant account and writes ``pixels/{id}.json`` describing the
send. The pixel endpoint (``cloud/pixel``) records every fetch. ``poll()``
pulls new fetches, drops the user's own (Sent Items views come from the same
public IP that polls), labels proxy and pre-fetch traffic, and keeps a local
SQLite copy that the GURI Tracking tab reads.

Files under ``%LOCALAPPDATA%\\GeoFooter``:
  aes_pixel.ini          enabled=0|1, base_url=https://track.i-txone.com
  pixels\\{id}.json       one per tracked send (written by VBA)
  pixel_opens.sqlite     sends, opens, poll state
"""

from __future__ import annotations

import ipaddress
import json
import os
import re
import sqlite3
import time
from dataclasses import dataclass
from pathlib import Path
from typing import Any, Dict, Iterable, List, Optional

DEFAULT_BASE_URL = "https://track.i-txone.com"
SECRET_NAME = "aes_pixel_key"
PIXEL_ID = re.compile(r"^[0-9a-f]{32}$")
# ReadNotify runs first at send time and rewrites recipients to
# user@domain.com[.flag...].read-notify.com.
_READNOTIFY_SUFFIX = re.compile(
    r"(?:\.(?:silent|noprint|translate|certified|selfdestruct|ensured))*\.read-notify\.com\b",
    re.IGNORECASE,
)

# Google proxies Gmail images through these ranges; the open is real but the IP
# and location are Google's.
_GOOGLE_PROXY_NETS = [ipaddress.ip_network(n) for n in ("66.249.80.0/20", "66.102.0.0/20", "74.125.0.0/16", "64.233.160.0/19")]
_APPLE_NETS = [ipaddress.ip_network("17.0.0.0/8")]


def _base_dir() -> Path:
    local = os.environ.get("LOCALAPPDATA", "")
    base = Path(local) / "GeoFooter" if local else Path.home() / ".geofooter"
    base.mkdir(parents=True, exist_ok=True)
    return base


def config_path() -> Path:
    return _base_dir() / "aes_pixel.ini"


def ledger_dir() -> Path:
    return _base_dir() / "pixels"


def db_path() -> Path:
    return _base_dir() / "pixel_opens.sqlite"


def load_config() -> Dict[str, Any]:
    cfg: Dict[str, Any] = {"enabled": False, "base_url": DEFAULT_BASE_URL}
    try:
        for line in config_path().read_text(encoding="utf-8-sig").splitlines():
            key, sep, value = line.partition("=")
            if not sep:
                continue
            key, value = key.strip().lower(), value.strip()
            if key == "enabled":
                cfg["enabled"] = value.lower() in {"1", "true", "yes", "on"}
            elif key == "base_url" and value:
                cfg["base_url"] = value.rstrip("/")
    except OSError:
        pass
    return cfg


def api_key() -> str:
    env = (os.environ.get("AES_PIXEL_KEY") or "").strip()
    if env:
        return env
    try:
        from aes.secret_store import get_secret
    except ImportError:
        return ""
    return (get_secret(SECRET_NAME) or "").strip()


# --------------------------------------------------------------------------- #
# Classification of a fetch
# --------------------------------------------------------------------------- #

@dataclass
class OpenLabel:
    client: str
    proxied: bool  # IP/location belong to a proxy, not the reader
    prefetch: bool  # fetched automatically, may not be a real read


def _in_nets(ip: str, nets: Iterable[ipaddress._BaseNetwork]) -> bool:
    try:
        addr = ipaddress.ip_address(ip)
    except ValueError:
        return False
    return any(addr in n for n in nets)


def classify(ip: str, ua: str) -> OpenLabel:
    u = (ua or "").strip()
    low = u.lower()
    if "googleimageproxy" in low or (_in_nets(ip, _GOOGLE_PROXY_NETS) and "google" in low):
        return OpenLabel("Gmail (Google image proxy)", True, False)
    if "yahoomailproxy" in low:
        return OpenLabel("Yahoo Mail (image proxy)", True, False)
    # Apple Mail Privacy Protection fetches on delivery with a bare UA.
    if u == "Mozilla/5.0" or _in_nets(ip, _APPLE_NETS):
        return OpenLabel("Apple Mail (Privacy Protection pre-fetch)", True, True)
    if "microsoft office" in low or "ms-office" in low or "outlook" in low:
        return OpenLabel("Outlook", False, False)
    if "thunderbird" in low:
        return OpenLabel("Thunderbird", False, False)
    if "iphone" in low or "ipad" in low:
        return OpenLabel("iOS Mail", False, False)
    if "android" in low:
        return OpenLabel("Android mail app", False, False)
    if "windows" in low or "macintosh" in low:
        return OpenLabel("Webmail / desktop browser", False, False)
    if not u:
        return OpenLabel("Unknown client", False, False)
    return OpenLabel(u[:60], False, False)


# --------------------------------------------------------------------------- #
# Local store
# --------------------------------------------------------------------------- #

def _connect(path: Optional[Path] = None) -> sqlite3.Connection:
    conn = sqlite3.connect(str(path or db_path()), timeout=5)
    conn.execute(
        "CREATE TABLE IF NOT EXISTS sends ("
        " pixel TEXT PRIMARY KEY, guri TEXT, sender TEXT, recipients TEXT,"
        " subject TEXT, sent_local TEXT)"
    )
    conn.execute(
        "CREATE TABLE IF NOT EXISTS opens ("
        " id INTEGER PRIMARY KEY, pixel TEXT, ts INTEGER, ip TEXT, ua TEXT,"
        " client TEXT, proxied INTEGER, prefetch INTEGER, own INTEGER)"
    )
    conn.execute("CREATE TABLE IF NOT EXISTS state (k TEXT PRIMARY KEY, v TEXT)")
    return conn


def _state_get(conn: sqlite3.Connection, key: str, default: str = "") -> str:
    row = conn.execute("SELECT v FROM state WHERE k = ?", (key,)).fetchone()
    return row[0] if row else default


def _state_set(conn: sqlite3.Connection, key: str, value: str) -> None:
    conn.execute("INSERT OR REPLACE INTO state (k, v) VALUES (?, ?)", (key, value))


def import_ledger(conn: sqlite3.Connection, folder: Optional[Path] = None) -> int:
    """Copy new pixels/{id}.json files written by VBA into the sends table."""
    added = 0
    folder = folder or ledger_dir()
    if not folder.is_dir():
        return 0
    known = {r[0] for r in conn.execute("SELECT pixel FROM sends")}
    for path in folder.glob("*.json"):
        pixel = path.stem.lower()
        if pixel in known or not PIXEL_ID.match(pixel):
            continue
        try:
            data = json.loads(path.read_text(encoding="utf-8-sig"))
        except (OSError, ValueError):
            continue
        recipients = _READNOTIFY_SUFFIX.sub("", "; ".join(
            p for p in (str(data.get(k) or "").strip() for k in ("to", "cc", "bcc")) if p
        ))
        conn.execute(
            "INSERT OR IGNORE INTO sends (pixel, sender, recipients, subject, sent_local)"
            " VALUES (?, ?, ?, ?, ?)",
            (pixel, str(data.get("sender") or ""), recipients,
             str(data.get("subject") or ""), str(data.get("sent_local") or "")),
        )
        added += 1
    return added


def store_opens(conn: sqlite3.Connection, payload: Dict[str, Any]) -> int:
    """Save a /opens response. Fetches from the polling IP are the user's own."""
    own_ips = set(filter(None, _state_get(conn, "own_ips").split(",")))
    client_ip = str(payload.get("client_ip") or "").strip()
    if client_ip and client_ip not in own_ips:
        own_ips.add(client_ip)
        _state_set(conn, "own_ips", ",".join(sorted(own_ips)))
    last = int(_state_get(conn, "last_id", "0") or 0)
    stored = 0
    for row in payload.get("opens") or []:
        try:
            rid = int(row.get("id"))
        except (TypeError, ValueError):
            continue
        ip, ua = str(row.get("ip") or ""), str(row.get("ua") or "")
        label = classify(ip, ua)
        conn.execute(
            "INSERT OR IGNORE INTO opens (id, pixel, ts, ip, ua, client, proxied, prefetch, own)"
            " VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)",
            (rid, str(row.get("pixel") or ""), int(row.get("ts") or 0), ip, ua,
             label.client, int(label.proxied), int(label.prefetch), int(ip in own_ips)),
        )
        last = max(last, rid)
        stored += 1
    if client_ip:
        conn.execute("UPDATE opens SET own = 1 WHERE ip = ?", (client_ip,))
    _state_set(conn, "last_id", str(last))
    return stored


def poll(*, timeout: float = 10.0) -> Dict[str, Any]:
    """Fetch new opens from the endpoint. Safe to call from a worker thread."""
    cfg = load_config()
    result: Dict[str, Any] = {"ledger": 0, "opens": 0, "error": ""}
    conn = _connect()
    try:
        result["ledger"] = import_ledger(conn)
        key = api_key()
        if not key:
            result["error"] = "no pixel API key configured"
            return result
        from aes.threat_intel import _http

        since = _state_get(conn, "last_id", "0")
        while True:
            resp = _http("GET", f"{cfg['base_url']}/opens", params={"since": since},
                         headers={"X-AES-Key": key}, timeout=timeout)
            if resp.status_code != 200:
                result["error"] = f"endpoint returned HTTP {resp.status_code}"
                break
            payload = resp.json()
            n = store_opens(conn, payload)
            result["opens"] += n
            if n < 1000:
                break
            since = _state_get(conn, "last_id", since)
        _state_set(conn, "last_poll", str(int(time.time())))
    except Exception as exc:  # network errors surface in the GURI status line
        result["error"] = f"{type(exc).__name__}: {exc}"
    finally:
        conn.commit()
        conn.close()
    return result


def sends_without_guri() -> List[Dict[str, str]]:
    conn = _connect()
    try:
        rows = conn.execute(
            "SELECT pixel, sender, recipients, subject, sent_local FROM sends"
            " WHERE guri IS NULL OR guri = ''"
        ).fetchall()
    finally:
        conn.close()
    keys = ("pixel", "sender", "recipients", "subject", "sent_local")
    return [dict(zip(keys, r)) for r in rows]


def set_guri(pixel: str, guri: str) -> None:
    conn = _connect()
    try:
        conn.execute("UPDATE sends SET guri = ? WHERE pixel = ?", (guri, pixel))
        conn.commit()
    finally:
        conn.close()


def summaries() -> List[Dict[str, Any]]:
    """One row per tracked send, newest first, excluding the user's own views."""
    conn = _connect()
    try:
        rows = conn.execute(
            "SELECT s.pixel, s.guri, s.sender, s.recipients, s.subject, s.sent_local,"
            " COUNT(o.id),"
            " SUM(CASE WHEN o.prefetch = 0 THEN 1 ELSE 0 END),"
            " MIN(o.ts), MAX(o.ts)"
            " FROM sends s LEFT JOIN opens o ON o.pixel = s.pixel AND o.own = 0"
            " GROUP BY s.pixel ORDER BY s.sent_local DESC"
        ).fetchall()
        last_client = {
            r[0]: r[1]
            for r in conn.execute(
                "SELECT pixel, client FROM opens WHERE own = 0 ORDER BY ts"
            )
        }
    finally:
        conn.close()
    out = []
    for pixel, guri, sender, recips, subject, sent, n, real, first, last in rows:
        out.append({
            "pixel": pixel, "guri": guri or "", "sender": sender, "recipients": recips,
            "subject": subject, "sent_local": sent, "opens": int(n or 0),
            "likely_reads": int(real or 0), "first_open": int(first or 0),
            "last_open": int(last or 0), "last_client": last_client.get(pixel, ""),
        })
    return out


def opens_for(pixel: str) -> List[Dict[str, Any]]:
    conn = _connect()
    try:
        rows = conn.execute(
            "SELECT ts, ip, client, proxied, prefetch, own FROM opens WHERE pixel = ? ORDER BY ts",
            (pixel,),
        ).fetchall()
    finally:
        conn.close()
    return [
        {"ts": r[0], "ip": r[1], "client": r[2], "proxied": bool(r[3]),
         "prefetch": bool(r[4]), "own": bool(r[5])}
        for r in rows
    ]
