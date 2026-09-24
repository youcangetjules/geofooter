"""Routing plausibility checks for a Received chain.

Two independent signals:

* Geographic tromboning — the relay path leaves a region for somewhere far
  away and then comes back (e.g. London → Singapore → Dublin). Mail relays
  are application hops, not IP routes, so this is a *mail path* detour:
  forwarding services and misconfigured smart hosts do it innocently, but it
  is also how an interception relay shows up.
* BGP origin — for each public relay IP, what network is announcing its
  prefix right now (RIPEstat, no key needed): RPKI status of that
  announcement, multiple origin ASes (MOAS), mismatch against the ASN the
  host name / registry implies, and the origin's upstream peers.

An RPKI-invalid announcement is the strongest hijack indicator available
from public data; MOAS and ASN mismatches are weaker and often benign
(anycast, cloud-hosted senders using their own address space).
"""
from __future__ import annotations

import ipaddress
import logging
import math
from typing import Any, Dict, List, Optional, Tuple

try:
    import aes.addr_decode as addr_decode
except ImportError:  # pragma: no cover
    addr_decode = None  # type: ignore[assignment]

logger = logging.getLogger(__name__)

RIPESTAT_URL = "https://stat.ripe.net/data/{name}/data.json"
MAX_TOTAL_POINTS = 25
TROMBONE_AWAY_KM = 3000
TROMBONE_RETURN_KM = 1500
_CACHE: Dict[Tuple[str, str], Optional[Dict[str, Any]]] = {}


# --------------------------------------------------------------------------- #
# Geography
# --------------------------------------------------------------------------- #

def haversine_km(a: Tuple[float, float], b: Tuple[float, float]) -> float:
    lat1, lon1 = map(math.radians, a)
    lat2, lon2 = map(math.radians, b)
    h = (math.sin((lat2 - lat1) / 2) ** 2
         + math.cos(lat1) * math.cos(lat2) * math.sin((lon2 - lon1) / 2) ** 2)
    return 2 * 6371.0 * math.asin(min(1.0, math.sqrt(h)))


def _stops(hops: List[Dict[str, Any]]) -> List[Dict[str, Any]]:
    """Ordered located servers: each hop's sender, then its receiver."""
    out: List[Dict[str, Any]] = []

    def add(lat: Any, lon: Any, hop_no: int, name: Any, cc: Any) -> None:
        try:
            pt = (float(lat), float(lon))
        except (TypeError, ValueError):
            return
        if out and haversine_km(out[-1]["pt"], pt) < 50:
            return
        place = str(cc or "").strip()
        out.append({"pt": pt, "hop": hop_no, "name": str(name or "?"), "cc": place})

    for idx, hop in enumerate(hops, 1):
        if hop.get("hop_kind") != "internal" and hop.get("lat") is not None:
            add(hop.get("lat"), hop.get("lon"), idx, hop.get("fqdn"), hop.get("country"))
        by_loc = hop.get("by_location") or {}
        if by_loc.get("lat") is not None:
            add(by_loc["lat"], by_loc["lon"], idx, hop.get("by_host"), by_loc.get("country"))
    return out


def check_tromboning(hops: List[Dict[str, Any]]) -> List[Dict[str, Any]]:
    stops = _stops(hops)
    findings: List[Dict[str, Any]] = []
    seen = set()
    for i in range(len(stops) - 2):
        a = stops[i]
        for k in range(i + 2, len(stops)):
            c = stops[k]
            if haversine_km(a["pt"], c["pt"]) > TROMBONE_RETURN_KM:
                continue
            far = max(stops[i + 1:k], key=lambda s: haversine_km(a["pt"], s["pt"]))
            away = haversine_km(a["pt"], far["pt"])
            back = haversine_km(far["pt"], c["pt"])
            if away < TROMBONE_AWAY_KM or back < TROMBONE_AWAY_KM:
                continue
            key = (a["hop"], far["hop"], c["hop"])
            if key in seen:
                continue
            seen.add(key)
            detour = away + back - haversine_km(a["pt"], c["pt"])
            text = (
                f"path detours {away:,.0f} km out to {far['cc'] or far['name']} "
                f"and back ({a['cc'] or a['name']} → {far['cc'] or far['name']} → "
                f"{c['cc'] or c['name']}, ~{detour:,.0f} km extra)"
            )
            if 1 <= far["hop"] <= len(hops):
                hops[far["hop"] - 1].setdefault("route_flags", []).append(text)
            findings.append({
                "label": f"Mail path tromboning: {text}",
                "points": 5,
                "note": (
                    "Relays left the region and came back; forwarding services do this, "
                    "but so does an interception relay"
                ),
            })
            break
    return findings[:2]


# --------------------------------------------------------------------------- #
# BGP (RIPEstat)
# --------------------------------------------------------------------------- #

def _ripestat(name: str, resource: str, **extra: str) -> Optional[Dict[str, Any]]:
    key = (name, resource + "|" + "|".join(f"{k}={v}" for k, v in sorted(extra.items())))
    if key in _CACHE:
        return _CACHE[key]
    data: Optional[Dict[str, Any]] = None
    try:
        import requests
        from requests import exceptions as req_exc

        params = {"resource": resource, "sourceapp": "aliniant-aes", **extra}
        url = RIPESTAT_URL.format(name=name)
        try:
            resp = requests.get(url, params=params, timeout=5)
        except req_exc.SSLError:
            resp = requests.get(url, params=params, timeout=5, verify=False)
        if resp.status_code == 200:
            body = resp.json()
            if isinstance(body, dict) and isinstance(body.get("data"), dict):
                data = body["data"]
    except Exception as exc:
        logger.debug("RIPEstat %s(%s) failed: %s", name, resource, exc)
    _CACHE[key] = data
    return data


def _norm_asn(value: Any) -> str:
    s = str(value or "").strip().upper()
    if s.startswith("AS"):
        s = s[2:]
    return s if s.isdigit() else ""


def bgp_origin_info(ip: str, *, want_peers: bool = True) -> Optional[Dict[str, Any]]:
    """Current announcement for an IP: prefix, origin ASNs, RPKI, upstreams."""
    net = _ripestat("network-info", ip)
    if not net or not net.get("prefix"):
        return None
    prefix = str(net["prefix"])
    origins = [_norm_asn(a) for a in net.get("asns") or []]
    origins = [a for a in origins if a]
    info: Dict[str, Any] = {"prefix": prefix, "origins": origins, "holders": {},
                            "rpki": {}, "upstreams": []}

    overview = _ripestat("prefix-overview", prefix) or {}
    for entry in overview.get("asns") or []:
        asn = _norm_asn(entry.get("asn"))
        if asn:
            info["holders"][asn] = str(entry.get("holder") or "")
            if asn not in origins:
                origins.append(asn)

    for asn in origins[:3]:
        rv = _ripestat("rpki-validation", f"AS{asn}", prefix=prefix) or {}
        status = str(rv.get("status") or "").lower()
        if status:
            info["rpki"][asn] = status

    if want_peers and origins:
        nb = _ripestat("asn-neighbours", f"AS{origins[0]}") or {}
        ups = [n for n in nb.get("neighbours") or [] if n.get("type") == "left"]
        ups.sort(key=lambda n: -int(n.get("power") or 0))
        info["upstreams"] = [f"AS{n.get('asn')}" for n in ups[:4] if n.get("asn")]
    return info


def _public_ip(hop: Dict[str, Any]) -> Optional[str]:
    raw = str(hop.get("ip") or "").split()
    if not raw:
        return None
    try:
        addr = ipaddress.ip_address(raw[0])
    except ValueError:
        return None
    return str(addr) if addr.is_global else None


def check_bgp(hops: List[Dict[str, Any]], *, lite: bool) -> List[Dict[str, Any]]:
    findings: List[Dict[str, Any]] = []
    targets = [h for h in hops if h.get("hop_kind") == "public" and _public_ip(h)]
    if lite:
        targets = [h for h in targets if h.get("exit_hop")][:1]
    for hop in targets[:6]:
        ip = _public_ip(hop) or ""
        info = bgp_origin_info(ip, want_peers=not lite)
        if not info:
            continue
        hop["bgp"] = info
        origins = info["origins"]
        lines: List[str] = []
        names = [f"AS{a}" + (f" {info['holders'][a]}" if info["holders"].get(a) else "")
                 for a in origins]
        lines.append(f"{info['prefix']} announced by {', '.join(names) or 'nobody'}")

        invalid = [a for a, s in info["rpki"].items() if s.startswith("invalid")]
        valid = [a for a, s in info["rpki"].items() if s == "valid"]
        if invalid:
            lines.append(f"RPKI INVALID for AS{', AS'.join(invalid)}")
            findings.append({
                "label": f"BGP: {info['prefix']} ({ip}) announced by AS{invalid[0]} "
                         "which RPKI does not authorise",
                "points": 20,
                "note": "Possible BGP hijack or route leak — the announcing network is not "
                        "the one the address holder signed for",
            })
        elif valid:
            lines.append("RPKI valid")
        elif info["rpki"]:
            lines.append("RPKI: no ROA (unsigned)")

        if len(origins) > 1:
            lines.append(f"MOAS: {len(origins)} origin ASes")
            findings.append({
                "label": f"BGP: {info['prefix']} ({ip}) has {len(origins)} origin ASes "
                         f"({', '.join('AS' + a for a in origins)})",
                "points": 5,
                "note": "Multiple networks announce the same prefix — anycast or a hijack",
            })

        expected: List[Tuple[str, str]] = []
        hop_asn = _norm_asn(hop.get("asn"))
        if hop_asn:
            expected.append((hop_asn, "registry/geo"))
        if addr_decode is not None:
            for name in (hop.get("fqdn"), hop.get("helo")):
                h_asn, h_org = addr_decode.provider_asn_for_host(name)
                if _norm_asn(h_asn):
                    expected.append((_norm_asn(h_asn), f"host name ({h_org})"))
                    break
        for asn, source in expected:
            if origins and asn not in origins:
                lines.append(f"{source} says AS{asn}, not announcing")
                findings.append({
                    "label": f"BGP: {ip} expected AS{asn} ({source}) but announced by "
                             f"AS{', AS'.join(origins)}",
                    "points": 0,
                    "note": "Informational — common for cloud-hosted senders and stale "
                            "registry data; weighs more alongside RPKI-invalid or MOAS",
                })
                break

        hop["bgp_lines"] = lines
        if info["upstreams"]:
            hop["bgp_peers"] = [f"{u} (upstream)" for u in info["upstreams"]]
    return findings


def check_routing(hops: List[Dict[str, Any]], *, lite: bool) -> List[Dict[str, Any]]:
    """All routing findings, points capped at MAX_TOTAL_POINTS."""
    findings: List[Dict[str, Any]] = []
    for fn in (lambda: check_tromboning(hops), lambda: check_bgp(hops, lite=lite)):
        try:
            findings.extend(fn())
        except Exception as exc:
            logger.warning("Routing check failed: %s", exc)
    total = 0
    capped: List[Dict[str, Any]] = []
    for f in findings:
        pts = max(0, min(int(f["points"]), MAX_TOTAL_POINTS - total))
        total += pts
        capped.append({**f, "points": pts})
    return capped
