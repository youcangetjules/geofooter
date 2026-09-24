"""Timestamp consistency checks for Received chains and the Date header.

Received headers are stored oldest-first after extraction (hop 1 = origin /
starting point, last hop = delivery). Each timestamp is written by the
*receiving* ("by") server using its own clock and UTC offset. Two classes of
evidence:

* Ordering  — a later hop stamped earlier than the hop above it means clock
  skew at best and an inserted/forged Received line at worst.
* Location  — a non-UTC offset that is not used anywhere in the country where
  the stamping server sits (or the Date header offset vs the origin country)
  suggests a relay or sender pretending to be somewhere it is not.

UTC (+0000) is never flagged — most mail infrastructure stamps UTC regardless
of location — and Google stamps US Pacific time on every server worldwide.
"""
from __future__ import annotations

import email.utils
from datetime import datetime, timedelta, timezone
from typing import Any, Dict, Iterable, List, Optional, Set

try:
    import pytz
except ImportError:  # pragma: no cover
    pytz = None  # type: ignore[assignment]

try:
    import pycountry
except ImportError:  # pragma: no cover
    pycountry = None  # type: ignore[assignment]

try:
    import aes_addr_decode as addr_decode
except ImportError:  # pragma: no cover
    addr_decode = None  # type: ignore[assignment]

ORDER_TOLERANCE_S = 300
ORDER_STRONG_S = 3600
MAX_TOTAL_POINTS = 25

_GOOGLE_PACIFIC = {-420, -480}


def iso2_for(country: Optional[str]) -> Optional[str]:
    """Country name or ISO code -> ISO alpha-2 (upper), else None."""
    if not country:
        return None
    c = str(country).strip()
    if not c or c.lower() in {"unknown", "—", "-", "eu", "n/a"}:
        return None
    if len(c) == 2 and c.isalpha():
        return c.upper()
    if pycountry is None:
        return None
    try:
        hit = pycountry.countries.get(alpha_3=c.upper()) if len(c) == 3 else None
        if not hit:
            hit = pycountry.countries.get(name=c) or pycountry.countries.get(common_name=c)
        if not hit:
            hit = pycountry.countries.lookup(c)
        return hit.alpha_2 if hit else None
    except (LookupError, KeyError, AttributeError):
        return None


def country_offsets(iso2: str, when_utc: datetime) -> Set[int]:
    """All UTC offsets (minutes) in use in a country at a given instant."""
    if pytz is None or not iso2:
        return set()
    try:
        zones = pytz.country_timezones.get(iso2.upper()) or []
    except Exception:
        return set()
    naive = when_utc.astimezone(timezone.utc).replace(tzinfo=None)
    out: Set[int] = set()
    for name in zones:
        try:
            off = pytz.timezone(name).utcoffset(naive)
            if off is not None:
                out.add(int(off.total_seconds() // 60))
        except Exception:
            continue
    return out


def fmt_offset(minutes: int) -> str:
    sign = "+" if minutes >= 0 else "-"
    m = abs(minutes)
    return f"UTC{sign}{m // 60:02d}:{m % 60:02d}"


def _fmt_span(seconds: float) -> str:
    s = abs(int(seconds))
    if s >= 86400:
        return f"{s / 86400:.1f} days"
    if s >= 3600:
        return f"{s / 3600:.1f}h"
    return f"{max(1, s // 60)}m"


def _parse_iso(value: Any) -> Optional[datetime]:
    if not value:
        return None
    try:
        return datetime.fromisoformat(str(value))
    except ValueError:
        return None


def _is_google_host(hop: Dict[str, Any]) -> bool:
    names = " ".join(str(hop.get(k) or "") for k in ("by_host", "fqdn", "helo")).lower()
    if "google.com" in names or "gmail.com" in names:
        return True
    dec = hop.get("ip_decode") or {}
    return "google" in str(dec.get("org_hint") or "").lower()


def _first_label(name: Any) -> str:
    return str(name or "").strip().lower().split(".", 1)[0]


def _stamping_server_country(hops: List[Dict[str, Any]], i: int) -> Optional[tuple]:
    """(iso2, source, strong) for the server that wrote hop i's timestamp.

    ``strong`` is True for decoded infrastructure host names; plain IP
    geolocation is weaker (GeoIP databases often misplace datacenter ranges).
    """
    hop = hops[i]
    by_host = str(hop.get("by_host") or "")
    if addr_decode is not None and by_host and not addr_decode.ip_literal(by_host):
        loc = hop.get("by_location") or addr_decode.decode_host_location(by_host)
        if loc and loc.get("confidence") in {"high", "medium"}:
            iso = iso2_for(loc.get("country"))
            if iso:
                return iso, f"host name {by_host}", loc.get("confidence") == "high"
    if i + 1 < len(hops) and by_host:
        # Oldest-first: the next hop's "from" is this hop's "by".
        nxt = hops[i + 1]
        lbl = _first_label(by_host)
        if lbl and lbl in {_first_label(nxt.get("fqdn")), _first_label(nxt.get("helo"))}:
            if nxt.get("hop_kind") == "public":
                iso = iso2_for(nxt.get("country"))
                if iso:
                    strong = str(nxt.get("location_source") or "").startswith("from host name")
                    return iso, f"IP geolocation of {by_host}", strong
    return None


def _add_flag(hop: Dict[str, Any], text: str) -> None:
    hop.setdefault("time_flags", []).append(text)


def check_time_consistency(
    hops: List[Dict[str, Any]],
    *,
    date_header: Optional[str] = None,
    origin_country: Optional[str] = None,
    now_utc: Optional[datetime] = None,
) -> List[Dict[str, Any]]:
    """Annotate hops with ``time_flags`` and return scored findings.

    Each finding: ``{"label", "points", "note"}``. Points per class are applied
    once (max), and the total is capped at ``MAX_TOTAL_POINTS``.
    """
    now_utc = now_utc or datetime.now(timezone.utc)
    findings: List[Dict[str, Any]] = []
    times = [_parse_iso(h.get("timestamp_utc")) for h in hops]

    # 1) Ordering: oldest-first list, so times should never decrease downwards.
    worst_skew = 0.0
    worst_desc = ""
    for i in range(1, len(hops)):
        earlier, later = times[i - 1], times[i]
        if earlier is None or later is None:
            continue
        skew = (earlier - later).total_seconds()
        if skew > ORDER_TOLERANCE_S:
            _add_flag(
                hops[i],
                f"stamped {_fmt_span(skew)} before the previous server received it "
                "(clock skew or forged hop)",
            )
            if skew > worst_skew:
                worst_skew = skew
                worst_desc = f"hop {i + 1} is {_fmt_span(skew)} earlier than hop {i}"
    if worst_skew:
        findings.append({
            "label": f"Received timestamps out of order ({worst_desc})",
            "points": 15 if worst_skew >= ORDER_STRONG_S else 5,
            "note": "Later hop stamped before an earlier one; >1h suggests an inserted/forged Received line",
        })

    if times and times[-1] is not None and (times[-1] - now_utc).total_seconds() > 600:
        _add_flag(hops[-1], f"timestamp is {_fmt_span((times[-1] - now_utc).total_seconds())} in the future")
        findings.append({
            "label": "Final Received timestamp is in the future",
            "points": 5,
            "note": "Receiving server clock wrong or header forged",
        })

    # 2) Offset vs stamping-server location.
    offset_hits: List[str] = []
    offset_pts = 0
    for i, hop in enumerate(hops):
        off = hop.get("utc_offset_min")
        t = times[i]
        if off is None or t is None or off == 0:
            continue
        if off in _GOOGLE_PACIFIC and _is_google_host(hop):
            continue
        where = _stamping_server_country(hops, i)
        if not where:
            continue
        iso, source, strong = where
        valid = country_offsets(iso, t)
        if not valid or off in valid:
            continue
        gap = min(abs(off - v) for v in valid)
        if gap < 60:
            continue
        valid_txt = ", ".join(fmt_offset(v) for v in sorted(valid)[:3])
        _add_flag(
            hop,
            f"clock {fmt_offset(off)} not used in {iso} ({valid_txt}) — location from {source}",
        )
        offset_hits.append(f"hop {i + 1}: {fmt_offset(off)} vs {iso}")
        offset_pts = max(offset_pts, 10 if (strong or gap >= 120) else 3)
    if offset_hits:
        findings.append({
            "label": "Hop clock offset does not match server location (" + "; ".join(offset_hits[:3]) + ")",
            "points": offset_pts,
            "note": (
                "Server claims a timezone not used where it is located; "
                "+10 for a decoded host location or 2h+ gap, +3 for a 1h gap on GeoIP alone"
            ),
        })

    # 3) Date header vs transport and origin.
    date_dt = None
    date_has_offset = False
    if date_header:
        try:
            date_dt = email.utils.parsedate_to_datetime(date_header.strip())
            date_has_offset = date_dt.tzinfo is not None and "-0000" not in date_header
            if date_dt.tzinfo is None:
                date_dt = date_dt.replace(tzinfo=timezone.utc)
        except (TypeError, ValueError):
            date_dt = None
    first_seen = next((t for t in times if t is not None), None)
    if date_dt is not None and first_seen is not None:
        lead = (date_dt - first_seen).total_seconds()
        if lead > 900:
            findings.append({
                "label": f"Date header is {_fmt_span(lead)} after the first server received the mail",
                "points": 10,
                "note": "A message cannot be written after it was sent; forged or rewritten Date",
            })
        elif lead < -72 * 3600:
            findings.append({
                "label": f"Date header is {_fmt_span(lead)} before the first Received hop",
                "points": 5,
                "note": "Old message replayed or held back before sending",
            })
    if date_dt is not None and date_has_offset:
        off = int(date_dt.utcoffset().total_seconds() // 60)
        iso = iso2_for(origin_country)
        if off != 0 and iso:
            valid = country_offsets(iso, date_dt)
            if valid and off not in valid and min(abs(off - v) for v in valid) >= 120:
                valid_txt = ", ".join(fmt_offset(v) for v in sorted(valid)[:3])
                findings.append({
                    "label": f"Sender clock {fmt_offset(off)} does not match origin country {iso} ({valid_txt})",
                    "points": 10,
                    "note": "Date header timezone disagrees with where the sending IP is",
                })

    total = 0
    capped: List[Dict[str, Any]] = []
    for f in findings:
        pts = max(0, min(int(f["points"]), MAX_TOTAL_POINTS - total))
        total += pts
        capped.append({**f, "points": pts})
    return capped


def any_flags(hops: Iterable[Dict[str, Any]]) -> bool:
    return any(h.get("time_flags") for h in hops)
