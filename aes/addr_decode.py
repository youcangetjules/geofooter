"""Offline decoding of IP addresses and infrastructure hostnames seen in Received hops.

Everything here is pure computation — no network calls — so it is safe for
compact scans. Hostname locations are inferred from provider naming
conventions (not published by the providers) and carry a confidence label.
"""

from __future__ import annotations

import ipaddress
import re
from typing import Any, Dict, List, Optional, Tuple

# ---------------------------------------------------------------------------
# IP addresses
# ---------------------------------------------------------------------------

# Provider ranges published in their SPF records / network docs.
# (cidr, organisation label, ASN — None when the range is shared / unknown)
_PROVIDER_NETS: List[Tuple[str, str, Optional[str]]] = [
    ("2001:4860::/32", "Google", "AS15169"),
    ("2404:6800::/32", "Google", "AS15169"),
    ("2607:f8b0::/32", "Google", "AS15169"),
    ("2800:3f0::/32", "Google", "AS15169"),
    ("2a00:1450::/32", "Google", "AS15169"),
    ("2c0f:fb50::/32", "Google", "AS15169"),
    ("35.190.247.0/24", "Google", "AS15169"),
    ("64.233.160.0/19", "Google", "AS15169"),
    ("66.102.0.0/20", "Google", "AS15169"),
    ("66.249.80.0/20", "Google", "AS15169"),
    ("72.14.192.0/18", "Google", "AS15169"),
    ("74.125.0.0/16", "Google", "AS15169"),
    ("108.177.8.0/21", "Google", "AS15169"),
    ("172.217.0.0/19", "Google", "AS15169"),
    ("173.194.0.0/16", "Google", "AS15169"),
    ("209.85.128.0/17", "Google", "AS15169"),
    ("216.58.192.0/19", "Google", "AS15169"),
    ("216.239.32.0/19", "Google", "AS15169"),
    ("2a01:111:f400::/48", "Microsoft Exchange Online Protection", "AS8075"),
    ("2a01:111:f403::/48", "Microsoft Exchange Online Protection", "AS8075"),
    ("2a01:111::/32", "Microsoft", "AS8075"),
    ("2603:10a6::/32", "Microsoft Exchange Online (Europe forests)", "AS8075"),
    ("2603:1000::/24", "Microsoft", "AS8075"),
    ("40.92.0.0/15", "Microsoft Exchange Online Protection", "AS8075"),
    ("40.107.0.0/16", "Microsoft Exchange Online Protection", "AS8075"),
    ("52.100.0.0/14", "Microsoft Exchange Online Protection", "AS8075"),
    ("104.47.0.0/17", "Microsoft Exchange Online Protection", "AS8075"),
    ("54.240.0.0/18", "Amazon SES", "AS16509"),
    ("23.249.208.0/20", "Amazon SES", "AS16509"),
    ("163.47.180.0/23", "Sailthru", "AS18877"),
]
_PROVIDER_NETS_PARSED = [
    (ipaddress.ip_network(n), name, asn) for n, name, asn in _PROVIDER_NETS
]

# Host-name suffixes that identify a provider ASN without an IP lookup.
_PROVIDER_HOST_ASN: List[Tuple[str, Optional[str], str]] = [
    (".prod.outlook.com", "AS8075", "Microsoft"),
    (".outlook.com", "AS8075", "Microsoft"),
    (".mail.protection.outlook.com", "AS8075", "Microsoft Exchange Online Protection"),
    (".protection.outlook.com", "AS8075", "Microsoft Exchange Online Protection"),
    (".office365.com", "AS8075", "Microsoft"),
    (".google.com", "AS15169", "Google"),
    (".gmail.com", "AS15169", "Google"),
    (".googlemail.com", "AS15169", "Google"),
    (".amazonses.com", "AS16509", "Amazon SES"),
    (".amazonaws.com", "AS16509", "Amazon"),
    (".sailthru.com", "AS18877", "Sailthru"),
    (".mimecast.com", None, "Mimecast"),
]

_V4_SPECIAL: List[Tuple[str, str, str]] = [
    # (network, label, scope) — scope: private | special | suspicious
    ("0.0.0.0/8", "'This network' (0/8) — not a real source", "suspicious"),
    ("10.0.0.0/8", "Private network (RFC 1918, 10/8)", "private"),
    ("100.64.0.0/10", "Carrier-grade NAT (RFC 6598) — sender is behind ISP NAT", "private"),
    ("127.0.0.0/8", "Loopback — same machine", "private"),
    ("169.254.0.0/16", "Link-local (APIPA) — no DHCP address", "private"),
    ("172.16.0.0/12", "Private network (RFC 1918, 172.16/12)", "private"),
    ("192.0.0.0/24", "IETF protocol assignments (192.0.0/24)", "special"),
    ("192.0.2.0/24", "Documentation range TEST-NET-1 — should never appear in real mail", "suspicious"),
    ("192.88.99.0/24", "6to4 relay anycast (deprecated)", "special"),
    ("192.168.0.0/16", "Private network (RFC 1918, 192.168/16)", "private"),
    ("198.18.0.0/15", "Benchmarking range (RFC 2544)", "special"),
    ("198.51.100.0/24", "Documentation range TEST-NET-2 — should never appear in real mail", "suspicious"),
    ("203.0.113.0/24", "Documentation range TEST-NET-3 — should never appear in real mail", "suspicious"),
    ("224.0.0.0/4", "Multicast", "special"),
    ("240.0.0.0/4", "Reserved (240/4)", "suspicious"),
]
_V4_SPECIAL_PARSED = [(ipaddress.ip_network(n), lbl, sc) for n, lbl, sc in _V4_SPECIAL]


def _provider_for(addr: ipaddress._BaseAddress) -> Tuple[Optional[str], Optional[str]]:
    """Return (organisation label, ASN) for a known provider address."""
    for net, name, asn in _PROVIDER_NETS_PARSED:
        if addr.version == net.version and addr in net:
            return name, asn
    return None, None


def provider_asn_for_host(hostname: Optional[str]) -> Tuple[Optional[str], Optional[str]]:
    """Return (ASN, organisation) inferred from a mail-infra hostname."""
    host = (hostname or "").strip().rstrip(".").lower()
    if not host:
        return None, None
    for suffix, asn, org in _PROVIDER_HOST_ASN:
        if host == suffix.lstrip(".") or host.endswith(suffix):
            return asn, org
    return None, None


def _v4_class(v4: ipaddress.IPv4Address) -> Tuple[str, str]:
    """(label, scope) for an IPv4 address; scope 'public' when nothing special."""
    for net, lbl, sc in _V4_SPECIAL_PARSED:
        if v4 in net:
            return lbl, sc
    if v4 == ipaddress.IPv4Address("255.255.255.255"):
        return "Limited broadcast", "suspicious"
    return "Public IPv4", "public"


def _interface_id_notes(v6: ipaddress.IPv6Address) -> List[str]:
    """Describe the low 64 bits (interface identifier)."""
    iid = int(v6) & ((1 << 64) - 1)
    b = iid.to_bytes(8, "big")
    if b[3] == 0xFF and b[4] == 0xFE:
        mac = bytes([b[0] ^ 0x02, b[1], b[2], b[5], b[6], b[7]])
        mac_s = ":".join(f"{x:02x}" for x in mac)
        kind = "locally administered" if mac[0] & 0x02 else "vendor-assigned (OUI " + mac_s[:8] + ")"
        return [f"Interface ID is EUI-64 — reveals hardware MAC {mac_s} ({kind})"]
    if b[0] in (0x00, 0x02) and b[1:4] == b"\x00\x5e\xfe":
        v4 = ipaddress.IPv4Address(b[4:8])
        return [f"ISATAP interface ID — embeds IPv4 {v4} ({_v4_class(v4)[0]})"]
    if iid == 0:
        return ["Interface ID is zero (subnet-router anycast)"]
    if iid <= 0xFFFF:
        return [f"Interface ID ::{iid:x} — manually assigned (typical of servers/routers)"]
    return ["Interface ID looks random (privacy/opaque) — not hardware-derived"]


def decode_ip(ip: str, *, provider_context: Optional[set] = None) -> Optional[Dict[str, Any]]:
    """Decode an IPv4/IPv6 literal.

    Returns dict with:
      kind          short type label (e.g. '6to4', 'Teredo', 'Private IPv4')
      scope         public | private | special | suspicious
      lines         human-readable decode lines
      embedded_ipv4 IPv4 carried inside an IPv6 address, if any
      geo_ip        public IPv4 usable for geolocation instead of the literal
      org_hint      organisation label when the address alone identifies it
      asn_hint      ASN (e.g. AS8075) when the address alone identifies it
    """
    try:
        addr = ipaddress.ip_address(str(ip).strip().strip("[]"))
    except ValueError:
        return None
    ctx = {c.lower() for c in (provider_context or set())}
    out: Dict[str, Any] = {
        "kind": "", "scope": "public", "lines": [], "embedded_ipv4": None,
        "geo_ip": None, "org_hint": None, "asn_hint": None,
    }
    provider, provider_asn = _provider_for(addr)

    if addr.version == 4:
        lbl, sc = _v4_class(addr)
        out.update(kind=lbl.split(" (")[0].split(" —")[0], scope=sc)
        if sc != "public":
            out["lines"].append(lbl)
        if provider:
            out["lines"].append(f"Inside {provider}'s published range")
            out["org_hint"] = provider
            out["asn_hint"] = provider_asn
        return out

    v6: ipaddress.IPv6Address = addr  # type: ignore[assignment]
    lines: List[str] = out["lines"]

    if v6 == ipaddress.IPv6Address("::"):
        out.update(kind="Unspecified", scope="suspicious")
        lines.append("Unspecified address (::) — not a real source")
        return out
    if v6.is_loopback:
        out.update(kind="Loopback", scope="private")
        lines.append("IPv6 loopback (::1) — same machine")
        return out
    if v6.ipv4_mapped:
        v4 = v6.ipv4_mapped
        lbl, sc = _v4_class(v4)
        out.update(kind="IPv4-mapped", scope=sc, embedded_ipv4=str(v4))
        lines.append(f"IPv4-mapped IPv6 — really IPv4 {v4} ({lbl})")
        if sc == "public":
            out["geo_ip"] = str(v4)
        return out
    if v6 in ipaddress.ip_network("64:ff9b::/96") or v6 in ipaddress.ip_network("64:ff9b:1::/48"):
        v4 = ipaddress.IPv4Address(int(v6) & 0xFFFFFFFF)
        lbl, sc = _v4_class(v4)
        out.update(kind="NAT64", scope=sc, embedded_ipv4=str(v4))
        lines.append(f"NAT64 — IPv6-only network reaching IPv4 {v4} ({lbl})")
        if sc == "public":
            out["geo_ip"] = str(v4)
        return out
    if v6.teredo:
        server, client = v6.teredo
        out.update(kind="Teredo", scope="suspicious", embedded_ipv4=str(client))
        port = (~(int(v6) >> 32) & 0xFFFF)
        lines.append(f"Teredo tunnel (2001::/32) — real client IPv4 {client}, UDP port {port}")
        lines.append(f"Teredo server {server}; Teredo is retired — unusual for mail servers")
        if _v4_class(client)[1] == "public":
            out["geo_ip"] = str(client)
        return out
    if v6.sixtofour:
        v4 = v6.sixtofour
        lbl, sc = _v4_class(v4)
        out.update(kind="6to4", embedded_ipv4=str(v4))
        if sc == "public":
            out["scope"] = "public"
            out["geo_ip"] = str(v4)
            lines.append(f"6to4 tunnel (2002::/16) — endpoint IPv4 {v4}")
        elif sc == "suspicious":
            out["scope"] = "suspicious"
            lines.append(f"6to4 format (2002::/16) wrapping non-routable IPv4 {v4} ({lbl})")
        else:
            out["scope"] = "private"
            lines.append(
                f"6to4 format (2002::/16) wrapping private IPv4 {v4} ({lbl}) — "
                "internal addressing, not a public tunnel"
            )
            if "google" in ctx and v4 in ipaddress.ip_network("10.0.0.0/8"):
                out["org_hint"] = "Google internal network"
                out["asn_hint"] = "AS15169"
                lines.append("Gmail uses 2002:a??:… (10.x.x.x) addresses between its internal mail servers")
        lines.extend(_interface_id_notes(v6))
        return out
    if v6 in ipaddress.ip_network("2001:db8::/32"):
        out.update(kind="Documentation", scope="suspicious")
        lines.append("Documentation prefix 2001:db8::/32 — should never appear in real mail (forged?)")
        return out
    if v6 in ipaddress.ip_network("fc00::/7"):
        gid = (int(v6) >> 80) & ((1 << 40) - 1)
        out.update(kind="Unique local (ULA)", scope="private")
        assigned = "locally assigned" if v6.packed[0] == 0xFD else "reserved fc00::/8"
        lines.append(f"Unique local IPv6 (fc00::/7, {assigned}) — private site network, global ID {gid:010x}")
        lines.extend(_interface_id_notes(v6))
        return out
    if v6.is_link_local:
        out.update(kind="Link-local", scope="private")
        lines.append("IPv6 link-local (fe80::/10) — same network segment only")
        lines.extend(_interface_id_notes(v6))
        return out
    if v6.is_multicast:
        out.update(kind="Multicast", scope="special")
        lines.append("IPv6 multicast")
        return out
    if v6 in ipaddress.ip_network("::/96"):
        v4 = ipaddress.IPv4Address(int(v6) & 0xFFFFFFFF)
        out.update(kind="IPv4-compatible", scope="suspicious", embedded_ipv4=str(v4))
        lines.append(f"Deprecated IPv4-compatible IPv6 — embeds {v4}")
        return out

    out["kind"] = "Public IPv6" if v6.is_global else "Special IPv6"
    out["scope"] = "public" if v6.is_global else "special"
    if provider:
        lines.append(f"Inside {provider}'s published range")
        out["org_hint"] = provider
        out["asn_hint"] = provider_asn
    lines.extend(_interface_id_notes(v6))
    return out


def ip_literal(token: Optional[str]) -> Optional[str]:
    """Return the token as a normalised IP string if it is an IP literal."""
    if not token:
        return None
    t = str(token).strip().strip("[]()").rstrip(";,")
    if t.lower().startswith("ipv6:"):
        t = t[5:]
    try:
        return str(ipaddress.ip_address(t))
    except ValueError:
        return None


# ---------------------------------------------------------------------------
# Hostname → location (naming conventions)
# ---------------------------------------------------------------------------

# city key → (city, country ISO2, lat, lon)
_CITIES: Dict[str, Tuple[str, str, float, float]] = {
    "amsterdam": ("Amsterdam", "NL", 52.37, 4.90),
    "dublin": ("Dublin", "IE", 53.35, -6.26),
    "vienna": ("Vienna", "AT", 48.21, 16.37),
    "helsinki": ("Helsinki", "FI", 60.17, 24.94),
    "paris": ("Paris", "FR", 48.86, 2.35),
    "marseille": ("Marseille", "FR", 43.30, 5.37),
    "frankfurt": ("Frankfurt", "DE", 50.11, 8.68),
    "berlin": ("Berlin", "DE", 52.52, 13.40),
    "geneva": ("Geneva", "CH", 46.20, 6.14),
    "zurich": ("Zurich", "CH", 47.38, 8.54),
    "london": ("London", "GB", 51.51, -0.13),
    "cardiff": ("Cardiff", "GB", 51.48, -3.18),
    "oslo": ("Oslo", "NO", 59.91, 10.75),
    "stavanger": ("Stavanger", "NO", 58.97, 5.73),
    "stockholm": ("Stockholm", "SE", 59.33, 18.07),
    "gavle": ("Gävle", "SE", 60.67, 17.14),
    "milan": ("Milan", "IT", 45.46, 9.19),
    "madrid": ("Madrid", "ES", 40.42, -3.70),
    "warsaw": ("Warsaw", "PL", 52.23, 21.01),
    "boydton": ("Boydton, VA", "US", 36.67, -78.39),
    "ashburn": ("Ashburn, VA", "US", 39.04, -77.49),
    "chicago": ("Chicago, IL", "US", 41.88, -87.63),
    "des moines": ("Des Moines, IA", "US", 41.59, -93.62),
    "san antonio": ("San Antonio, TX", "US", 29.42, -98.49),
    "dallas": ("Dallas, TX", "US", 32.78, -96.80),
    "quincy": ("Quincy, WA", "US", 47.23, -119.85),
    "phoenix": ("Phoenix, AZ", "US", 33.45, -112.07),
    "san jose": ("San Jose, CA", "US", 37.34, -121.89),
    "cheyenne": ("Cheyenne, WY", "US", 41.14, -104.82),
    "columbus": ("Columbus, OH", "US", 39.96, -83.00),
    "boardman": ("Boardman, OR", "US", 45.84, -119.70),
    "san francisco": ("San Francisco, CA", "US", 37.77, -122.42),
    "new york": ("New York, NY", "US", 40.71, -74.01),
    "atlanta": ("Atlanta, GA", "US", 33.75, -84.39),
    "miami": ("Miami, FL", "US", 25.76, -80.19),
    "seattle": ("Seattle, WA", "US", 47.61, -122.33),
    "los angeles": ("Los Angeles, CA", "US", 34.05, -118.24),
    "denver": ("Denver, CO", "US", 39.74, -104.99),
    "boston": ("Boston, MA", "US", 42.36, -71.06),
    "toronto": ("Toronto", "CA", 43.65, -79.38),
    "quebec": ("Quebec City", "CA", 46.81, -71.21),
    "montreal": ("Montreal", "CA", 45.50, -73.57),
    "vancouver": ("Vancouver", "CA", 49.28, -123.12),
    "sao paulo": ("São Paulo", "BR", -23.55, -46.63),
    "singapore": ("Singapore", "SG", 1.35, 103.82),
    "hong kong": ("Hong Kong", "HK", 22.32, 114.17),
    "tokyo": ("Tokyo", "JP", 35.68, 139.69),
    "osaka": ("Osaka", "JP", 34.69, 135.50),
    "seoul": ("Seoul", "KR", 37.57, 126.98),
    "kuala lumpur": ("Kuala Lumpur", "MY", 3.14, 101.69),
    "pune": ("Pune", "IN", 18.52, 73.86),
    "mumbai": ("Mumbai", "IN", 19.08, 72.88),
    "chennai": ("Chennai", "IN", 13.08, 80.27),
    "delhi": ("Delhi", "IN", 28.61, 77.21),
    "bangalore": ("Bangalore", "IN", 12.97, 77.59),
    "sydney": ("Sydney", "AU", -33.87, 151.21),
    "melbourne": ("Melbourne", "AU", -37.81, 144.96),
    "dubai": ("Dubai", "AE", 25.20, 55.27),
    "bahrain": ("Manama", "BH", 26.23, 50.59),
    "johannesburg": ("Johannesburg", "ZA", -26.20, 28.05),
    "cape town": ("Cape Town", "ZA", -33.92, 18.42),
}

# Exchange Online first-label datacenter codes.
# value: {region-group or "*": (city key, confidence)}
# Region groups: EU (EUR/GBR/DEU/FRA/CHE/SWE/NOR/...), NA, APAC.
_M365_DC: Dict[str, Dict[str, Tuple[str, str]]] = {
    "AM": {"*": ("amsterdam", "high")},
    "AS": {"EU": ("amsterdam", "medium")},
    "DB": {"*": ("dublin", "high")},
    "DU": {"EU": ("dublin", "medium")},
    "VI": {"*": ("vienna", "high")},
    "VE": {"EU": ("vienna", "low")},
    "HE": {"*": ("helsinki", "high")},
    "PA": {"*": ("paris", "high")},
    "PR": {"EU": ("paris", "medium")},
    "MR": {"EU": ("marseille", "medium")},
    "FR": {"*": ("frankfurt", "high")},
    "BE": {"EU": ("berlin", "medium")},
    "GV": {"*": ("geneva", "high")},
    "ZR": {"*": ("zurich", "high")},
    "LO": {"*": ("london", "high")},
    "LN": {"*": ("london", "high")},
    "CW": {"EU": ("cardiff", "high")},
    "OS": {"EU": ("oslo", "medium"), "APAC": ("osaka", "medium")},
    "ST": {"EU": ("stavanger", "low")},
    "BN": {"*": ("boydton", "high")},
    "BL": {"*": ("boydton", "medium")},
    "BY": {"*": ("boydton", "medium")},
    "CH": {"*": ("chicago", "high")},
    "DM": {"*": ("des moines", "high")},
    "IA": {"NA": ("des moines", "medium")},
    "SA": {"NA": ("san antonio", "high")},
    "SN": {"NA": ("san antonio", "high")},
    "DS": {"NA": ("dallas", "low")},
    "MW": {"NA": ("quincy", "medium")},
    "CO": {"NA": ("quincy", "medium")},
    "PH": {"NA": ("phoenix", "medium")},
    "SJ": {"*": ("san jose", "high")},
    "CY": {"NA": ("cheyenne", "medium")},
    "YT": {"*": ("toronto", "medium")},
    "YQ": {"*": ("quebec", "medium")},
    "SG": {"*": ("singapore", "high")},
    "SI": {"APAC": ("singapore", "medium")},
    "HK": {"*": ("hong kong", "high")},
    "TY": {"*": ("tokyo", "high")},
    "KL": {"APAC": ("kuala lumpur", "medium")},
    "SE": {"APAC": ("seoul", "medium")},
    "SL": {"APAC": ("seoul", "medium")},
    "PN": {"APAC": ("pune", "medium")},
    "PS": {"APAC": ("pune", "medium")},
    "MA": {"APAC": ("chennai", "medium")},
    "BM": {"APAC": ("mumbai", "medium")},
    "SY": {"*": ("sydney", "high")},
    "ME": {"*": ("melbourne", "high")},
}

_M365_REGION_GROUP = {
    "EUR": "EU", "GBR": "EU", "DEU": "EU", "FRA": "EU", "CHE": "EU", "SWE": "EU",
    "NOR": "EU", "POL": "EU", "ITA": "EU", "ESP": "EU",
    "NAM": "NA", "CAN": "NA", "LAM": "NA", "BRA": "NA",
    "APC": "APAC", "JPN": "APAC", "KOR": "APAC", "IND": "APAC", "AUS": "APAC",
    "SGP": "APAC",
}

_M365_ROLES = {
    "MB": "Mailbox server", "CA": "Client Access (front door)", "WS": "Web Services",
    "FE": "Front End", "BE": "Back End", "EP": "Edge Proxy", "BP": "Backend Proxy",
    "DL": "Directory Lookup", "FT": "Frontend Transport (EOP filtering)",
    "PEPF": "Perimeter / Edge Protection front end",
    "EPF": "Edge Protection front end (inbound filtering)",
    "PPF": "PPF-series pooled server",
}

_M365_ENVS = {"PR": "production", "P": "production", "SD": "staging", "T": "test"}

# Region / geo codes in the forest label (EURPRD08, EURP250, GBRP265, eop-nam12).
_M365_REGION_NAMES = {
    "EUR": "Europe", "GBR": "United Kingdom", "DEU": "Germany", "FRA": "France",
    "CHE": "Switzerland", "SWE": "Sweden", "NOR": "Norway", "POL": "Poland",
    "ITA": "Italy", "ESP": "Spain", "NAM": "North America", "CAN": "Canada",
    "LAM": "Latin America", "BRA": "Brazil", "APC": "Asia Pacific", "JPN": "Japan",
    "KOR": "Korea", "IND": "India", "AUS": "Australia", "SGP": "Singapore",
    "ZAF": "South Africa", "ARE": "UAE", "ISR": "Israel", "QAT": "Qatar",
}

_AWS_REGIONS: Dict[str, str] = {
    "us-east-1": "ashburn", "us-east-2": "columbus", "us-west-1": "san francisco",
    "us-west-2": "boardman", "ca-central-1": "montreal", "sa-east-1": "sao paulo",
    "eu-west-1": "dublin", "eu-west-2": "london", "eu-west-3": "paris",
    "eu-central-1": "frankfurt", "eu-central-2": "zurich", "eu-north-1": "stockholm",
    "eu-south-1": "milan", "eu-south-2": "madrid", "me-south-1": "bahrain",
    "me-central-1": "dubai", "af-south-1": "cape town", "ap-south-1": "mumbai",
    "ap-southeast-1": "singapore", "ap-southeast-2": "sydney",
    "ap-northeast-1": "tokyo", "ap-northeast-2": "seoul", "ap-northeast-3": "osaka",
    "ap-east-1": "hong kong",
}

# Airport / metro codes used by carriers and CDNs in router & relay names.
# Only codes that do not collide with common mail-hostname words.
_IATA: Dict[str, str] = {
    "lhr": "london", "lon": "london", "fra": "frankfurt", "ams": "amsterdam",
    "cdg": "paris", "par": "paris", "dub": "dublin", "mad": "madrid",
    "mrs": "marseille", "mxp": "milan", "mil": "milan", "zrh": "zurich",
    "gva": "geneva", "vie": "vienna", "arn": "stockholm", "sto": "stockholm",
    "osl": "oslo", "hel": "helsinki", "waw": "warsaw", "iad": "ashburn",
    "ewr": "new york", "jfk": "new york", "nyc": "new york", "lga": "new york",
    "bos": "boston", "ord": "chicago", "chi": "chicago", "dfw": "dallas",
    "atl": "atlanta", "mia": "miami", "sea": "seattle", "sjc": "san jose",
    "sfo": "san francisco", "lax": "los angeles", "den": "denver",
    "phx": "phoenix", "yyz": "toronto", "yul": "montreal", "yvr": "vancouver",
    "gru": "sao paulo", "nrt": "tokyo", "hnd": "tokyo", "tyo": "tokyo",
    "kix": "osaka", "icn": "seoul", "sin": "singapore", "hkg": "hong kong",
    "syd": "sydney", "mel": "melbourne", "bom": "mumbai", "maa": "chennai",
    "blr": "bangalore", "dxb": "dubai", "jnb": "johannesburg",
}

_MIMECAST_REGIONS = {
    "eu": "Europe (Mimecast EU grid)", "uk": "United Kingdom", "de": "Germany",
    "us": "United States", "ca": "Canada", "za": "South Africa", "au": "Australia",
}


def _place(city_key: str, *, source: str, confidence: str, detail: str) -> Dict[str, Any]:
    city, cc, lat, lon = _CITIES[city_key]
    return {
        "city": city, "country": cc, "lat": lat, "lon": lon,
        "source": source, "confidence": confidence, "detail": detail,
    }


def _parse_forest_label(label: str) -> Optional[Dict[str, str]]:
    """EURPRD08 / EURP250 / eop-nam12 / eop-EUR03 → region, kind, number."""
    lab = (label or "").strip().lower()
    m = re.match(r"^eop-([a-z]{3})(\d+)$", lab)
    if m:
        return {"region": m.group(1).upper(), "kind": "EOP forest", "num": m.group(2)}
    m = re.match(r"^([a-z]{3})(prd|p)(\d+)$", lab)
    if m:
        kind = "production forest" if m.group(2) == "prd" else "production pod"
        return {"region": m.group(1).upper(), "kind": kind, "num": m.group(3)}
    return None


def m365_region_group(forest_label: str) -> Optional[str]:
    parsed = _parse_forest_label(forest_label)
    if not parsed:
        return None
    return _M365_REGION_GROUP.get(parsed["region"])


_MS_MAIL_SUFFIXES = ("outlook.com", "office365.com", "exchangelabs.com", "outlook.office.com")


def decode_m365_host(hostname: Optional[str]) -> Optional[Dict[str, Any]]:
    """Break a Microsoft 365 mail hostname into its constituent parts.

    Handles Exchange Online servers (AM0PR10CA0036, DU0P250MB0793,
    AM1PEPF000252DE, AM7PPF30B196146), EOP filtering nodes
    (BN8NAM12FT012.eop-nam12…), outbound EOP relays
    (mail-am6eur05on2078.outbound…) and tenant MX hosts
    (contoso-com.mail.protection.outlook.com). Datacenter codes are
    Microsoft naming conventions, not published data, so each location
    carries a confidence label.

    Returns None when the name is not recognisably Microsoft 365.
    """
    host = (hostname or "").strip().rstrip(".").lower()
    if not host or ip_literal(host):
        return None
    labels = host.split(".")
    is_ms = any(host == s or host.endswith("." + s) for s in _MS_MAIL_SUFFIXES)
    if not is_ms and len(labels) > 1:
        return None

    out: Dict[str, Any] = {
        "service": "Exchange Online Protection" if "protection.outlook.com" in host else "Exchange Online",
        "dc": None, "dc_num": None, "city_key": None, "confidence": None,
        "role": None, "role_name": None, "serial": None,
        "env": None, "forest": None,
        "region": None, "region_name": None, "forest_label": None, "forest_kind": None,
        "tenant": None,
    }

    first = labels[0].upper()
    patterns = (
        # Classic Exchange Online: AM0PR10CA0036 / DU0P250MB0793 / CWLP265MB1234
        (r"^([A-Z]{2})([0-9A-Z])(PR|P|SD|T)(\d+)(CA|MB|WS|FE|BE|EP|BP|DL)(\d+)$",
         ("dc", "dc_num", "env", "forest", "role", "serial")),
        # Edge / protection front ends: AM1PEPF000252DE, BL02EPF0002992C, AM7PPF30B196146
        (r"^([A-Z]{2})([0-9A-Z]{1,2}?)(PEPF|EPF|PPF)([0-9A-F]+)$",
         ("dc", "dc_num", "role", "serial")),
        # EOP filtering node: BN8NAM12FT012 / DB5EUR03FT045
        (r"^([A-Z]{2})([0-9A-Z])([A-Z]{3})(\d+)(FT)(\d+)$",
         ("dc", "dc_num", "region", "forest", "role", "serial")),
        # Outbound EOP relay: MAIL-AM6EUR05ON2078
        (r"^MAIL-([A-Z]{2})([0-9A-Z])([A-Z]{3})(\d+)(ON)(\d+)$",
         ("dc", "dc_num", "region", "forest", "role", "serial")),
    )
    for rx, names in patterns:
        m = re.match(rx, first)
        if m:
            out.update(dict(zip(names, m.groups())))
            break
    else:
        m = None

    # Server FQDNs also live under mail.protection.outlook.com, so only an
    # unrecognised first label means a tenant MX (contoso-com.mail.protection…).
    m_tenant = re.match(r"^([a-z0-9-]+)\.mail\.protection\.outlook\.com$", host)
    if m_tenant and not m:
        out["tenant"] = m_tenant.group(1)
        out["role"] = "MX"
        out["role_name"] = "tenant inbound MX"
        return out

    if len(labels) > 1:
        parsed = _parse_forest_label(labels[1])
        if parsed:
            out["forest_label"] = labels[1].upper()
            out["forest_kind"] = parsed["kind"]
            out["region"] = out["region"] or parsed["region"]
            out["forest"] = out["forest"] or parsed["num"]

    if out["role"] == "ON":
        out["role_name"] = "outbound relay (EOP → internet)"
    elif out["role"]:
        out["role_name"] = _M365_ROLES.get(out["role"], out["role"])
    if out["env"]:
        out["env"] = _M365_ENVS.get(out["env"], out["env"])
    if out["region"]:
        out["region_name"] = _M365_REGION_NAMES.get(out["region"], out["region"])

    if out["dc"]:
        hit = m365_city(out["dc"], _M365_REGION_GROUP.get(out["region"] or ""))
        if hit:
            out["city_key"], out["confidence"] = hit

    if not (m or out["forest_label"]):
        return None
    return out


def describe_m365_host(hostname: Optional[str]) -> Optional[str]:
    """One-line human breakdown of a Microsoft 365 hostname."""
    d = decode_m365_host(hostname)
    if not d:
        return None
    bits: List[str] = []
    if d["tenant"]:
        guess = d["tenant"].replace("-", ".")
        return f"{d['service']} · tenant MX for '{d['tenant']}' (≈ {guess})"
    if d["dc"]:
        code = f"{d['dc']}{d['dc_num'] or ''}"
        if d["city_key"]:
            city, cc = _CITIES[d["city_key"]][0], _CITIES[d["city_key"]][1]
            conf = "" if d["confidence"] == "high" else f", {d['confidence']} confidence"
            bits.append(f"{city}, {cc} (datacenter {code}{conf})")
        else:
            bits.append(f"datacenter {code} (location not mapped)")
    if d["role_name"]:
        role = d["role_name"]
        if d["serial"]:
            role += f" #{d['serial']}"
        bits.append(role)
    if d["forest_label"]:
        bits.append(f"{d['region_name'] or d['region']} {d['forest_kind']} {d['forest']} ({d['forest_label']})")
    elif d["region"]:
        bits.append(f"{d['region_name']} forest {d['forest']}")
    if d["env"] and d["env"] != "production":
        bits.append(d["env"])
    return f"{d['service']} · " + " · ".join(bits) if bits else None


def m365_city(dc_code: str, region_group: Optional[str]) -> Optional[Tuple[str, str]]:
    """(city key, confidence) for an Exchange Online datacenter code."""
    entry = _M365_DC.get((dc_code or "").upper())
    if not entry:
        return None
    if region_group and region_group in entry:
        return entry[region_group]
    if "*" in entry:
        return entry["*"]
    if region_group is None and len(entry) == 1:
        key, conf = next(iter(entry.values()))
        return key, "low"
    return None


def decode_host_location(hostname: Optional[str]) -> Optional[Dict[str, Any]]:
    """Infer a physical location from provider hostname conventions."""
    if not hostname:
        return None
    host = str(hostname).strip().rstrip(".").lower()
    if not host or ip_literal(host):
        return None
    labels = host.split(".")

    # Exchange Online / EOP: AS2PR08MB9127.eurprd08.prod.outlook.com, AM7PPF30B196146…
    m365 = decode_m365_host(host)
    if m365 and m365.get("city_key"):
        detail = f"Microsoft 365 datacenter {m365['dc']}{m365['dc_num'] or ''}"
        if m365.get("forest_label"):
            detail += f" in {m365['forest_label']}"
        if m365.get("role_name"):
            detail += f" · {m365['role_name']}"
        return _place(
            m365["city_key"], source="Microsoft 365 host name",
            confidence=m365["confidence"], detail=detail,
        )

    # Amazon SES / AWS: e226-3.smtp-out.eu-west-1.amazonses.com
    for lab in labels:
        if lab in _AWS_REGIONS and ("amazonses.com" in host or "amazonaws.com" in host):
            return _place(
                _AWS_REGIONS[lab], source="AWS region in host name", confidence="high",
                detail=f"AWS region {lab}",
            )

    # Mimecast: eu-smtp-inbound-1.mimecast.com
    if host.endswith("mimecast.com") or host.endswith("mimecast.co.za"):
        m_mc = re.match(r"^([a-z]{2})-", labels[0])
        if m_mc and m_mc.group(1) in _MIMECAST_REGIONS:
            return {
                "city": "", "country": m_mc.group(1).upper() if m_mc.group(1) != "eu" else "EU",
                "lat": None, "lon": None, "source": "Mimecast grid prefix",
                "confidence": "medium", "detail": _MIMECAST_REGIONS[m_mc.group(1)],
            }

    # Carrier / CDN airport codes: ae-1.r20.lhr01.example.net, fra-edge-2.example.com
    if len(labels) >= 3:
        for lab in labels[:-2]:
            for tok in re.split(r"[-_]", lab):
                m_tok = re.match(r"^([a-z]{3})\d{0,3}$", tok)
                if m_tok and m_tok.group(1) in _IATA:
                    code = m_tok.group(1)
                    return _place(
                        _IATA[code], source="Airport code in host name", confidence="low",
                        detail=f"'{code.upper()}' metro code",
                    )
    return None


def location_label(loc: Dict[str, Any]) -> str:
    """Short 'City, CC (inferred, conf)' label for display."""
    city = loc.get("city") or ""
    cc = loc.get("country") or ""
    place = ", ".join(p for p in (city, cc) if p)
    return f"{place} — {loc.get('detail', '')} ({loc.get('confidence', 'low')} confidence)"
