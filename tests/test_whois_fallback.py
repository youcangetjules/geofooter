"""Registrar lookup must survive a broken whois library and TLDs with no RDAP."""

import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

import aes.geolocate_headers as gh  # noqa: E402

NIC_US = """Domain Name: zoom.us
Registry Domain ID: D123-US
Registrar: MarkMonitor, Inc.
Creation Date: 2002-04-24T15:03:39Z
Registry Expiry Date: 2027-04-23T23:59:59Z
"""


def _service(monkeypatch, answers):
    svc = gh.WhoisService()
    asked = []

    def fake(server, text, timeout=6.0):
        asked.append((server, text))
        value = answers[server]
        if isinstance(value, Exception):
            raise value
        return value

    monkeypatch.setattr(gh.WhoisService, "_port43_query", staticmethod(fake))
    return svc, asked


def test_direct_whois_follows_the_iana_referral(monkeypatch):
    svc, asked = _service(monkeypatch, {"whois.iana.org": "domain: US\nwhois:        whois.nic.us\n", "whois.nic.us": NIC_US})
    out = svc._port43_whois("zoom.us")
    assert asked == [("whois.iana.org", "us"), ("whois.nic.us", "zoom.us")]
    assert out == {"registrar": "MarkMonitor, Inc.", "creation_date": "2002-04-24",
                   "expiration_date": "2027-04-23", "status": "Unknown"}


def test_direct_whois_gives_up_quietly(monkeypatch):
    svc, _ = _service(monkeypatch, {"whois.iana.org": "domain: ZZ\n"})
    assert svc._port43_whois("example.zz") is None                 # no referral
    svc, _ = _service(monkeypatch, {"whois.iana.org": OSError("blocked")})
    assert svc._port43_whois("example.us") is None                 # network refused
    svc, _ = _service(monkeypatch, {"whois.iana.org": "whois: whois.nic.us\n", "whois.nic.us": "No Data Found\n"})
    assert svc._port43_whois("nothing.us") is None                 # nothing parseable
    assert svc._port43_whois("bad domain!") is None


def test_broken_library_and_missing_rdap_fall_through_to_direct_whois(monkeypatch):
    class WrongWhois:                                              # the other "whois" package
        query = staticmethod(lambda domain: None)

    monkeypatch.setattr(gh, "whois", WrongWhois)
    monkeypatch.setattr(gh.WhoisService, "_rdap_whois", lambda self, domain: None)
    svc, asked = _service(monkeypatch, {"whois.iana.org": "whois: whois.nic.us\n", "whois.nic.us": NIC_US})
    out = svc.get_whois_info("zoom.us")
    assert out["registrar"] == "MarkMonitor, Inc." and out["creation_date"] == "2002-04-24"
