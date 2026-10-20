"""Hop ASN comes from the sending side, never the receiving by-host."""

from aes.geolocate_headers import HTMLReportGenerator


def _gen(rdap_asn: str):
    gen = HTMLReportGenerator()

    def fake_rdap(hop, ip):
        hop["asn"] = rdap_asn
        hop["asn_bgp"] = rdap_asn
        hop["asn_org"] = "registry org"
        hop["prefix"] = "149.72.0.0/16"
        hop["rir"] = "ARIN"

    gen._apply_rdap_fields = fake_rdap
    gen._ripestat_prefix = lambda ip: {}
    gen._whois_ip_fallback = lambda ip: {}
    return gen


def test_google_by_host_does_not_replace_sendgrid_asn():
    hop = {
        "hop_kind": "public",
        "fqdn": "o1.ptr8966.carto.com",
        "helo": "",
        "by_host": "mx.google.com",
        "asn": "",
    }
    _gen("AS11377")._fill_ip_registration(hop, "149.72.58.126")
    assert hop["asn"] == "AS11377"


def test_sender_hostname_still_beats_bgp_transit():
    hop = {
        "hop_kind": "public",
        "fqdn": "mta12.sailthru.com",
        "helo": "",
        "by_host": "mx.google.com",
        "asn": "",
    }
    _gen("AS16509")._fill_ip_registration(hop, "163.47.180.10")
    assert hop["asn"] == "AS18877"


def test_sendgrid_net_maps_to_sendgrid():
    from aes.addr_decode import provider_asn_for_host

    assert provider_asn_for_host("xvfrmmbr.outbound-mail.sendgrid.net")[0] == "AS11377"
