"""BGP and tromboning evidence must name the hops involved and reach the report."""

import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

import aes.geolocate_headers as gh  # noqa: E402
import aes.route_checks as rc  # noqa: E402

LONDON, DUBLIN, SINGAPORE = (51.5, -0.12), (53.35, -6.26), (1.35, 103.82)


def _hop(pt, cc, name):
    return {"hop_kind": "public", "lat": pt[0], "lon": pt[1], "country": cc, "fqdn": name}


def _detour():
    return [_hop(LONDON, "GB", "a.example"), _hop(SINGAPORE, "SG", "b.example"), _hop(DUBLIN, "IE", "c.example")]


def test_tromboning_names_all_three_hops():
    hops = _detour()
    findings = rc.check_tromboning(hops)
    assert len(findings) == 1
    assert findings[0]["kind"] == "tromboning"
    assert findings[0]["hops"] == [1, 2, 3]
    assert all(h.get("route_flags") for h in hops)
    assert "leaves the region" in hops[0]["route_flags"][0]
    assert "returns to the region" in hops[2]["route_flags"][0]


def test_direct_path_is_clean_only_when_something_was_checked():
    hops = [_hop(LONDON, "GB", "a"), _hop(DUBLIN, "IE", "b"), _hop((48.85, 2.35), "FR", "c")]
    assert rc.check_tromboning(hops) == []
    assert rc.integrity_summary(hops, [], lite=True)["clean"] is True
    assert rc.integrity_summary([{"hop_kind": "internal"}], [], lite=True)["clean"] is False


def test_bgp_findings_carry_their_hop(monkeypatch):
    hops = [{"hop_kind": "internal"}, {"hop_kind": "public", "ip": "198.51.100.7", "exit_hop": True, "asn": "AS64500"}]
    monkeypatch.setattr(rc, "_public_ip", lambda h: h.get("ip"))
    monkeypatch.setattr(rc, "bgp_origin_info", lambda ip, want_peers=True: {
        "prefix": "198.51.100.0/24", "origins": ["64501", "64502"], "holders": {},
        "rpki": {"64501": "invalid_asn"}, "upstreams": [],
    })
    findings = rc.check_bgp(hops, lite=True)
    kinds = {f["kind"]: f["hops"] for f in findings}
    assert kinds == {"bgp-rpki-invalid": [2], "bgp-multiple-origins": [2], "bgp-origin-mismatch": [2]}
    summary = rc.integrity_summary(hops, findings, lite=True)
    assert [i["kind"] for i in summary["items"]][0] == "bgp-rpki-invalid"   # strongest first
    assert summary["bgp_checked"] == 1 and summary["clean"] is False


def test_report_block_names_hops_and_states_coverage():
    hops = _detour()
    findings = rc.check_tromboning(hops)
    block = gh._route_integrity_html(rc.integrity_summary(hops, findings, lite=True))
    assert "Route Integrity" in block and "Hops 1 \u2192 2 \u2192 3" in block
    assert "Tromboning" in block and "exit hop only" in block
    clean = gh._route_integrity_html({"items": [], "located": 4, "bgp_checked": 1, "lite": False, "clean": True})
    assert "No BGP hijack or tromboning evidence found" in clean and "4 located" in clean
    unknown = gh._route_integrity_html({"items": [], "located": 1, "bgp_checked": 0, "lite": True, "clean": False})
    assert "Not enough route data" in unknown
    assert gh._route_integrity_html({}) == ""
