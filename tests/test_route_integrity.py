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


def test_banner_always_says_scan_done_with_an_assessment():
    assert rc.banner_note({"items": [], "bgp_checked": 1}) == ("BGP scan: no risks", False)
    nothing = ("BGP scan: n/a - internal route", False)
    assert rc.banner_note({"items": [], "bgp_checked": 0}) == nothing
    assert rc.banner_note(None) == nothing
    # Asked, no answer: never passed off as clean.
    assert rc.banner_note({"items": [], "bgp_checked": 0, "bgp_attempted": 1}) == (
        "BGP scan: UNVERIFIED - no reply", True)


def _public(ip, **extra):
    return dict({"hop_kind": "public", "ip": ip, "fqdn": "mx.example"}, **extra)


def test_basic_scan_assesses_the_first_internet_hop_when_no_exit_hop_is_marked(monkeypatch):
    asked = []

    def fake(ip, *, want_peers=True):
        asked.append(ip)
        return {"prefix": "8.8.8.0/24", "origins": ["15169"], "holders": {}, "rpki": {"15169": "valid"},
                "upstreams": []}

    monkeypatch.setattr(rc, "bgp_origin_info", fake)
    hops = [{"hop_kind": "internal", "ip": "10.0.0.1"}, _public("8.8.8.8"), _public("1.1.1.1")]
    findings = rc.check_bgp(hops, lite=True)
    assert asked == ["8.8.8.8"]
    summary = rc.integrity_summary(hops, findings, lite=True)
    assert (summary["bgp_attempted"], summary["bgp_checked"]) == (1, 1)
    assert rc.banner_note(summary)[0] == "BGP scan: no risks"


def test_marked_exit_hop_is_still_preferred(monkeypatch):
    asked = []
    monkeypatch.setattr(rc, "bgp_origin_info", lambda ip, *, want_peers=True: asked.append(ip))
    rc.check_bgp([_public("8.8.8.8"), _public("1.1.1.1", exit_hop=True)], lite=True)
    assert asked == ["1.1.1.1"]


def test_unanswered_lookup_is_reported_and_not_cached(monkeypatch):
    monkeypatch.setattr(rc, "bgp_origin_info", lambda ip, *, want_peers=True: None)
    hops = [_public("8.8.8.8")]
    summary = rc.integrity_summary(hops, rc.check_bgp(hops, lite=True), lite=True)
    assert rc.banner_note(summary) == ("BGP scan: UNVERIFIED - no reply", True)

    import requests
    calls = []

    def boom(*a, **k):
        calls.append(1)
        raise requests.ConnectionError("down")

    monkeypatch.setattr(requests, "get", boom)
    monkeypatch.setattr(rc, "_CACHE", {})
    assert rc._ripestat("network-info", "8.8.8.8") is None
    assert len(calls) == 3 and rc._CACHE == {}       # retried, failure not remembered


def test_banner_names_the_strongest_bgp_risk_and_warns():
    summary = {"bgp_checked": 2, "items": [
        {"kind": "bgp-rpki-invalid", "hops": [1]}, {"kind": "tromboning", "hops": [1, 2, 3]},
        {"kind": "bgp-origin-mismatch", "hops": [1]}]}
    assert rc.banner_note(summary) == ("BGP scan: RISK - hijack suspected (+1 more)", True)


def test_banner_treats_an_origin_mismatch_alone_as_an_advisory():
    summary = {"bgp_checked": 1, "items": [{"kind": "bgp-origin-mismatch", "hops": [1]}]}
    assert rc.banner_note(summary) == ("BGP scan: no risks (1 advisory)", False)


def test_status_strip_carries_the_bgp_result(monkeypatch):
    seen = {}

    def fake_render(self, line, *, warn, risk_score=0):
        seen["line"], seen["warn"] = line, warn
        return "", 0, 0

    cls = next(v for v in vars(gh).values()
               if isinstance(v, type) and hasattr(v, "_build_top_of_mail_banners"))
    monkeypatch.setattr(cls, "_render_status_banner_png", fake_render)
    monkeypatch.setattr(cls, "_sender_status_label", lambda *a, **k: "Known/Trust")
    monkeypatch.setattr(cls, "_beacon_blocking_enabled", lambda self: False)
    me = object.__new__(cls)
    kwargs = dict(block_state={"trusted": True}, attachment_count=0, beacon_count=0, beacon_blocked=None,
                  attachment_not_ok=0, attachments_quarantined=0, link_findings=[], risk_score=5)
    cls._build_top_of_mail_banners(me, route_integrity={"items": [], "bgp_checked": 1}, **kwargs)
    assert seen["line"].endswith("BGP scan: no risks") and seen["warn"] is False
    cls._build_top_of_mail_banners(
        me, route_integrity={"bgp_checked": 1, "items": [{"kind": "bgp-rpki-invalid"}]}, **kwargs)
    assert "BGP scan: RISK - hijack suspected" in seen["line"] and seen["warn"] is True


def test_origin_comes_from_dns_and_unread_rpki_is_not_called_clean(monkeypatch):
    monkeypatch.setattr(rc, "_CACHE", {})
    monkeypatch.setattr(rc, "_dns_origin", lambda ip: ("8.8.8.0/24", ["15169"]))
    monkeypatch.setattr(rc, "_ripestat", lambda *a, **k: None)      # registry silent
    info = rc.bgp_origin_info("8.8.8.8", want_peers=False)
    assert info["prefix"] == "8.8.8.0/24" and info["origins"] == ["15169"] and info["rpki_unknown"]
    hops = [_public("8.8.8.8")]
    summary = rc.integrity_summary(hops, rc.check_bgp(hops, lite=True), lite=True)
    assert rc.banner_note(summary) == ("BGP scan: origin OK, RPKI unverified", False)


def test_registry_calls_stop_at_the_scan_deadline(monkeypatch):
    import requests
    calls = []
    monkeypatch.setattr(requests, "get", lambda *a, **k: calls.append(1) or (_ for _ in ()).throw(requests.Timeout()))
    monkeypatch.setattr(rc, "_CACHE", {})
    monkeypatch.setattr(rc, "_DEADLINE", rc.time.monotonic() - 1)
    assert rc._ripestat("rpki-validation", "AS1", prefix="1.0.0.0/24") is None and calls == []
