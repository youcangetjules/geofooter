"""The report attached to scanned mail must work for someone without AES."""

import json
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

import aes.geolocate_headers as gh  # noqa: E402

REPORT = """<!DOCTYPE html><html><head><title>t</title></head><body class="x">
<h1>Full Analysis</h1>
<p>Sender IP 203.0.113.9</p>
<a href="aes://open-report?path=C%3A%5Cx.html" target="_blank" style="color:#fff">Risk LOW (20/100)</a>
<a href='aes://block-beacons?sender=a%40b.c'>BB</a>
<a href="https://example.org/kept">kept</a>
</body></html>"""


def _report(tmp_path):
    src = tmp_path / "fullscan_abc123.html"
    src.write_text(REPORT, encoding="utf-8")
    return src


def test_share_copy_has_no_aes_links_and_keeps_content(tmp_path):
    src = _report(tmp_path)
    out = Path(gh.write_shareable_report(src.as_uri()))
    assert out.name == gh.AES_REPORT_ATTACH_NAME
    assert out.parent == tmp_path / "share" / "fullscan_abc123"
    page = out.read_text(encoding="utf-8")
    assert "aes://" not in page
    assert "Risk LOW (20/100)" in page and ">BB<" not in page and "BB" in page
    assert 'href="https://example.org/kept"' in page
    assert "Sender IP 203.0.113.9" in page
    assert page.index("produced by Aliniant Email Scanner") > page.index("<body")
    assert src.read_text(encoding="utf-8") == REPORT          # the original is untouched


def test_plain_path_is_accepted_and_missing_file_gives_nothing(tmp_path):
    src = _report(tmp_path)
    assert gh.write_shareable_report(str(src)).endswith(gh.AES_REPORT_ATTACH_NAME)
    assert gh.write_shareable_report(str(tmp_path / "nope.html")) == ""
    assert gh.write_shareable_report("") == ""


def test_off_switch(tmp_path, monkeypatch):
    monkeypatch.setenv("LOCALAPPDATA", str(tmp_path))
    assert gh.report_attachment_enabled() is True              # default
    cfg = tmp_path / "GeoFooter"
    cfg.mkdir()
    (cfg / "aes_report_attachment.json").write_text(json.dumps({"enabled": False}), encoding="utf-8")
    assert gh.report_attachment_enabled() is False
    (cfg / "aes_report_attachment.json").write_text("not json", encoding="utf-8")
    assert gh.report_attachment_enabled() is True


# --- host name on the report's IP line ----------------------------------------

def test_hostname_comes_from_the_matching_hop():
    hops = [
        {"ip": "198.51.100.7", "reverse_dns": "Unknown", "hostname": "mx1.example.net."},
        {"ip": "203.0.113.9 (internal)", "reverse_dns": "relay.example.org"},
    ]
    assert gh.hostname_for_ip("198.51.100.7", hops) == "mx1.example.net"
    assert gh.hostname_for_ip("203.0.113.9", hops) == "relay.example.org"


def test_hostname_falls_back_to_reverse_dns_and_caches(monkeypatch):
    import dns.resolver

    calls = []

    class FakeResolver:
        lifetime = timeout = 0

        def resolve(self, name, kind):
            calls.append((str(name), kind))
            return ["host.example.com."]

    monkeypatch.setattr(dns.resolver, "Resolver", FakeResolver)
    gh._PTR_CACHE.clear()
    assert gh.hostname_for_ip("192.0.2.44", []) == "host.example.com"
    assert gh.hostname_for_ip("192.0.2.44", []) == "host.example.com"
    assert calls == [("44.2.0.192.in-addr.arpa.", "PTR")]


def test_no_hostname_is_empty_not_an_error(monkeypatch):
    import dns.resolver

    class Failing:
        lifetime = timeout = 0

        def resolve(self, name, kind):
            raise dns.resolver.NXDOMAIN()

    monkeypatch.setattr(dns.resolver, "Resolver", Failing)
    gh._PTR_CACHE.clear()
    assert gh.hostname_for_ip("192.0.2.45", None) == ""
    assert gh.hostname_for_ip("", None) == ""


# --- ASPR / RN indicators on the footer's bottom row ---------------------------

def _settings(tmp_path, monkeypatch, *, pixel=None, classify=None):
    monkeypatch.setenv("LOCALAPPDATA", str(tmp_path))
    folder = tmp_path / "GeoFooter"
    folder.mkdir(exist_ok=True)
    if pixel is not None:
        (folder / "aes_pixel.ini").write_text(pixel, encoding="utf-8")
    if classify is not None:
        (folder / "classify_result.json").write_text(json.dumps(classify), encoding="utf-8")


def test_readnotify_category_from_last_send(tmp_path, monkeypatch):
    _settings(tmp_path, monkeypatch)
    assert gh.readnotify_category() == ""                               # no file
    _settings(tmp_path, monkeypatch, classify={"read_notify": False, "read_notify_options": {"certified": True}})
    assert gh.readnotify_category() == ""                               # off wins
    _settings(tmp_path, monkeypatch, classify={"read_notify": True, "read_notify_options": {"activate": True}})
    assert gh.readnotify_category() == "Tracking"
    _settings(tmp_path, monkeypatch, classify={"read_notify": True, "read_notify_options": {
        "activate": True, "certified": True, "block_print": True}})
    assert gh.readnotify_category() == "Certified + No-Print"
    _settings(tmp_path, monkeypatch, classify={"read_notify": True, "read_notify_options": {
        "activate": True, "invisible": True, "certified": True}})
    assert gh.readnotify_category() == "Invisible"                      # cannot be combined


def test_status_cells_show_both_indicators(tmp_path, monkeypatch):
    monkeypatch.setattr(gh, "aspr_active", lambda: True)
    monkeypatch.setattr(gh, "readnotify_category", lambda: "")
    cells = gh._footer_status_cells("#d5dee4")
    assert "ASPR:" in cells and "&nbsp;ACTIVE" in cells and "RN:" in cells and "&nbsp;OFF" in cells
    monkeypatch.setattr(gh, "aspr_active", lambda: False)
    monkeypatch.setattr(gh, "readnotify_category", lambda: "Certified")
    cells = gh._footer_status_cells("#d5dee4")
    assert "&nbsp;NOT ACTIVE" in cells and "&nbsp;CERTIFIED" in cells

