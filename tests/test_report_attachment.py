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


def test_status_indicators_show_both_values(tmp_path, monkeypatch):
    monkeypatch.setattr(gh, "aspr_active", lambda: True)
    monkeypatch.setattr(gh, "readnotify_category", lambda: "")
    flow = gh._footer_status_flow("#d5dee4")
    assert "ASPR:" in flow and ">ACTIVE<" in flow and "RN:" in flow and ">OFF<" in flow
    assert "<td" not in flow and "<table" not in flow          # inline, so the row can wrap
    monkeypatch.setattr(gh, "aspr_active", lambda: False)
    monkeypatch.setattr(gh, "readnotify_category", lambda: "Certified")
    flow = gh._footer_status_flow("#d5dee4")
    assert ">NOT&nbsp;ACTIVE<" in flow and ">CERTIFIED<" in flow


# --- reflowing strip ----------------------------------------------------------

def test_flow_separators_break_only_after_dots():
    out = gh.flow_separators(
        "Scan Result - Basic | <a href=\"x\" style=\"color:#fff; line-height:12px; vertical-align:middle;\">Risk LOW</a>"
        " | Links: %% <font style='color:#0f0; line-height:16px;'>12 OK</font>",
        color="#ffffff", weight="bold",
    )
    assert out.count("&#9679;") == 2                             # one dot between each pair
    assert "Scan&nbsp;Result&nbsp;-&nbsp;Basic" in out and "Risk&nbsp;LOW" in out
    assert "Links:&nbsp;<font" in out and "12&nbsp;OK" in out
    assert "vertical-align" not in out and "line-height" not in out
    assert 'href="x"' in out                                     # tags and attributes untouched
    assert "&#9679;</span> <span" in out                         # the only plain space is after a dot


def test_strip_has_no_fixed_width_content(monkeypatch):
    """Nothing but the logo may pin the strip's width, or a narrow window cuts it off."""
    cls = next(v for v in vars(gh).values() if isinstance(v, type) and "_build_html" in vars(v))
    obj = object.__new__(cls)
    for name in ("_write_risk_report_file", "_write_full_scan_report_file", "_build_full_details_html"):
        monkeypatch.setattr(cls, name, lambda *a, **k: "", raising=False)
    monkeypatch.setattr(gh, "aspr_active", lambda: True)
    monkeypatch.setattr(gh, "readnotify_category", lambda: "")
    page = obj._build_html({
        "risk_level": "LOW", "risk_score": 20, "hre": 1,
        "summary_line1_html": "Aliniant AES Scan Result - Basic Scan | {{AES_RISK}} | {{AES_HRE}}",
        "summary_line2_html": "Sender IP 203.0.113.9 | Testville | {{AES_FULLSCAN}}",
        "sender_email": "someone@example.org", "sender_domain": "example.org", "guri": "T",
    }, "compact")
    import re
    rows = page[page.index("aes-footer-start"):]
    assert "nowrap" not in rows
    widths = [int(w) for w in re.findall(r"<td[^>]* width='([0-9]+)'", rows)]
    assert max(widths) <= 78                                      # logo cell and the AURA chip only
    for code in ("SL", "BA", "BB", "TS", "NT"):
        if code != "SL":                                          # SL needs a links report
            assert f"{code}&nbsp;" in rows
    assert "&#8203;" in rows                                      # break points between chips


# --- footer mark named after its content --------------------------------------

def test_mark_file_name_follows_its_content(tmp_path, monkeypatch):
    monkeypatch.setenv("LOCALAPPDATA", str(tmp_path))
    first = tmp_path / "a.png"
    first.write_bytes(b"design one")
    one = gh._versioned_mark_path(first)
    assert one.name.startswith("aes_mark_footer_") and one.suffix == ".png"
    assert one.read_bytes() == b"design one"
    assert gh._versioned_mark_path(first) == one                 # unchanged image, same name
    first.write_bytes(b"design two")
    two = gh._versioned_mark_path(first)
    assert two != one and two.read_bytes() == b"design two"      # redesign, new name


def test_footer_marker_points_at_the_versioned_file(tmp_path, monkeypatch):
    monkeypatch.setenv("LOCALAPPDATA", str(tmp_path))
    out = gh._footer_mark_html()
    assert "cid:aesfootermark" in out
    import re
    path = re.search(r"AES-Mark-Img: (.*?) -->", out).group(1)
    assert Path(path).name.startswith("aes_mark_footer_") and Path(path).is_file()


# --- origin from the relay hops when the headers name no sender IP -------------

INTERNAL_HOPS = [
    {"ip": "fe80::72ab:5895:eaa8:7a9c (internal)", "fqdn": "AM0PR08MB4275.eurprd08.prod.outlook.com",
     "hop_kind": "internal", "asn": "\u2014", "city": "Amsterdam", "country": "NL"},
    {"ip": "2603:10a6:208:140::11", "fqdn": "AM0PR08MB4275.eurprd08.prod.outlook.com",
     "hop_kind": "public", "asn": "AS8075", "city": "Amsterdam", "country": "NL"},
    {"fqdn": "DB8PR08MB5355.eurprd08.prod.outlook.com", "hop_kind": "unknown", "city": "Dublin", "country": "IE"},
]


def test_origin_uses_the_first_public_hop():
    origin = gh.origin_from_hops(INTERNAL_HOPS)
    assert origin["ip"] == "2603:10a6:208:140::11"
    assert origin["host"] == "AM0PR08MB4275.eurprd08.prod.outlook.com"
    assert (origin["city"], origin["country"], origin["asn"]) == ("Amsterdam", "NL", "AS8075")
    assert origin["org"]                                          # filled from the host name


def test_origin_without_a_public_hop_gives_location_but_no_ip():
    origin = gh.origin_from_hops([INTERNAL_HOPS[0], INTERNAL_HOPS[2]])
    assert origin["ip"] == "" and origin["city"] == "Amsterdam"
    assert gh.origin_from_hops([]) == {} and gh.origin_from_hops([{"hop_kind": "unknown"}]) == {}


def test_share_note_sits_inside_the_report_card(tmp_path):
    src = tmp_path / "fullscan_card.html"
    src.write_text('<html><body><div class="wrap"><h1>Full Analysis</h1></div></body></html>', encoding="utf-8")
    page = Path(gh.write_shareable_report(str(src))).read_text(encoding="utf-8")
    note = page.index("produced by Aliniant Email Scanner")
    assert page.index('<div class="wrap">') < note < page.index("<h1>")



def test_attachment_count_is_what_the_reader_sees_not_every_inline_picture():
    import aes.geolocate_headers as gh

    cls = next(v for v in vars(gh).values()
               if isinstance(v, type) and hasattr(v, "_visible_attachment_counts"))
    me = object.__new__(cls)

    def shown(counts, summary):
        return cls._visible_attachment_counts(me, counts, {"aes-attachments": summary})

    # Nine scanned parts, none in the attachment well.
    assert shown((9, 9, 0), "count=9; visible=0; ok=0; not_ok=0") == (0, 0, 0)
    assert shown((9, 9, 0), "count=9; visible=2; ok=0; not_ok=0") == (2, 2, 0)
    # A part that fails its scan is counted even when it hides inline.
    assert shown((9, 8, 1), "count=9; visible=0; ok=0; not_ok=0") == (1, 0, 1)
    # Older export without the field: unchanged.
    assert shown((9, 9, 0), "count=9; ok=0; not_ok=0") == (9, 9, 0)
