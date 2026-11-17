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
