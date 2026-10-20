"""Tests for the AES read-tracking pixel (server + client side)."""

from __future__ import annotations

import importlib.util
import json
import os
import sys
import threading
import urllib.error
import urllib.request
from http.server import ThreadingHTTPServer
from pathlib import Path

import pytest

ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

PIXEL = "0123456789abcdef0123456789abcdef"


@pytest.fixture()
def tracking(tmp_path, monkeypatch):
    monkeypatch.setenv("LOCALAPPDATA", str(tmp_path))
    import aes.pixel_tracking as pt

    return pt


@pytest.fixture()
def server(tmp_path, monkeypatch):
    monkeypatch.setenv("AES_PIXEL_KEY", "k" * 40)
    monkeypatch.setenv("AES_PIXEL_DB", str(tmp_path / "srv" / "opens.sqlite"))
    spec = importlib.util.spec_from_file_location(
        "aes_pixel_server", ROOT / "cloud" / "pixel" / "aes_pixel_server.py"
    )
    mod = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(mod)
    mod._init_db()
    httpd = ThreadingHTTPServer(("127.0.0.1", 0), mod.Handler)
    thread = threading.Thread(target=httpd.serve_forever, daemon=True)
    thread.start()
    yield f"http://127.0.0.1:{httpd.server_address[1]}"
    httpd.shutdown()


def _get(url, headers=None):
    req = urllib.request.Request(url, headers=headers or {})
    try:
        with urllib.request.urlopen(req, timeout=5) as resp:
            return resp.status, resp.headers, resp.read()
    except urllib.error.HTTPError as exc:
        return exc.code, exc.headers, exc.read()


def test_server_records_pixel_and_returns_gif(server):
    status, headers, body = _get(f"{server}/p/{PIXEL}.gif", {"User-Agent": "Microsoft Office/16.0"})
    assert status == 200
    assert headers["Content-Type"] == "image/gif"
    assert "no-store" in headers["Cache-Control"]
    assert body.startswith(b"GIF89a")

    status, _, raw = _get(f"{server}/opens?since=0", {"X-AES-Key": "k" * 40})
    data = json.loads(raw)
    assert status == 200
    assert [o["pixel"] for o in data["opens"]] == [PIXEL]
    assert data["opens"][0]["ua"] == "Microsoft Office/16.0"
    assert data["client_ip"] == "127.0.0.1"


def test_server_ignores_malformed_id_but_still_serves_gif(server):
    status, _, body = _get(f"{server}/p/not-a-pixel.gif")
    assert status == 200 and body.startswith(b"GIF89a")
    _, _, raw = _get(f"{server}/opens", {"X-AES-Key": "k" * 40})
    assert json.loads(raw)["opens"] == []


def test_server_opens_requires_key(server):
    assert _get(f"{server}/opens")[0] == 401
    assert _get(f"{server}/opens", {"X-AES-Key": "wrong"})[0] == 401


def test_classify_labels_proxies_and_prefetch(tracking):
    gmail = tracking.classify("66.249.84.1", "Mozilla/5.0 (via ggpht.com GoogleImageProxy)")
    assert gmail.proxied and not gmail.prefetch
    apple = tracking.classify("172.225.0.1", "Mozilla/5.0")
    assert apple.prefetch
    outlook = tracking.classify("81.2.69.160", "Microsoft Office/16.0 (Windows NT 10.0; Microsoft Outlook 16.0)")
    assert outlook.client == "Outlook" and not outlook.proxied


def test_own_ip_views_are_excluded_even_when_seen_earlier(tracking):
    conn = tracking._connect()
    conn.execute(
        "INSERT INTO sends (pixel, sender, recipients, subject, sent_local) VALUES (?, ?, ?, ?, ?)",
        (PIXEL, "julian.garrett@aliniant.com", "a@b.com", "Quote", "2026-10-20 18:00:00"),
    )
    tracking.store_opens(conn, {"client_ip": "", "opens": [
        {"id": 1, "pixel": PIXEL, "ts": 100, "ip": "5.5.5.5", "ua": "Microsoft Office/16.0"},
        {"id": 2, "pixel": PIXEL, "ts": 200, "ip": "9.9.9.9", "ua": "Microsoft Office/16.0"},
    ]})
    tracking.store_opens(conn, {"client_ip": "5.5.5.5", "opens": []})
    conn.commit()
    conn.close()

    (row,) = tracking.summaries()
    assert row["opens"] == 1
    assert row["first_open"] == 200
    own = [h for h in tracking.opens_for(PIXEL) if h["own"]]
    assert [h["ip"] for h in own] == ["5.5.5.5"]


def test_ledger_import_strips_readnotify_suffix(tracking):
    folder = tracking.ledger_dir()
    folder.mkdir(parents=True)
    (folder / f"{PIXEL}.json").write_text(
        json.dumps({
            "id": PIXEL, "sent_local": "2026-10-20 18:00:00",
            "sender": "julian.garrett@aliniant.com",
            "to": "bob@example.com.silent.read-notify.com", "cc": "", "bcc": "",
            "subject": "Quote — 20 Oct",
        }),
        encoding="utf-8-sig",
    )
    conn = tracking._connect()
    assert tracking.import_ledger(conn) == 1
    assert tracking.import_ledger(conn) == 0
    conn.commit()
    conn.close()
    (send,) = tracking.sends_without_guri()
    assert send["recipients"] == "bob@example.com"
    assert send["subject"] == "Quote — 20 Oct"


def test_config_defaults_off(tracking):
    assert tracking.load_config() == {"enabled": False, "base_url": tracking.DEFAULT_BASE_URL}
    tracking.config_path().write_text("enabled=1\nbase_url=https://track.example.com/\n", encoding="utf-8")
    assert tracking.load_config() == {"enabled": True, "base_url": "https://track.example.com"}
