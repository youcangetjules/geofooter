#!/usr/bin/env python3
"""AES read-tracking pixel endpoint (stdlib only).

Runs behind nginx on a loopback port. Two routes:

  GET /p/<32 hex>.gif   1x1 transparent GIF; records time, client IP, user agent
  GET /opens?since=N    rows with id > N as JSON; needs header X-AES-Key
  GET /healthz          "ok"

Configuration comes from the environment (systemd EnvironmentFile):

  AES_PIXEL_KEY    shared secret for /opens (required)
  AES_PIXEL_DB     SQLite path (default /var/lib/aes-pixel/opens.sqlite)
  AES_PIXEL_PORT   loopback port (default 8791)

The service makes no outbound connections.
"""

from __future__ import annotations

import hmac
import json
import os
import re
import sqlite3
import threading
import time
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from urllib.parse import parse_qs, urlparse

GIF = (
    b"GIF89a\x01\x00\x01\x00\x80\x00\x00\x00\x00\x00\xff\xff\xff!\xf9\x04\x01"
    b"\x00\x00\x00\x00,\x00\x00\x00\x00\x01\x00\x01\x00\x00\x02\x02D\x01\x00;"
)
PIXEL_PATH = re.compile(r"^/p/([0-9a-f]{32})\.gif$")
RETENTION_DAYS = 400
# Caps a single pixel id so a scraper hammering one URL cannot fill the disk.
MAX_HITS_PER_PIXEL = 500
MAX_ROWS_PER_POLL = 1000

DB_PATH = os.environ.get("AES_PIXEL_DB", "/var/lib/aes-pixel/opens.sqlite")
API_KEY = os.environ.get("AES_PIXEL_KEY", "")
PORT = int(os.environ.get("AES_PIXEL_PORT", "8791"))

_lock = threading.Lock()
_inserts = 0


def _connect() -> sqlite3.Connection:
    conn = sqlite3.connect(DB_PATH, timeout=5)
    conn.execute("PRAGMA journal_mode=WAL")
    return conn


def _init_db() -> None:
    os.makedirs(os.path.dirname(DB_PATH) or ".", exist_ok=True)
    with _connect() as conn:
        conn.execute(
            "CREATE TABLE IF NOT EXISTS opens ("
            " id INTEGER PRIMARY KEY AUTOINCREMENT,"
            " pixel TEXT NOT NULL,"
            " ts INTEGER NOT NULL,"
            " ip TEXT,"
            " ua TEXT)"
        )
        conn.execute("CREATE INDEX IF NOT EXISTS idx_opens_pixel ON opens (pixel)")
    _purge()


def _purge() -> None:
    cutoff = int(time.time()) - RETENTION_DAYS * 86400
    with _connect() as conn:
        conn.execute("DELETE FROM opens WHERE ts < ?", (cutoff,))


def _record(pixel: str, ip: str, ua: str) -> None:
    global _inserts
    with _lock, _connect() as conn:
        (hits,) = conn.execute("SELECT COUNT(*) FROM opens WHERE pixel = ?", (pixel,)).fetchone()
        if hits >= MAX_HITS_PER_PIXEL:
            return
        conn.execute(
            "INSERT INTO opens (pixel, ts, ip, ua) VALUES (?, ?, ?, ?)",
            (pixel, int(time.time()), ip[:64], ua[:400]),
        )
        _inserts += 1
    if _inserts % 1000 == 0:
        _purge()


class Handler(BaseHTTPRequestHandler):
    server_version = "aes-pixel"
    sys_version = ""

    def log_message(self, fmt: str, *args) -> None:  # nginx already logs requests
        return

    def _client_ip(self) -> str:
        return (self.headers.get("X-Real-IP") or self.client_address[0] or "").strip()

    def _send(self, status: int, body: bytes, ctype: str, extra: dict | None = None) -> None:
        self.send_response(status)
        self.send_header("Content-Type", ctype)
        self.send_header("Content-Length", str(len(body)))
        self.send_header("Cache-Control", "no-store, no-cache, must-revalidate, max-age=0")
        self.send_header("Pragma", "no-cache")
        self.send_header("Expires", "0")
        for k, v in (extra or {}).items():
            self.send_header(k, v)
        self.end_headers()
        self.wfile.write(body)

    def do_GET(self) -> None:
        url = urlparse(self.path)
        m = PIXEL_PATH.match(url.path)
        if m:
            try:
                _record(m.group(1), self._client_ip(), self.headers.get("User-Agent") or "")
            except sqlite3.Error:
                pass
            self._send(200, GIF, "image/gif")
            return
        if url.path.startswith("/p/"):
            # Unknown or malformed ids still get an image so nothing breaks in the mail.
            self._send(200, GIF, "image/gif")
            return
        if url.path == "/healthz":
            self._send(200, b"ok", "text/plain")
            return
        if url.path == "/opens":
            self._opens(url)
            return
        self._send(404, b"not found", "text/plain")

    def _opens(self, url) -> None:
        supplied = self.headers.get("X-AES-Key") or ""
        if not API_KEY or not hmac.compare_digest(supplied, API_KEY):
            self._send(401, b'{"error":"unauthorized"}', "application/json")
            return
        try:
            since = int((parse_qs(url.query).get("since") or ["0"])[0])
        except ValueError:
            since = 0
        with _connect() as conn:
            rows = conn.execute(
                "SELECT id, pixel, ts, ip, ua FROM opens WHERE id > ? ORDER BY id LIMIT ?",
                (since, MAX_ROWS_PER_POLL),
            ).fetchall()
        payload = {
            "client_ip": self._client_ip(),
            "opens": [
                {"id": r[0], "pixel": r[1], "ts": r[2], "ip": r[3], "ua": r[4]} for r in rows
            ],
        }
        self._send(200, json.dumps(payload).encode("utf-8"), "application/json")


def main() -> None:
    if not API_KEY:
        raise SystemExit("AES_PIXEL_KEY is not set")
    _init_db()
    httpd = ThreadingHTTPServer(("127.0.0.1", PORT), Handler)
    httpd.daemon_threads = True
    httpd.serve_forever()


if __name__ == "__main__":
    main()
