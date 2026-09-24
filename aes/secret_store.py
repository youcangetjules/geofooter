#!/usr/bin/env python3
"""AES secret storage — Windows DPAPI (current user) for API keys.

Secrets are written under %LOCALAPPDATA%\\GeoFooter\\secrets\\ as DPAPI-protected
blobs. Only the same Windows user on this machine can decrypt them. Raw keys are
never written to plaintext JSON settings files.
"""

from __future__ import annotations

import os
import re
from pathlib import Path
from typing import Any, Dict, Optional

SECRET_NAME_ABUSEIPDB = "abuseipdb"


def _secrets_dir() -> Path:
    local = os.environ.get("LOCALAPPDATA", "")
    base = Path(local) / "GeoFooter" / "secrets" if local else Path(r"C:\GeoFooter\secrets")
    base.mkdir(parents=True, exist_ok=True)
    return base


def _blob_path(name: str) -> Path:
    safe = re.sub(r"[^a-zA-Z0-9_.-]+", "_", name.strip().lower())
    return _secrets_dir() / f"{safe}.dpapi"


def _hint_path(name: str) -> Path:
    safe = re.sub(r"[^a-zA-Z0-9_.-]+", "_", name.strip().lower())
    return _secrets_dir() / f"{safe}.hint"


def _protect(plaintext: str) -> bytes:
    import win32crypt  # type: ignore

    blob = win32crypt.CryptProtectData(
        plaintext.encode("utf-8"),
        "AES API key",
        None,
        None,
        None,
        0,
    )
    return bytes(blob)


def _unprotect(blob: bytes) -> str:
    import win32crypt  # type: ignore

    _desc, data = win32crypt.CryptUnprotectData(blob, None, None, None, 0)
    if isinstance(data, bytes):
        return data.decode("utf-8", errors="strict")
    return str(data)


def set_secret(name: str, value: str) -> None:
    """Store or clear a named secret. Empty value deletes the stored secret."""
    path = _blob_path(name)
    hint = _hint_path(name)
    cleaned = (value or "").strip()
    if not cleaned:
        if path.is_file():
            path.unlink()
        if hint.is_file():
            hint.unlink()
        return
    path.write_bytes(_protect(cleaned))
    try:
        os.chmod(path, 0o600)
    except OSError:
        pass
    hint.write_text(_mask_hint(cleaned), encoding="utf-8")
    try:
        os.chmod(hint, 0o600)
    except OSError:
        pass


def get_secret(name: str) -> str:
    """Return the decrypted secret, or '' if missing/unreadable."""
    path = _blob_path(name)
    if not path.is_file():
        return ""
    try:
        return _unprotect(path.read_bytes()).strip()
    except Exception:
        return ""


def secret_configured(name: str) -> bool:
    return _blob_path(name).is_file() and bool(get_secret(name))


def secret_hint(name: str) -> str:
    """Non-secret display hint (e.g. ••••a712)."""
    hint = _hint_path(name)
    if hint.is_file():
        try:
            return hint.read_text(encoding="utf-8").strip()
        except OSError:
            pass
    value = get_secret(name)
    return _mask_hint(value) if value else ""


def _mask_hint(value: str) -> str:
    if not value:
        return ""
    if len(value) <= 4:
        return "****"
    return "****" + value[-4:]


def get_abuseipdb_api_key() -> str:
    """Resolve AbuseIPDB key: env override, then DPAPI store."""
    env = (os.environ.get("AES_ABUSEIPDB_API_KEY") or "").strip()
    if env:
        return env
    return get_secret(SECRET_NAME_ABUSEIPDB)


def set_abuseipdb_api_key(value: str) -> None:
    set_secret(SECRET_NAME_ABUSEIPDB, value)


def abuseipdb_status() -> dict:
    configured = bool(get_abuseipdb_api_key())
    return {
        "configured": configured,
        "hint": secret_hint(SECRET_NAME_ABUSEIPDB) if configured else "",
        "source": (
            "env"
            if (os.environ.get("AES_ABUSEIPDB_API_KEY") or "").strip()
            else ("dpapi" if configured else "none")
        ),
    }


def migrate_plaintext_if_present(name: str = SECRET_NAME_ABUSEIPDB) -> str:
    """If a plaintext drop-file exists, DPAPI-encrypt it and delete the plaintext.

    Looks for %LOCALAPPDATA%\\GeoFooter\\secrets\\<name>.key
    Returns 'migrated' | 'already' | 'missing' | 'error:...'
    """
    plain = _secrets_dir() / f"{name}.key"
    if secret_configured(name):
        if plain.is_file():
            try:
                plain.unlink()
            except OSError:
                pass
        return "already"
    if not plain.is_file():
        return "missing"
    try:
        raw = plain.read_text(encoding="utf-8").strip()
        if not raw:
            plain.unlink(missing_ok=True)
            return "missing"
        set_secret(name, raw)
        plain.unlink(missing_ok=True)
        return "migrated"
    except Exception as exc:  # noqa: BLE001
        return f"error:{exc}"


# ---------------------------------------------------------------------------
# Non-secret API endpoint config (URL is not sensitive; key stays in DPAPI)
# ---------------------------------------------------------------------------

DEFAULT_ABUSEIPDB_BASE_URL = "https://api.abuseipdb.com/api/v2"


def _api_config_path() -> Path:
    local = os.environ.get("LOCALAPPDATA", "")
    base = Path(local) / "GeoFooter" if local else Path(r"C:\GeoFooter")
    base.mkdir(parents=True, exist_ok=True)
    return base / "aes_api.json"


def _normalize_abuseipdb_base_url(url: str) -> str:
    text = (url or "").strip().rstrip("/")
    if not text:
        return DEFAULT_ABUSEIPDB_BASE_URL
    # Accept either .../api/v2 or a full .../check URL and normalize to base.
    if text.lower().endswith("/check"):
        text = text[: -len("/check")].rstrip("/")
    if not re.match(r"^https?://", text, flags=re.IGNORECASE):
        text = "https://" + text
    return text


def get_abuseipdb_base_url() -> str:
    """Configured AbuseIPDB API base URL (no trailing slash), e.g. .../api/v2."""
    env = (os.environ.get("AES_ABUSEIPDB_BASE_URL") or "").strip()
    if env:
        return _normalize_abuseipdb_base_url(env)
    path = _api_config_path()
    if path.is_file():
        try:
            import json

            data = json.loads(path.read_text(encoding="utf-8"))
            return _normalize_abuseipdb_base_url(str(data.get("abuseipdb_base_url") or ""))
        except Exception:
            pass
    return DEFAULT_ABUSEIPDB_BASE_URL


def set_abuseipdb_base_url(url: str) -> str:
    """Persist AbuseIPDB base URL; returns the normalized value stored."""
    import json

    normalized = _normalize_abuseipdb_base_url(url)
    path = _api_config_path()
    payload: dict = {}
    if path.is_file():
        try:
            payload = json.loads(path.read_text(encoding="utf-8"))
            if not isinstance(payload, dict):
                payload = {}
        except Exception:
            payload = {}
    payload["abuseipdb_base_url"] = normalized
    path.write_text(json.dumps(payload, indent=2), encoding="utf-8")
    try:
        os.chmod(path, 0o600)
    except OSError:
        pass
    return normalized


def abuseipdb_check_url() -> str:
    """Full check endpoint URL derived from the configured base."""
    return get_abuseipdb_base_url().rstrip("/") + "/check"


def aes_http_get(
    url: str,
    *,
    headers: Optional[Dict[str, str]] = None,
    params: Optional[Dict[str, Any]] = None,
    timeout: float = 10,
) -> tuple[Any, bool]:
    """HTTP GET that tolerates corporate TLS interception.

    Tries normal certificate verification first; on SSL failure retries once
    with verify=False so AbuseIPDB (and similar) still work behind SSL inspection.
    Returns (response, used_insecure_tls).
    """
    import requests
    from requests import exceptions as req_exc

    try:
        resp = requests.get(
            url,
            headers=headers,
            params=params,
            timeout=timeout,
        )
        return resp, False
    except req_exc.SSLError:
        try:
            import urllib3

            urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)
        except Exception:
            pass
        resp = requests.get(
            url,
            headers=headers,
            params=params,
            timeout=timeout,
            verify=False,
        )
        return resp, True
