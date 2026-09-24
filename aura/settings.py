"""Aura automation settings (per-user JSON under %LOCALAPPDATA%\\GeoFooter)."""

from __future__ import annotations

import json
from dataclasses import asdict, dataclass
from datetime import date, datetime
from pathlib import Path
from typing import Any, Optional

from geofooter.paths import user_data_dir

# No backdating: AES broker mail dated before this is never auto-queued.
DEFAULT_DETECT_SINCE = "2026-09-24"
DEFAULT_FOLLOW_UP_DAYS = 30


def settings_path() -> Path:
    return user_data_dir() / "aura_settings.json"


@dataclass
class AuraSettings:
    detect_since: str = DEFAULT_DETECT_SINCE
    send_account: str = ""
    follow_up_days: int = DEFAULT_FOLLOW_UP_DAYS

    def detect_since_date(self) -> date:
        try:
            return date.fromisoformat(self.detect_since.strip()[:10])
        except Exception:
            return date.fromisoformat(DEFAULT_DETECT_SINCE)

    def is_in_window(self, when: Optional[datetime]) -> bool:
        """True when a message date is on/after detect_since. Unknown dates are rejected."""
        if when is None:
            return False
        if when.tzinfo is not None:
            when = when.astimezone()
        return when.date() >= self.detect_since_date()


def load_settings(path: Optional[Path] = None) -> AuraSettings:
    path = path or settings_path()
    if not path.is_file():
        return AuraSettings()
    try:
        raw: Any = json.loads(path.read_text(encoding="utf-8"))
    except Exception:
        return AuraSettings()
    if not isinstance(raw, dict):
        return AuraSettings()
    try:
        days = int(raw.get("follow_up_days") or DEFAULT_FOLLOW_UP_DAYS)
    except (TypeError, ValueError):
        days = DEFAULT_FOLLOW_UP_DAYS
    return AuraSettings(
        detect_since=str(raw.get("detect_since") or DEFAULT_DETECT_SINCE).strip(),
        send_account=str(raw.get("send_account") or "").strip().lower(),
        follow_up_days=max(1, days),
    )


def save_settings(settings: AuraSettings, path: Optional[Path] = None) -> Path:
    path = path or settings_path()
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(asdict(settings), indent=2), encoding="utf-8")
    return path
