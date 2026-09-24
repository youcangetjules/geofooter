"""Identity profile for removal letter templates (local JSON)."""

from __future__ import annotations

import json
import os
from dataclasses import asdict, dataclass, field
from pathlib import Path
from typing import Any, Dict, List


def data_dir() -> Path:
    local = os.environ.get("LOCALAPPDATA", "")
    return Path(local) / "GeoFooter" if local else Path(r"C:\GeoFooter")


def profile_path() -> Path:
    return data_dir() / "broker_identity_profile.json"


@dataclass
class IdentityProfile:
    full_name: str = ""
    aliases: List[str] = field(default_factory=list)
    emails: List[str] = field(default_factory=list)
    phones: List[str] = field(default_factory=list)
    addresses: List[str] = field(default_factory=list)
    date_of_birth: str = ""  # optional YYYY-MM-DD
    notes: str = ""

    def to_dict(self) -> Dict[str, Any]:
        return asdict(self)

    @classmethod
    def from_dict(cls, raw: Any) -> "IdentityProfile":
        if not isinstance(raw, dict):
            return cls()

        def _list(key: str) -> List[str]:
            val = raw.get(key)
            if isinstance(val, list):
                return [str(x).strip() for x in val if str(x).strip()]
            if isinstance(val, str) and val.strip():
                return [line.strip() for line in val.splitlines() if line.strip()]
            return []

        return cls(
            full_name=str(raw.get("full_name") or "").strip(),
            aliases=_list("aliases"),
            emails=_list("emails"),
            phones=_list("phones"),
            addresses=_list("addresses"),
            date_of_birth=str(raw.get("date_of_birth") or "").strip(),
            notes=str(raw.get("notes") or "").strip(),
        )


def load_profile() -> IdentityProfile:
    path = profile_path()
    if not path.is_file():
        return IdentityProfile()
    try:
        return IdentityProfile.from_dict(json.loads(path.read_text(encoding="utf-8")))
    except Exception:
        return IdentityProfile()


def save_profile(profile: IdentityProfile) -> Path:
    path = profile_path()
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(
        json.dumps(profile.to_dict(), indent=2, ensure_ascii=False),
        encoding="utf-8",
    )
    return path
