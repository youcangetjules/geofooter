#!/usr/bin/env python3
"""Footer one-liners (witticisms) and Poe suggestions.

The built-in banks are the defaults. A saved file under
%LOCALAPPDATA%\\GeoFooter\\aes_witticisms.json replaces them per risk band.
The Poe API key itself stays in the DPAPI secret store.
"""

from __future__ import annotations

import json
import os
import random
import urllib.error
import urllib.request
from pathlib import Path
from typing import Dict, List

LEVELS = ("LOW", "RAISED", "HIGH", "CRITICAL")
LEVEL_LABELS = {
    "LOW": "Low",
    "RAISED": "Raised",
    "HIGH": "High",
    "CRITICAL": "Critical (score over 75)",
}
DEFAULT_MODEL = "Claude-Sonnet-4.6"
POE_URL = "https://api.poe.com/v1/chat/completions"

DEFAULT_LINES: Dict[str, List[str]] = {
    "LOW": [
        "Relax. This one could not organise a phishing trip.",
        "Safer than your browser history. Marginally.",
        "We poked it with a stick. It apologised.",
        "Green light. You may unclench.",
        "It brought ID and a sensible jumper.",
        "Nothing suspicious, which is honestly a bit suspicious. Still fine.",
        "Cleared. The scanner is going back to its nap.",
        "You can read this one without adult supervision.",
    ],
    "RAISED": [
        "Amber. Read it like it owes you money.",
        "Not a villain. Do not give it your passwords anyway.",
        "We did not hate it. We did not trust it either.",
        "Fine to open. Terrible idea to click the shiny bit.",
        "Raised eyebrow fitted at no extra charge.",
        "Probably legit. 'Probably' is doing a lot of work.",
        "Hover before you click. Future you says thanks.",
        "Interesting email. Keep your wallet in your pocket.",
    ],
    "HIGH": [
        "This one has 'trust me' energy. Do not.",
        "High. The links are decorative. Leave them that way.",
        "If it wants a login, it can want it from someone else.",
        "We would not let this one borrow a pen.",
        "Smile, nod, and click absolutely nothing.",
        "The scanner put it in the naughty corner.",
        "Pretty. Pushy. Put the mouse down.",
        "Charming, and almost certainly up to something.",
    ],
    "CRITICAL": [
        "So risky we took its HTML away.",
        "Plain text, because the fancy version was up to something.",
        "This one does not get buttons. It knows what it did.",
        "We sent the layout to its room.",
        "Too spicy for formatting. It can sit in plain text and think.",
    ],
}


def witticisms_path() -> Path:
    local = os.environ.get("LOCALAPPDATA", "")
    base = Path(local) / "GeoFooter" if local else Path.home() / "GeoFooter"
    base.mkdir(parents=True, exist_ok=True)
    return base / "aes_witticisms.json"


def _clean_line(text: str) -> str:
    line = " ".join(str(text or "").split())
    if len(line) >= 2 and line[0] == line[-1] and line[0] in "\"'":
        line = line[1:-1].strip()
    return line[:180]


def normalize_banks(raw: object) -> Dict[str, List[str]]:
    """Return one list per risk band. Missing bands keep the built-in lines."""
    source = raw if isinstance(raw, dict) else {}
    nested = source.get("lines") if isinstance(source.get("lines"), dict) else source
    banks: Dict[str, List[str]] = {}
    for level in LEVELS:
        items = nested.get(level) if isinstance(nested, dict) else None
        if not isinstance(items, list):
            banks[level] = list(DEFAULT_LINES[level])
            continue
        seen = set()
        cleaned: List[str] = []
        for item in items:
            line = _clean_line(str(item or ""))
            key = line.lower()
            if line and key not in seen:
                seen.add(key)
                cleaned.append(line)
        banks[level] = cleaned
    return banks


def load_witticisms() -> Dict[str, object]:
    """Saved banks plus the Poe model name. Missing file means the defaults."""
    model = DEFAULT_MODEL
    raw: object = None
    path = witticisms_path()
    if path.is_file():
        try:
            raw = json.loads(path.read_text(encoding="utf-8"))
        except (OSError, json.JSONDecodeError):
            raw = None
    if isinstance(raw, dict):
        chosen = str(raw.get("model") or "").strip()
        if chosen:
            model = chosen[:80]
    return {"model": model, "lines": normalize_banks(raw)}


def save_witticisms(lines: Dict[str, List[str]], model: str = "") -> None:
    payload = {
        "model": (model or DEFAULT_MODEL).strip()[:80] or DEFAULT_MODEL,
        "lines": normalize_banks({"lines": lines}),
    }
    path = witticisms_path()
    tmp = path.with_suffix(".json.tmp")
    tmp.write_text(json.dumps(payload, indent=2, ensure_ascii=False), encoding="utf-8")
    os.replace(tmp, path)


def pick_witticism(level: str) -> str:
    """One line for this risk band. An empty saved band falls back to the defaults."""
    band = str(level or "").upper()
    if band not in DEFAULT_LINES:
        band = "HIGH"
    banks = load_witticisms()["lines"]
    choices = banks.get(band) if isinstance(banks, dict) else None
    if not choices:
        choices = DEFAULT_LINES[band]
    return random.choice(list(choices))


def suggest_witticism(
    level: str,
    existing: List[str],
    api_key: str,
    model: str = "",
) -> str:
    """Ask Poe for one new line. Raises RuntimeError with a short message."""
    key = (api_key or "").strip()
    if not key:
        raise RuntimeError("Add a Poe API key first.")
    band = str(level or "LOW").upper()
    if band not in LEVEL_LABELS:
        band = "LOW"
    samples = [line for line in existing if line][:12]
    sample_block = "\n".join(f"- {line}" for line in samples) or "- (none yet)"
    prompt = (
        "Write one cheeky, funny one-line comment for the corner of an email "
        "security scan footer. British, dry, and a bit rude. No hashtags, no "
        "emoji, no quotation marks, under 90 characters.\n"
        f"Risk band: {LEVEL_LABELS[band]}.\n"
        "Do not repeat any of these:\n"
        f"{sample_block}\n"
        "Reply with only the line."
    )
    body = json.dumps(
        {
            "model": (model or DEFAULT_MODEL).strip() or DEFAULT_MODEL,
            "messages": [{"role": "user", "content": prompt}],
            "max_tokens": 80,
            "temperature": 0.9,
        }
    ).encode("utf-8")
    request = urllib.request.Request(
        POE_URL,
        data=body,
        headers={
            "Authorization": f"Bearer {key}",
            "Content-Type": "application/json",
            "Accept": "application/json",
        },
        method="POST",
    )
    try:
        with urllib.request.urlopen(request, timeout=25) as response:
            payload = json.loads(response.read().decode("utf-8"))
    except urllib.error.HTTPError as exc:
        if exc.code == 401:
            raise RuntimeError("Poe rejected that API key.") from None
        if exc.code == 402:
            raise RuntimeError("The Poe account has no credits left.") from None
        if exc.code == 429:
            raise RuntimeError("Poe is rate-limiting requests. Try again in a minute.") from None
        raise RuntimeError(f"Poe returned HTTP {exc.code}.") from None
    except (urllib.error.URLError, TimeoutError, json.JSONDecodeError) as exc:
        raise RuntimeError(f"Could not reach Poe ({type(exc).__name__}).") from None
    try:
        text = payload["choices"][0]["message"]["content"]
    except (KeyError, IndexError, TypeError) as exc:
        raise RuntimeError("Poe replied in an unexpected shape.") from exc
    line = _clean_line(str(text or ""))
    if not line:
        raise RuntimeError("Poe returned an empty line.")
    return line
