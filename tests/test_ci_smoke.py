#!/usr/bin/env python3
"""Linux/CI smoke tests for the package layout (no Outlook / pywin32)."""

from __future__ import annotations

import compileall
import importlib
from pathlib import Path

import pytest

ROOT = Path(__file__).resolve().parents[1]


def test_packages_compile() -> None:
    for name in ("aes", "guri", "aura", "geofooter"):
        ok = compileall.compile_dir(str(ROOT / name), quiet=1)
        assert ok, f"compile failed under {name}/"


@pytest.mark.parametrize(
    "mod",
    [
        "geofooter",
        "geofooter.paths",
        "geofooter.version",
        "geofooter.crashlog",
        "aes",
        "aes.addr_decode",
        "aes.score_history",
        "aes.secret_store",
        "aes.scanners",
        "aes.scanners.link",
        "guri",
        "guri.core",
        "aura",
        "aura.detect",
        "aura.catalog",
        "aura.settings",
        "aura.automation",
    ],
)
def test_import_module(mod: str) -> None:
    importlib.import_module(mod)


def test_version_aligned_with_version_file() -> None:
    from geofooter.version import VERSION

    text = (ROOT / "VERSION").read_text(encoding="utf-8").strip()
    assert text == VERSION


def test_vba_modules_are_ascii() -> None:
    # The VBA editor imports .bas/.cls as ANSI; UTF-8 punctuation becomes mojibake.
    bad = []
    for path in sorted((ROOT / "VBA").glob("*.[bc][al][ss]")):
        for lineno, line in enumerate(path.read_bytes().splitlines(), 1):
            if any(b > 0x7F for b in line):
                bad.append(f"{path.name}:{lineno}")
    assert not bad, "non-ASCII in VBA modules: " + ", ".join(bad[:20])


def test_install_root_markers() -> None:
    assert (ROOT / "aes" / "geolocate_headers.py").is_file()
    assert (ROOT / "guri" / "gui.py").is_file()
    assert (ROOT / "environment.yml").is_file()
