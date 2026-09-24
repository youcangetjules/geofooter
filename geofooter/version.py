#!/usr/bin/env python3
"""GeoFooter / AES / GURI suite version and release metadata.

Single source of truth for About UI, window titles, footers, ribbon host,
and packaging. Bump VERSION + VERSION_STAMP + RELEASE_NOTES together when
shipping a meaningful change; mirror the same version in:
  - VERSION (plain text, install root)
  - CHANGELOG.md
  - AesRibbonHost/AesRibbonHost.csproj <Version>

SemVer (MAJOR.MINOR.PATCH) — almost everything is a PATCH:
  PATCH (~90%)  bugfixes, wording, icons, scoring tweaks, docs
  MINOR         new backward-compatible features (tabs, ribbon buttons, modes)
  MAJOR         rare — breaking Outlook/VBA contracts, schema, or public CLI
"""

from __future__ import annotations

# Semver: MAJOR.MINOR.PATCH
VERSION = "1.2.1"
VERSION_TUPLE = (1, 2, 1)
# Local release / build timestamp for this VERSION (YYYY-MM-DD HH:MM)
VERSION_STAMP = "2026-09-24 18:00"

APP_NAME = "GURI"
APP_FULL_NAME = "GURI - Aliniant Smart Ass Email"
# The "Smart Ass" joke: Smart Ass Email ≈ Aliniant Smart Scraper
APP_SILLY_QUOTE = '"Aliniant Smart Scraper"'
APP_ORG = "Aliniant"
SUITE_NAME = "GeoFooter"
SUITE_PRODUCTS = ("AES", "GURI", "Aura", "AesRibbonHost")
COPYRIGHT_YEAR = "2026"

# Short release notes for the About tab (newest first).
# Each entry: (version, timestamp, notes)
RELEASE_NOTES = (
    (
        "1.2.1",
        "2026-09-24 18:00",
        "Suite reorganised into aes/, guri/, aura/, geofooter/ packages; VBA\\ "
        "holds only Outlook modules; GURI lists live in datastore/.",
    ),
    (
        "1.2.0",
        "2026-09-24 15:25",
        "Aura (data-broker removal) ribbon + GURI tab; threat-intel RiskTable "
        "consensus scoring; Create-GURI DB persist (avg_risk TEXT); clickable "
        "footer metrics; IP allocation decode; suite versioning + public GitHub.",
    ),
    (
        "1.1.0",
        "2026-08-10 14:16",
        "User rules with compile + local enforcement; Locations file scrape; "
        "identity / not-for-me deadlines; timeline importance colours; "
        "scrape lookback months; About + version.py.",
    ),
    (
        "1.0.0",
        "2026-01-02 09:00",
        "GURI database viewer, Outlook scrape, deadlines, learning loop, Ollama.",
    ),
)


def version_string() -> str:
    """Human-readable version with timestamp, e.g. 1.2.0 (2026-09-24 15:25)."""
    stamp = (VERSION_STAMP or "").strip()
    if stamp:
        return f"{VERSION} ({stamp})"
    return VERSION


def suite_banner() -> str:
    """One-line product + version for logs / footers."""
    return f"{SUITE_NAME} / {'+'.join(SUITE_PRODUCTS)} v{version_string()}"


def about_summary() -> str:
    return (
        f"{APP_FULL_NAME}  {APP_SILLY_QUOTE}\n"
        f"Suite {SUITE_NAME}  v{version_string()}\n"
        f"Products: {', '.join(SUITE_PRODUCTS)}\n"
        f"© {COPYRIGHT_YEAR} {APP_ORG}"
    )
