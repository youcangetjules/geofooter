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
VERSION = "1.3.44"
VERSION_TUPLE = (1, 3, 44)
# Local release / build timestamp for this VERSION (YYYY-MM-DD HH:MM)
VERSION_STAMP = "2026-09-25 12:30"

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
        "1.3.44",
        "2026-09-25 12:30",
        "The footer A spans all three rows of the scan strip.",
    ),
    (
        "1.3.43",
        "2026-09-25 12:25",
        "Links, OK, and NOK each sit in their own footer cell so they share the top-row baseline.",
    ),
    (
        "1.3.42",
        "2026-09-25 12:20",
        "Footer lines are centred on the full strip. Links counts share one baseline, and the copyright uses the © symbol.",
    ),
    (
        "1.3.41",
        "2026-09-25 12:05",
        "The second footer line renders the relay note instead of raw HTML, and the Quick Actions row shows the copyright on the right.",
    ),
    (
        "1.3.40",
        "2026-09-25 11:58",
        "The footer A attaches to the message you are reading, not only to a popped-out window.",
    ),
    (
        "1.3.39",
        "2026-09-25 11:55",
        "Footer asides are longer and funnier, and they still stay on one line.",
    ),
    (
        "1.3.38",
        "2026-09-25 11:55",
        "Footer asides stay on one line. The top-row counts share one baseline. The orange A is the logo image again.",
    ),
    (
        "1.3.37",
        "2026-09-25 11:50",
        "Footer buttons are shorter and their fill stays inside the border. The orange A is drawn in the footer. A random cheeky line sits on the top right.",
    ),
    (
        "1.3.36",
        "2026-09-25 11:40",
        "A scan that dies or overruns writes a failure marker, so Outlook does not stay on AES PROC. Repeated failures pause the queue.",
    ),
    (
        "1.3.35",
        "2026-09-25 10:30",
        "Actions to take opens newest first. Column headers sort, and a second click reverses. Timeline days read 24th and 25th, paler.",
    ),
    (
        "1.3.34",
        "2026-09-25 00:20",
        "Footer top-row labels and dots share one vertical centre.",
    ),
    (
        "1.3.33",
        "2026-09-24 22:55",
        "The footer A is embedded in the message so Outlook can show it.",
    ),
    (
        "1.3.32",
        "2026-09-24 22:50",
        "The orange A sits on the left of the footer. A faint line separates each row.",
    ),
    (
        "1.3.31",
        "2026-09-24 22:45",
        "A message already converted to text stays text. Button clicks no longer write HTML and a scan at the same time.",
    ),
    (
        "1.3.30",
        "2026-09-24 22:40",
        "NT stays neutral: first click drops trust and blocks beacons. Second click is Full No Trust and converts that mail to text. Footer buttons replace the footer.",
    ),
    (
        "1.3.29",
        "2026-09-24 22:30",
        "Beacon Blocking settings are the default. A footer button on one mail overrides them for that sender.",
    ),
    (
        "1.3.28",
        "2026-09-24 22:20",
        "When beacon blocking covers the sender, detected beacons count as blocked.",
    ),
    (
        "1.3.27",
        "2026-09-24 22:18",
        "The whole HRE label is the link, not only the count.",
    ),
    (
        "1.3.26",
        "2026-09-24 22:15",
        "HRE opens the risk breakdown. Footer row padding is half of what it was.",
    ),
    (
        "1.3.25",
        "2026-09-24 22:12",
        "Scans no longer restore HTML automatically, so a mitigated mail stops flipping back to HTML.",
    ),
    (
        "1.3.24",
        "2026-09-24 22:10",
        "HRE sits to the right of Links: green at 0, yellow at 1–3, red above that.",
    ),
    (
        "1.3.23",
        "2026-09-24 22:08",
        "Footer rows are evenly spaced, with the same hairline between row 1 and 2 as between 2 and 3.",
    ),
    (
        "1.3.22",
        "2026-09-24 22:05",
        "Attachment counts use the same green or red as OK and NOK.",
    ),
    (
        "1.3.21",
        "2026-09-24 22:00",
        "BB is red when global beacon blocking covers the sender, not only the per-sender list.",
    ),
    (
        "1.3.20",
        "2026-09-24 21:55",
        "Risk label is light green, amber, or red by band. Scores above 75 become text with a restore-HTML link.",
    ),
    (
        "1.3.19",
        "2026-09-24 21:50",
        "Beacon 0 Blocked and link OK counts are green; blocked beacons and link NOK counts are red.",
    ),
    (
        "1.3.18",
        "2026-09-24 21:45",
        "A 12px white spacer row sits above the footer so Outlook cannot attach it to the message.",
    ),
    (
        "1.3.17",
        "2026-09-24 21:42",
        "Every item on the footer top row is 12px bold.",
    ),
    (
        "1.3.16",
        "2026-09-24 21:40",
        "Blocked chips are red, trusted is a lighter solid green, and the gap above the footer is email white.",
    ),
    (
        "1.3.15",
        "2026-09-24 21:35",
        "Footer uses Segoe UI, 5px cells beside each green dot, a darker slate, a smaller second line, and taller buttons.",
    ),
    (
        "1.3.14",
        "2026-09-24 21:30",
        "Footer uses a grey slate, green dots between facts, a larger second line, tighter buttons, and 5px above the strip.",
    ),
    (
        "1.3.13",
        "2026-09-24 21:25",
        "Logging settings include a max log size in MiB. 0 is unlimited; over the cap, the oldest lines are removed.",
    ),
    (
        "1.3.12",
        "2026-09-24 21:20",
        "Footer strip redesigned: one dark slate background, risk dot, hairline dividers, white pill actions.",
    ),
    (
        "1.3.11",
        "2026-09-24 21:12",
        "The selected tab is a shade darker than the window background.",
    ),
    (
        "1.3.10",
        "2026-09-24 21:10",
        "Email timeline shows the day of the month at 20% / 72pt, and a solid bar between 23:00 and 00:00.",
    ),
    (
        "1.3.9",
        "2026-09-24 21:05",
        "Quick Actions label uses dark or white text so it contrasts with the scan strip.",
    ),
    (
        "1.3.8",
        "2026-09-24 21:00",
        "Quick Action chips sit together with 20px between them.",
    ),
    (
        "1.3.7",
        "2026-09-24 20:10",
        "Centred top-bar clock. PostgreSQL and Database live on the Settings tab.",
    ),
    (
        "1.3.6",
        "2026-09-24 20:05",
        "One Refresh at the top right, a Push/Pull pill beside it, and a Settings tab after Ollama.",
    ),
    (
        "1.3.5",
        "2026-09-24 20:00",
        "Tray icon registers as AGS (Aliniant Geosense Suite) and is turned on in Windows taskbar settings.",
    ),
    (
        "1.3.4",
        "2026-09-24 19:35",
        "Welcome Refresh reloads Outlook. The click was ignored because Qt passed a bool the handler rejected.",
    ),
    (
        "1.3.3",
        "2026-09-24 19:30",
        "Quick Action chips are wider and spaced evenly, with Quick Actions: on the left of that centred row.",
    ),
    (
        "1.3.2",
        "2026-09-24 19:25",
        "Footer Quick Actions row (SL, BA, BB, TS/ST, NT). Reply chains keep one scan result at the bottom.",
    ),
    (
        "1.3.1",
        "2026-09-24 19:45",
        "Master installer scripts/Install_GeoFooter.bat. It states that PostgreSQL is required.",
    ),
    (
        "1.3.0",
        "2026-09-24 19:30",
        "Aura automation: removal emails for the whole catalogue drafted in Outlook for "
        "review (never auto-sent); AES auto-queues broker mail dated 2026-09-24 onwards; "
        "sent drafts tracked with 30-day follow-up.",
    ),
    (
        "1.2.10",
        "2026-09-24 18:45",
        "CI: tests/conftest.py puts the install root on sys.path for pytest.",
    ),
    (
        "1.2.9",
        "2026-09-24 18:40",
        "Unblock Conda CI: fix missing import os in scripts/fix_duplicate_guris.py.",
    ),
    (
        "1.2.8",
        "2026-09-24 18:30",
        "MSCANSelfUpdate: wake VBE with Alt+F11; manual import fallback when Application.VBE stays Nothing despite AccessVBOM=1.",
    ),
    (
        "1.2.7",
        "2026-09-24 18:25",
        "Force CRLF on all VBA modules; .gitattributes locks *.bas/*.cls.",
    ),
    (
        "1.2.6",
        "2026-09-24 18:20",
        "In-Outlook VBA updater (Alt+F8 UpdateAesFromDisk) for machines where external import can't reach VBE.",
    ),
    (
        "1.2.5",
        "2026-09-24 17:40",
        "VBA modules are pure ASCII (no more garbled dashes in Outlook dialogs); CI guard.",
    ),
    (
        "1.2.4",
        "2026-09-24 17:30",
        "VBA import script compiles and saves the project so Outlook restarts keep the new code.",
    ),
    (
        "1.2.3",
        "2026-09-24 17:10",
        "Conda environment.yml + CI; SemVer bump on every patch; GURI/Aura cold-start "
        "and AES toolbar recovery after VBA re-import.",
    ),
    (
        "1.2.2",
        "2026-09-24 16:35",
        "Automated Outlook VBA sync (scripts/Import_VBA_to_Outlook); docs bug tracker.",
    ),
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
