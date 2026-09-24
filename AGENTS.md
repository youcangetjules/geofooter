# GeoFooter / AES / GURI — Agent Guide

## What this suite is

**Aliniant Email Scanner (AES)** and related tools for Outlook email security analysis, plus **GURI** (Globally Unique Record Identifier) management.

| Product | Role |
| --------- | ------ |
| **AES (GeoFooter)** | Outlook add-in path: watch inbox → export headers/body → Python scan → inject risk banner + HTML footer into the mail |
| **GURI** | Desktop GUI to browse/create/search unique IDs for emails and documents; used by AES when stamping analysis. Includes **Aura** tab (Incogni-style opt-out / erasure request tracking). |
| **Aura** | Aliniant Universal Removal Application — data-broker removal workspace (launched from the Outlook ribbon; lives as a GURI tab). |
| **AesRibbonHost** | Optional COM/ribbon host for AES UI in Outlook |

Primary user flow: new mail arrives in Outlook → VBA queues a **compact** scan → `aes\geolocate_headers.py` runs out-of-process → VBA embeds a PNG status strip and footer (actions via `aes://`).

## Base technologies

- **Outlook VBA** (`VBA\*.bas`, `*.cls`) — events, queue, header export, footer inject, COM
- **Python 3** — scanners, dialogs, GURI GUI (`PySide6`, `Pillow`, `requests`, `dnspython`, `python-whois`, `ipwhois`, `pycountry`, `pytz`, `pywin32`)
- **WScript** — delayed queue ticks / job nudges (Outlook has no reliable `Application.OnTime` here)
- **SQLite** (default) / optional **PostgreSQL** or MySQL for GURI storage
- **Windows Defender / Norton** — optional AV on exported attachments (full/deep; not compact lite)

## Repo layout (live code)

```text
<install root>\
  aes\                      ← AES Python
    geolocate_headers.py    ← main scan engine (compact | full | deep)
    settings_dialog.py / classify_dialog.py / diagnostics_dialog.py  ← PySide6 dialogs
    action_handler.py       ← handles aes:// footer actions (registered in HKCU)
    secret_store.py, addr_decode.py, time_checks.py, route_checks.py, threat_intel.py, score_history.py
    scanners\               ← attachment / body / link / av / norton / windows_defender
  guri\                     ← core.py (GURIDatabase), gui.py (desktop app), gui_service.py, helpers
  aura\                     ← data-broker catalog, profile, request store, GURI tab
  geofooter\                ← paths.py (install root), version.py, crashlog.py
  VBA\                      ← Outlook modules only: MSCAN*.bas / *.cls, IMPORT.txt, ribbon XML
  assets\                   ← brand, icons, pages, outlook forms, sample data
  datastore\                ← local DBs + GURI lists (recipients, important domains) — gitignored
  docs\                     ← feature guides (not live code)
  scripts\                  ← one-off migrators / launchers / icon builders
  tests\                    ← ad-hoc test scripts
  output\                   ← banners, reports, applied footers (runtime)
  AesRibbonHost\            ← ribbon host project
```

Do **not** keep obsolete trees (`Old\`, `_archive\`) in the published repo, and do not add loose `.py` files at the root or under `VBA\`. New Python goes in the matching package.

**Entry scripts** (`aes\geolocate_headers.py`, `aes\*_dialog.py`, `aes\action_handler.py`, `guri\gui.py`, `guri\gui_service.py`) are launched **by file path** from VBA, the ribbon host or `.bat` files. Each one starts with a short `sys.path` bootstrap that drops its own package folder and inserts the install root, so imports are always absolute (`from guri.core import …`, `import aes.addr_decode as addr_decode`). Keep that bootstrap if you add a new entry script.

**Install root** — set in GURI → Database → Install root (writes `%LOCALAPPDATA%\GeoFooter\install_root.txt`). All suite folders above are relative to it. Override with env `GEOFOOTER_ROOT`. See `geofooter\paths.py` / `VBA\MSCANPaths.bas` (`GetAesScript`, `GetGeolocateScript`, `GetGuriGuiScript`).

## Operating constraints agents must respect

1. **Outlook UI thread** — never run `ProcessEmail` / heavy export synchronously from `ItemAdd` / `NewMailEx`. Queue + delayed VBS tick only.
2. **VBA changes require re-import** into Outlook (`VBA\IMPORT.txt`). Editing `.bas` on disk does nothing until imported.
3. **Keep modules in sync** — e.g. `MSCANModule1` calling `MSCANModSenderRules.QuarantineAllAttachments` needs that module imported too.
4. **Compact vs full/deep** — automatic new-mail scans are compact/lite (lighter hop/AV/export). Full/Deep keep heavy fidelity.
5. **Commit + push after each patch** — when a meaningful patch lands, create a captioned git commit and `git push` to `origin` in the same turn (unless the user says not to for that change). Never commit secrets (DPAPI settings written by `aes\secret_store.py`, API keys, `*.dpapi`, live DB configs, `datastore\`).
6. **Work log** — whenever you make meaningful code or project-doc changes, append an entry to **`WORKLOG.md`** (newest first). Include what changed, why, and any follow-up (e.g. VBA re-import). Skip trivial typos-only edits unless the user asks.
7. **Versioning** — SemVer `MAJOR.MINOR.PATCH` in `VERSION` + `geofooter\version.py` + `CHANGELOG.md` + AesRibbonHost `<Version>`. **~90% of bumps are PATCH**; MINOR for new features; MAJOR only for breaking contracts. Keep versions aligned across those files.
8. **Commit captions** — every commit message **must** open with a short caption stating what was done (imperative, specific). Example: `Fix Create-GURI silent insert failure for long document paths`. Do not use vague messages like `update` or `wip`.
9. **Install root** — never hard-code `C:\GeoFooter`. Use `geofooter.paths` (Python) / `MSCANPaths` (VBA). User sets root in GURI → Database → Install root.

## Quick pointers

- Import order / ThisOutlookSession setup: `VBA\IMPORT.txt`
- Dependencies: `requirements.txt`
- Runtime data: `%LOCALAPPDATA%\GeoFooter\` (logs, jobs, sender history, settings)
- Change history: `WORKLOG.md` · release history: `CHANGELOG.md` · bugs: `docs/bug_tracker.md` · current: `VERSION`
- Agent persona / skills: `SKILLS.md`
