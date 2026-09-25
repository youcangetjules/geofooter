# Changelog

All notable changes to the **GeoFooter / AES / GURI / Aura** suite are documented here.

Format follows [Keep a Changelog](https://keepachangelog.com/en/1.1.0/).
Versioning follows [Semantic Versioning](https://semver.org/spec/v2.0.0.html):
**MAJOR.MINOR.PATCH** — roughly **90% of releases are patches**.

| Bump | When |
|------|------|
| **PATCH** | Default — fixes, wording, scoring tweaks, icons, docs |
| **MINOR** | New backward-compatible features (ribbon buttons, tabs, modes) |
| **MAJOR** | Rare — breaking Outlook/VBA contracts, DB schema, or public CLI |

The canonical version string lives in `geofooter/version.py` and the plain-text `VERSION` file.
Bump those together with this changelog and `AesRibbonHost/AesRibbonHost.csproj` `<Version>`.

Every git commit must start with a short **caption** (one line) stating what was done,
e.g. `Fix Create-GURI silent insert failure for long document paths`.

## [Unreleased]

## [1.3.42] — 2026-09-25

### Fixed

- The Links count, including NOK, sits on the same baseline as the rest of the top row.
- The scan line, the sender line, and the Quick Actions are centred on the full footer. The logo and the side text no longer shift that centre.
- The bottom-right notice uses the © symbol.

## [1.3.41] — 2026-09-25

### Fixed

- The second footer line was printing the outbound-relay note as raw HTML. That note now renders, and the line is centred.

### Added

- The Quick Actions row shows “Copyright 2026 Aliniant Labs” on the far right.

## [1.3.40] — 2026-09-25

### Fixed

- The footer A was only sought in a separate message window. It now attaches to the mail selected in the reading pane.

## [1.3.39] — 2026-09-25

### Changed

- The cheeky footer lines are longer and funnier, and they still stay on one line.

## [1.3.38] — 2026-09-25

### Changed

- The cheeky footer line stays on one line.
- Top-row counts, including NOK, share the same baseline.
- The orange A is the logo image again, attached to the message after the footer is applied.

## [1.3.37] — 2026-09-25

### Changed

- Quick Action buttons are 16px tall. The fill and the border are the same box, so the colour no longer spills past the line.
- The orange A is drawn in the footer itself, so Outlook does not need a separate image.
- The top-right of the footer shows a random cheeky line for the risk band.

## [1.3.36] — 2026-09-25

### Fixed

- Compact scans crashed before writing a footer, so Outlook held AES PROC for five minutes per mail.
- A scan that overruns its budget, or dies before writing a footer, leaves a failure marker. DNS and WHOIS lookups are time-limited. Three failures in a row pause the queue for two minutes.

## [1.3.35] — 2026-09-25

### Changed

- Actions to take opens in time order, newest first. A column header sorts by that column; a second click reverses it.
- Timeline day numbers are paler and carry the day suffix (24th, 25th).

## [1.3.34] — 2026-09-25

### Changed

- Footer top-row labels, counts, and the green dots share one vertical centre.

## [1.3.33] — 2026-09-24

### Fixed

- The footer A was a blocked file link, which Outlook drew as an empty bar. It is now embedded in the message.

## [1.3.32] — 2026-09-24

### Changed

- The orange A is on the left of the footer, with the white background removed.
- A faint line separates each footer row.

## [1.3.31] — 2026-09-24

### Fixed

- A mail already converted to text stays text. A footer button no longer writes HTML into that message and starts a scan at the same time, which was flipping it back and forth.

## [1.3.30] — 2026-09-24

### Changed

- NT has three states. The chip stays the neutral colour. The first click removes Trust and blocks beacons. The second click is Full No Trust: attachments are blocked and that mail becomes plain text, with a restore link that is not opened until you click it.
- BA, BB, TS, and the first NT click replace the footer on the open message.

## [1.3.29] — 2026-09-24

### Changed

- Beacon Blocking in AES Settings is the default for every sender. The BB button on a mail overrides that default for that sender, including turning blocking off when the global rule is on.

## [1.3.28] — 2026-09-24

### Fixed

- With beacon blocking on for the sender, a detected beacon is counted as blocked. The line no longer says “1 detected 0 Blocked” when those beacons are blocked.

## [1.3.27] — 2026-09-24

### Changed

- The whole **HRE** label links to the risk breakdown. The count keeps its green, yellow, or red colour.

## [1.3.26] — 2026-09-24

### Changed

- **HRE** is a link to the risk-score breakdown, which lists the elements behind the count.
- Vertical padding on each footer row is 5px, half of the previous 10px.

## [1.3.25] — 2026-09-24

### Fixed

- A mail converted to text was being turned back into HTML on the next scan, then converted to text again. Scans no longer restore that HTML. Use the restore link in the text to put the HTML back.

## [1.3.24] — 2026-09-24

### Added

- **HRE** (high-risk elements) is shown to the right of Links. The count is green at 0, yellow at 1–3, and red above that.

## [1.3.23] — 2026-09-24

### Changed

- The three footer rows use the same vertical padding. The hairline between the detail line and Quick Actions is repeated between the title line and the detail line, at the same width.

## [1.3.22] — 2026-09-24

### Changed

- Attachment counts sit inside the status colour: green with **OK**, red with **NOK**.

## [1.3.21] — 2026-09-24

### Fixed

- **BB** stays white unless the sender is on the per-sender beacon list. With “block all scanned mail” on, and the sender not whitelisted, **BB** is now red.

## [1.3.20] — 2026-09-24

### Changed

- The risk words and score are colour-coded: **LOW** light green, **RAISED** (26–50) amber, **HIGH** (51–75) red.
- A score above 75 still strips attachments and converts the message to text. That text includes a link to restore the original HTML.

## [1.3.19] — 2026-09-24

### Changed

- Attachments still show a green **OK** when there are none or every file is fine.
- Beacons show **0 Blocked** in that same green. If any beacon is blocked, that count is **Blocked** in red.
- Links show the clear count as **OK** in green. Dangerous links add **NOK** in red beside it.

## [1.3.18] — 2026-09-24

### Fixed

- The footer no longer sits flush against the message. Outlook was dropping the div padding, so a 12px white spacer row is now part of the footer HTML.

## [1.3.17] — 2026-09-24

### Changed

- Attachments, beacons, and links on the footer top row use the same 12px bold type as the rest of that line.

## [1.3.16] — 2026-09-24

### Changed

- Quick Action chips show status: blocked attachments or beacons are red, a trusted sender is a lighter solid green. The 5px gap above the footer is the email background (white), not the slate of the strip.

## [1.3.15] — 2026-09-24

### Changed

- The footer is set in Segoe UI. Each green dot has a 5px-wide cell on either side, which Outlook keeps. The strip is 10% darker (`#39444d`). The second line is 1pt smaller. The action buttons are 4px taller.

## [1.3.14] — 2026-09-24

### Changed

- The footer background is a grey slate (`#3f4c55`, the previous slate at 80% over white). Facts are separated by a small green dot with 5px on each side. The second line is 2pt larger. The button row has 5px less space above and below the chips. There is 5px between the bottom of the message and the top of the footer.

## [1.3.13] — 2026-09-24

### Added

- AES Settings → Logging has **Max log size** in MiB. The default is 0, which means unlimited. When `VBA_Log.txt` grows past the limit, the oldest lines are removed.

## [1.3.12] — 2026-09-24

### Changed

- The AES footer strip is redesigned: one flat dark-slate background (`#0f1f2a`), white heading, muted detail line, and hairline `|` dividers instead of the green-to-red gradient. Risk shows as a small coloured dot beside the label.
- Quick Action chips are white pills with slate letters so they lift off the strip. Blocked attachments/beacons fill red; a trusted sender fills green.

## [1.3.11] — 2026-09-24

### Changed

- The selected tab is a shade darker than the window background.

## [1.3.10] — 2026-09-24

### Changed

- The email timeline draws the day of the month behind the hours at 20% opacity and 72pt.
- The grid line between 23:00 and 00:00 is a solid dark bar.

## [1.3.9] — 2026-09-24

### Changed

- **Quick Actions:** uses near-black or white, whichever contrasts with the scan-strip colour. White on the low-risk green was washing out.

## [1.3.8] — 2026-09-24

### Changed

- Footer Quick Action chips (SL, BA, BB, TS, NT) sit together in the centre of the strip, with 20px between each chip.

## [1.3.7] — 2026-09-24

### Changed

- The top bar shows a centred live clock, for example **Now: 24 September 2026; 20:08:56**.
- **PostgreSQL** and **Database…** moved from the top bar into the Settings tab.

## [1.3.6] — 2026-09-24

### Changed

- Welcome no longer has its own Refresh buttons. One **Refresh** sits at the top right. The pill beside it is **Push** (AES tells GURI as soon as a scan finishes) or **Pull** (GURI polls Outlook on a timer).
- **Settings** is the rightmost tab, next to Ollama.

## [1.3.5] — 2026-09-24

### Fixed

- The suite tray icon was registered as GURI under Python, so it did not appear as AGS in Settings > Personalisation > Taskbar > Other system tray icons. It now registers as **AGS (Aliniant Geosense Suite)** and turns that entry On.

## [1.3.4] — 2026-09-24

### Fixed

- Welcome Refresh (timeline, matching emails, actions, statistics) ignored the click: Qt passed a bool into handlers that did not accept it, so the dashboard never reloaded. Refresh now redraws and pulls a fresh Outlook scrape.

## [1.3.3] — 2026-09-24

### Changed

- Quick Action chips are about three times as wide. "Quick Actions:" sits on the left of that row; the label and the chips share the footer width in equal slots so the row stays centred.

## [1.3.2] — 2026-09-24

### Added

- Footer third row, **Quick Actions**: Show Links `<SL>`, Block Attachments `<BA>`, Block Beacons `<BB>`, Trust Sender `<TS>`, Mark Not Trusted `<NT>`. Chips are white on green. `<BA>` / `<BB>` turn red when that block is already on. A trusted sender shows `<ST>` in green instead of `<TS>`.

### Fixed

- A reply chain no longer keeps every earlier AES scan. Previous results are removed and one scan result is placed at the bottom of the message.

## [1.3.1] — 2026-09-24

### Added

- Master installer `scripts/Install_GeoFooter.bat` (runs `scripts/Install-GeoFooter.ps1`) installs both the AES email tool (venv, `aes://`, Outlook VBA, ribbon) and the GURI GUI (same venv, desktop shortcut, `Launch_GURI_GUI.bat` uses `.venv`). It tells you PostgreSQL is required and copies `guri_postgres_config.example.json` when the real config is missing.

## [1.3.0] — 2026-09-24

### Added

- Aura automation (`aura/automation.py`): creates an Outlook draft removal email for every catalogue broker with an opt-out address, plus AES-detected brokers. Drafts are saved to Drafts (category "Aura") and are **never sent automatically**.
- Aura tab **Review** sub-tab: Open draft, Regenerate, Skip, Prepare all now. Aura runs on open and every 5 minutes, and marks requests as sent (30-day follow-up) when the email shows up in Sent Items.
- "Send from" Outlook account in the Aura identity form (`aura/settings.py`, `aura_settings.json`).
- AES scans auto-queue detected broker mail to Aura, but only mail dated on or after the Aura start date (default 2026-09-24).
- `draft_entry_id` column on Aura requests (auto-migrated).
- `tests/test_aura_automation.py` (fake draft backend, no Outlook needed); runs in Conda CI.

### Changed

- Web-form-only brokers are tracked as "Manual web form (not automated)" and are not submitted.

## [1.2.10] — 2026-09-24

### Fixed

- Conda CI smoke tests failed with `ModuleNotFoundError` for `geofooter` / `aes` / `guri` / `aura`: added `tests/conftest.py` to put the install root on `sys.path`.

## [1.2.9] — 2026-09-24

### Fixed

- Conda CI flake8 F821: `scripts/fix_duplicate_guris.py` used `os` without importing it (blocked every push since the workflow started linting `scripts/`). Path now uses `geofooter.paths.datastore_path` instead of a hard-coded `C:\GeoFooter`.

## [1.2.8] — 2026-09-24

### Fixed

- `MSCANSelfUpdate`: no longer tells you to set AccessVBOM when it is already 1. Nudges Alt+F11 and retries; if `Application.VBE` is still Nothing, opens the VBA folder and pastes a remove/import checklist (File > Import does not need VBOM access).

## [1.2.7] — 2026-09-24

### Fixed

- VBA modules MSCANModule1.bas and MSCANToolbar.bas were LF-only; Outlook File > Import needs CRLF. Added `.gitattributes` (`*.bas` / `*.cls` eol=crlf).

## [1.2.6] — 2026-09-24

### Added

- `VBA/MSCANSelfUpdate.bas`: `UpdateAesFromDisk` re-imports all `VBA\MSCAN*` modules (list read from `VBA\IMPORT.txt`), then compiles and saves, from inside Outlook. It's needed because Outlook returns no `Application.VBE` to external scripts on some machines, even with AccessVBOM=1.

### Changed

- `Import_VBA_to_Outlook.ps1` also imports `MSCANSelfUpdate.bas`, and its failure message points at the in-Outlook updater.

## [1.2.5] — 2026-09-24

### Fixed

- Garbled punctuation (`â€”`) in VBA dialog titles and messages. All `VBA\*.bas` / `*.cls` are now pure ASCII (97 replacements), because the VBA editor imports them as ANSI.

### Added

- CI smoke test `test_vba_modules_are_ascii` so non-ASCII can't creep back into VBA modules.

## [1.2.4] — 2026-09-24

### Fixed

- `scripts/Import_VBA_to_Outlook.ps1` now runs VBE Compile + Save after importing. Without the save, restarting Outlook reloaded the old `VbaProject.OTM`, so the buttons ran pre-1.2.1 code that looks for `guri_gui.py` / `VBA\geolocate_headers.py` and did nothing.

## [1.2.3] — 2026-09-24

### Added

- `environment.yml` for Conda/CI; `tests/test_ci_smoke.py` package smoke tests.
- `MSCANToolbar.RecoverAesUi` to rebuild the AES CommandBar after a VBA re-import.

### Fixed

- Conda GitHub Action failed (missing `environment.yml`).
- GURI/Aura ribbon: always direct-launch `guri\gui.py --raise` so a cold start works; broader Python path search; set process working directory to install root.
- Ribbon error text points at Compile + `RecoverAesUi` when the VBA toolbar is missing.

### Changed

- AGENTS: bump SemVer on every meaningful patch commit.

## [1.2.2] — 2026-09-24

### Added

- `scripts/Import_VBA_to_Outlook.ps1` (+ `.bat`) to remove/re-import AES `MSCAN*` modules into the running Outlook VBA project (AccessVBOM required; optional `-SyncThisOutlookSession`).
- `docs/bug_tracker.md` for open/closed suite defects.

### Changed

- `VBA/IMPORT.txt`, README, and AGENTS point at the automated import path.

## [1.2.1] — 2026-09-24

### Changed

- Python code reorganised into packages under the install root:
  - `aes/` — scan engine (`geolocate_headers.py`), checks, dialogs, `action_handler.py`, `secret_store.py`
  - `aes/scanners/` — attachment / body / link / AV / Norton / Defender scanners
  - `guri/` — `core.py`, `gui.py`, `gui_service.py` and helpers (the `guri_` prefix is dropped)
  - `aura/` — data-broker removal (was `broker_removal/`)
  - `geofooter/` — `paths.py`, `version.py`, `crashlog.py`
- `VBA\` now holds only Outlook modules; its Python scripts, `icons\` and `pages\` moved to `aes\` and `assets\`.
- GURI recipients, important domains and reply status are read from `datastore\`.
- VBA, the ribbon host, and the launchers in `scripts\` resolve every script through the install root.

### Fixed

- The newer `link_scanner` copy (previously shadowed by a stale root duplicate) is now the one that loads.

## [1.2.0] — 2026-09-24

### Added

- **Aura** (Aliniant Universal Removal Application): Outlook ribbon / toolbar button, `aes_aura` icons, GURI `--aura` / IPC `AURA`, broker-removal workspace tab.
- Threat-intelligence RiskTable multi-engine **consensus scoring** (VirusTotal lone hits such as 1/91 stay clean).
- Clickable footer metrics (Attachments / Beacons / Links) via `aes://open-report`.
- Suite versioning: `VERSION`, `CHANGELOG.md`, expanded `version.py`, ribbon host assembly version aligned.
- Agent/project docs: `AGENTS.md`, `SKILLS.md`, `WORKLOG.md`.

### Changed

- Create-GURI persists document paths: `guri_records.avg_risk` widened to **TEXT**; failed inserts raise instead of silent success.
- Document-type radios share one exclusive group across Standard / Aliniant / Sensitive.
- RiskTable banner wording: `{n} found on RiskTable`.
- Threat-intel all-OK line hints: “Run Full or Deep scan to see these in full.”
- IP / RDAP allocation preference over misleading BGP org labels where applicable.

### Security

- Local DB credentials (`guri_*_config.json`) are **not** published; use `*.example.json` templates only.
- Rotate any passwords that previously lived in local git history before this public release.

## [1.1.0] — 2026-08-10

### Added

- User rules with compile + local enforcement; Locations file scrape.
- Identity / not-for-me deadlines; timeline importance colours; scrape lookback months.
- About tab + `version.py`.

## [1.0.0] — 2026-01-02

### Added

- GURI database viewer, Outlook scrape, deadlines, learning loop, Ollama.

[Unreleased]: https://github.com/youcangetjules/geofooter/compare/v1.2.0...HEAD
[1.2.0]: https://github.com/youcangetjules/geofooter/releases/tag/v1.2.0
[1.1.0]: https://github.com/youcangetjules/geofooter/compare/v1.0.0...v1.1.0
[1.0.0]: https://github.com/youcangetjules/geofooter/releases/tag/v1.0.0
