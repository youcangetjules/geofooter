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
