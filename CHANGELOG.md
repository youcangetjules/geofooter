# Changelog

All notable changes to the **GeoFooter / AES / GURI / Aura** suite are documented here.

Format follows [Keep a Changelog](https://keepachangelog.com/en/1.1.0/).
Versioning follows [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

The canonical version string lives in `version.py` and the plain-text `VERSION` file.
Bump those together with this changelog and `AesRibbonHost/AesRibbonHost.csproj` `<Version>`.

## [Unreleased]

## [2.0.0] — 2026-09-24

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

[Unreleased]: https://github.com/youcangetjules/geofooter/compare/v2.0.0...HEAD
[2.0.0]: https://github.com/youcangetjules/geofooter/releases/tag/v2.0.0
[1.1.0]: https://github.com/youcangetjules/geofooter/compare/v1.0.0...v1.1.0
[1.0.0]: https://github.com/youcangetjules/geofooter/releases/tag/v1.0.0
