# GeoFooter

**Aliniant Email Scanner (AES)**, **GURI**, and **Aura** — scan and track emails, stamp unique record IDs, and manage data-broker opt-outs.

| Product | Role |
|---------|------|
| **AES (GeoFooter)** | Outlook watch → Python scan → risk banner + HTML footer (`aes://` actions) |
| **GURI** | Desktop GUI for unique IDs, scrape, deadlines, library |
| **Aura** | Aliniant Universal Removal Application (broker opt-out workspace) |
| **AesRibbonHost** | Optional COM/ribbon host for AES UI in Outlook |

**Version:** see [`VERSION`](VERSION) / [`version.py`](version.py) — currently **1.2.0**.  
**Changes:** [`CHANGELOG.md`](CHANGELOG.md) · day-to-day notes in [`WORKLOG.md`](WORKLOG.md).

## Paths

Set **Install root** in GURI → **Database** (saves `%LOCALAPPDATA%\GeoFooter\install_root.txt`).
Suite folders (`assets/`, `docs/`, `scripts/`, `tests/`, `datastore/`, `debuglog/`, `crashlogs/`, `VBA/`) are relative to that root.
Override with env `GEOFOOTER_ROOT`. See `geofooter_paths.py` / `VBA\MSCANPaths.bas`.

| Folder | Contents |
|--------|----------|
| `VBA/` | Canonical Outlook VBA + scan/dialog Python |
| `assets/` | Brand, ribbon icons, Outlook forms, sample data |
| `docs/` | Feature guides |
| `scripts/` | Migrations, launchers, icon build helpers |
| `tests/` | Ad-hoc test scripts |

## Requirements

- Windows + Microsoft Outlook (classic)
- Python 3.10+ with deps from [`requirements.txt`](requirements.txt)
- Optional: PostgreSQL for GURI (`guri_postgres_config.example.json`)
- Optional: .NET Framework 4.8 SDK to build `AesRibbonHost`

## Quick start

1. Copy `guri_postgres_config.example.json` → `guri_postgres_config.json` and set local credentials (**never commit** the real file).
2. Import VBA modules per [`VBA/IMPORT.txt`](VBA/IMPORT.txt).
3. Prefer scanners under `VBA\` (canonical), e.g. `VBA\geolocate_headers.py`.
4. Ribbon: close Outlook → `AesRibbonHost\install.ps1` → reopen.

## Versioning

We use **SemVer** (`MAJOR.MINOR.PATCH`). **About 90% of bumps are patches.**

| Bump | When |
|------|------|
| **PATCH** | Default — fixes, wording, scoring tweaks, icons, docs |
| **MINOR** | New backward-compatible features (Aura, tabs, scan modes) |
| **MAJOR** | Rare — breaking Outlook/VBA contracts, schema, or public CLI |

**Git commits:** every commit message starts with a short caption of what was done
(imperative, specific), e.g. `Widen avg_risk to TEXT so Create-GURI saves file paths`.

When releasing:

1. Update `VERSION`, `version.py` (`VERSION`, `VERSION_STAMP`, `RELEASE_NOTES`), `CHANGELOG.md`, and `AesRibbonHost\AesRibbonHost.csproj` `<Version>`.
2. Commit with a caption, tag `vX.Y.Z`, push tags to GitHub.
3. Note VBA re-import / ribbon reinstall in the changelog when required.

```bash
git tag -a v1.2.0 -m "Ship suite 1.2.0: Aura, TI consensus, Create-GURI persist"
git push origin v1.2.0
```

## Agent / contributor notes

See [`AGENTS.md`](AGENTS.md) and [`SKILLS.md`](SKILLS.md). Do not commit secrets, DPAPI blobs, or live DB passwords.

## License / ownership

© Aliniant — internal / product suite. Public mirror: [youcangetjules/geofooter](https://github.com/youcangetjules/geofooter).
