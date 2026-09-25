# GeoFooter

**Aliniant Email Scanner (AES)**, **GURI**, and **Aura** — scan and track emails, stamp unique record IDs, and manage data-broker opt-outs.

| Product             | Role                                                                       |
| ------------------- | -------------------------------------------------------------------------- |
| **AES (GeoFooter)** | Outlook watch → Python scan → risk banner + HTML footer (`aes://` actions) |
| **GURI**            | Desktop GUI for unique IDs, scrape, deadlines, library                     |
| **Aura**            | Aliniant Universal Removal Application (broker opt-out workspace)          |
| **AesRibbonHost**   | Optional COM/ribbon host for AES UI in Outlook                             |

**Version:** see [`VERSION`](VERSION) / [`geofooter/version.py`](geofooter/version.py) — currently **1.3.40**.  
**Changes:** [`CHANGELOG.md`](CHANGELOG.md) · day-to-day notes in [`WORKLOG.md`](WORKLOG.md) · bugs in [`docs/bug_tracker.md`](docs/bug_tracker.md).

## Paths / Conda

Set **Install root** in GURI → **Database** (saves `%LOCALAPPDATA%\GeoFooter\install_root.txt`).
All suite folders below are relative to that root.
Override with env `GEOFOOTER_ROOT`. See `geofooter/paths.py` / `VBA\MSCANPaths.bas`.

**Conda (optional, also used by CI):**

```bash
conda env create -f environment.yml
conda activate geofooter
# Windows only (Outlook COM):
pip install pywin32>=306
```

| Folder       | Contents                                                                     |
| ------------ | ---------------------------------------------------------------------------- |
| `aes/`       | Scan engine (`geolocate_headers.py`), dialogs, `aes://` handler, `scanners/` |
| `guri/`      | GURI core (`core.py`) and desktop app (`gui.py`)                             |
| `aura/`      | Data-broker removal workspace (GURI tab)                                     |
| `geofooter/` | Shared install-root paths, version, crash logging                            |
| `VBA/`       | Outlook VBA modules only (`.bas` / `.cls`)                                   |
| `assets/`    | Brand, ribbon icons, pages, Outlook forms, sample data                       |
| `datastore/` | Local DBs and GURI lists (not committed)                                     |
| `docs/`      | Feature guides                                                               |
| `scripts/`   | Migrations, launchers, icon build helpers                                    |
| `tests/`     | Ad-hoc test scripts                                                          |

## Requirements

- Windows + Microsoft Outlook (classic)
- Python 3.10+ with deps from [`requirements.txt`](requirements.txt)
- PostgreSQL for GURI (`guri_postgres_config.example.json`). SQLite is deprecated.
- Optional: .NET Framework 4.8 SDK to build `AesRibbonHost`

## Quick start

On Windows, run [`scripts/Install_GeoFooter.bat`](scripts/Install_GeoFooter.bat). One run installs both products: the **AES email tool** (Python scan engine, `aes://`, Outlook VBA, ribbon) and the **GURI GUI** (desktop shortcut plus `scripts\Launch_GURI_GUI.bat`). **PostgreSQL is required** for GURI. Install it from [postgresql.org/download/windows](https://www.postgresql.org/download/windows/), create database `guri_db`, and set the password in `guri_postgres_config.json`.

Manual steps, if you prefer:

1. Copy `guri_postgres_config.example.json` → `guri_postgres_config.json` and set local credentials (**never commit** the real file).
2. Import VBA modules per [`VBA/IMPORT.txt`](VBA/IMPORT.txt), or run [`scripts/Import_VBA_to_Outlook.bat`](scripts/Import_VBA_to_Outlook.bat) with Outlook open (AccessVBOM required — see [`docs/bug_tracker.md`](docs/bug_tracker.md)).
3. Register the `aes://` footer actions: `.venv\Scripts\python.exe aes\action_handler.py --register`.
4. Ribbon: close Outlook → `AesRibbonHost\install.ps1` → reopen.
5. GURI: `scripts\Launch_GURI_GUI.bat` (or `python guri\gui.py` from the install root).

## Versioning

We use **SemVer** (`MAJOR.MINOR.PATCH`). **About 90% of bumps are patches.**

| Bump      | When                                                         |
| --------- | ------------------------------------------------------------ |
| **PATCH** | Default — fixes, wording, scoring tweaks, icons, docs        |
| **MINOR** | New backward-compatible features (Aura, tabs, scan modes)    |
| **MAJOR** | Rare — breaking Outlook/VBA contracts, schema, or public CLI |

**Git commits:** every commit message starts with a short caption of what was done
(imperative, specific), e.g. `Widen avg_risk to TEXT so Create-GURI saves file paths`.

When releasing:

1. Update `VERSION`, `geofooter/version.py` (`VERSION`, `VERSION_STAMP`, `RELEASE_NOTES`), `CHANGELOG.md`, and `AesRibbonHost\AesRibbonHost.csproj` `<Version>`.
2. Commit with a caption, tag `vX.Y.Z`, push tags to GitHub.
3. Note VBA re-import / ribbon reinstall in the changelog when required.

```bash
git tag -a v1.2.0 -m "Ship suite 1.2.0: Aura, TI consensus, Create-GURI persist"
git push origin v1.2.0
```

## Agent / contributor notes

See [`AGENTS.md`](AGENTS.md) and [`SKILLS.md`](SKILLS.md). Do not commit secrets, DPAPI blobs, or live DB passwords.

All of the API Keys for the Threat websites have free tiers.  

We did that on purpose :) - pretty powerful even at that level.

## License / ownership

© Aliniant — internal / product suite. Public mirror: [youcangetjules/geofooter](https://github.com/youcangetjules/geofooter).
If you can download it and install it - have fun!  
Need help with it - julian.garrett@aliniant.com  
I take cash and card. And feature requests.  
Rule for life - drive it like its stolen :)
