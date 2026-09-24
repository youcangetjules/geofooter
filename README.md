# GeoFooter

**Aliniant Email Scanner (AES)**, **GURI**, and **Aura** — scan and track emails, stamp unique record IDs, and manage data-broker opt-outs.

| Product | Role |
|---------|------|
| **AES (GeoFooter)** | Outlook watch → Python scan → risk banner + HTML footer (`aes://` actions) |
| **GURI** | Desktop GUI for unique IDs, scrape, deadlines, library |
| **Aura** | Aliniant Universal Removal Application (broker opt-out workspace) |
| **AesRibbonHost** | Optional COM/ribbon host for AES UI in Outlook |

**Version:** see [`VERSION`](VERSION) / [`version.py`](version.py) — currently **2.0.0**.  
**Changes:** [`CHANGELOG.md`](CHANGELOG.md) · day-to-day notes in [`WORKLOG.md`](WORKLOG.md).

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

We use **SemVer** (`MAJOR.MINOR.PATCH`):

| Bump | When |
|------|------|
| MAJOR | Breaking Outlook/VBA contracts, schema, or public CLI |
| MINOR | New features (Aura, new scan modes, tabs) backward-compatible |
| PATCH | Fixes and wording |

When releasing:

1. Update `VERSION`, `version.py` (`VERSION`, `VERSION_STAMP`, `RELEASE_NOTES`), `CHANGELOG.md`, and `AesRibbonHost\AesRibbonHost.csproj` `<Version>`.
2. Commit, tag `vX.Y.Z`, push tags to GitHub.
3. Note VBA re-import / ribbon reinstall in the changelog when required.

```bash
git tag -a v2.0.0 -m "GeoFooter suite 2.0.0"
git push origin v2.0.0
```

## Agent / contributor notes

See [`AGENTS.md`](AGENTS.md) and [`SKILLS.md`](SKILLS.md). Do not commit secrets, DPAPI blobs, or live DB passwords.

## License / ownership

© Aliniant — internal / product suite. Public mirror: [youcangetjules/geofooter](https://github.com/youcangetjules/geofooter).
