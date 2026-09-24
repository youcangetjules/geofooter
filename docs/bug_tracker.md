# GeoFooter / AES / GURI — Bug tracker

Living list of known defects and follow-ups. Newest open items at the top of each section.

Use this for **bugs and regressions**, not day-to-day build notes (`WORKLOG.md`) or shipped releases (`CHANGELOG.md`).

## How to use

1. Add a new entry under **Open** with the next free `BUG-NNN` id.
2. Fill Severity, Area, Status, and the short repro / impact.
3. When fixed, move the entry to **Closed** and set **Fixed in** (commit hash and/or version).
4. Prefer one bug per entry. Link related commits or PR URLs when useful.

### Severity

| Level | Meaning |
| ----- | ------- |
| **S1** | Data loss, wrong security decision, or Outlook hang on the UI thread |
| **S2** | Feature broken for a main path (scan, footer, GURI save, ribbon) |
| **S3** | Workaround exists; annoying or partial failure |
| **S4** | Cosmetic, docs, or rare edge case |

### Areas

`AES` · `VBA` · `GURI` · `Aura` · `Ribbon` · `Install` · `Docs`

### Status

`open` · `investigating` · `fixed` · `wontfix` · `duplicate`

---

## Open

### BUG-002 — VBA modules on disk are ahead of Outlook after 1.2.1 package move
- **Severity:** S2
- **Area:** VBA / Install
- **Status:** open
- **Reported:** 2026-09-24
- **Summary:** Outlook still runs the pre-restructure module set until the updated `.bas`/`.cls` files are imported. Scans may still look for old script paths (`VBA\geolocate_headers.py`, root `guri_gui.py`, etc.).
- **Repro:** Edit `VBA\MSCANPaths.bas` (or pull 1.2.1+); open Outlook without re-import; trigger a scan or Settings/GURI from the toolbar.
- **Workaround:** Re-import per `VBA\IMPORT.txt`, or use `scripts\Import_VBA_to_Outlook.ps1` once AccessVBOM is enabled. Then Debug → Compile VBAProject → Save → restart Outlook.
- **Follow-up:** Finish documenting the auto-import script; confirm AccessVBOM registry path on this Outlook build.

### BUG-001 — Outlook Trust Center UI has no “Trust access to the VBA project object model” checkbox
- **Severity:** S3
- **Area:** Install / VBA
- **Status:** open
- **Reported:** 2026-09-24
- **Summary:** Macro Settings shows Enable all macros, but not AccessVBOM. External sync of VBA modules fails until the registry value is set.
- **Workaround:**
  ```powershell
  New-Item -Path "HKCU:\Software\Microsoft\Office\16.0\Outlook\Security" -Force | Out-Null
  Set-ItemProperty -Path "HKCU:\Software\Microsoft\Office\16.0\Outlook\Security" -Name "AccessVBOM" -Type DWord -Value 1
  ```
  Fully quit Outlook (tray too), reopen, then run `scripts\Import_VBA_to_Outlook.ps1 -EnableAccessVBOM` if needed.
- **Notes:** “Enable all macros” only allows macros to *run*; it does not grant project object model access.

---

## Closed

_(none yet)_

---

## Template

```markdown
### BUG-NNN — short title
- **Severity:** S2
- **Area:** AES
- **Status:** open
- **Reported:** YYYY-MM-DD
- **Summary:** One or two sentences.
- **Repro:** Steps or “n/a”.
- **Expected / actual:** …
- **Workaround:** … or “none”.
- **Fixed in:** (when closed) commit / version
```
