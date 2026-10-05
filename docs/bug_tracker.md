# GeoFooter / AES / GURI — Bug tracker

Living list of known defects and follow-ups. Newest open items at the top of each section.

Use this for **bugs and regressions**, not day-to-day build notes (`WORKLOG.md`) or shipped releases (`CHANGELOG.md`).

## How to use

1. Add a new entry under **Open** with the next free `BUG-NNN` id.
2. Fill Severity, Area, Status, and the short repro / impact.
3. When fixed, move the entry to **Closed** and set **Fixed in** (commit hash and/or version).
4. Prefer one bug per entry. Link related commits or PR URLs when useful.

### Severity

| Level  | Meaning                                                              |
| ------ | -------------------------------------------------------------------- |
| **S1** | Data loss, wrong security decision, or Outlook hang on the UI thread |
| **S2** | Feature broken for a main path (scan, footer, GURI save, ribbon)     |
| **S3** | Workaround exists; annoying or partial failure                       |
| **S4** | Cosmetic, docs, or rare edge case                                    |

### Areas

`AES` · `VBA` · `GURI` · `Aura` · `Ribbon` · `Install` · `Docs`

### Status

`open` · `investigating` · `fixed` · `wontfix` · `duplicate`

---

## Open

### BUG-008 — Footer never applies while a draft is open, and a stale job empties the mail body

- **Severity:** S1
- **Area:** VBA
- **Status:** fixed (pending operator verification after VBA re-import)
- **Reported:** 2026-10-05
- **Summary:** Two faults in one incident. (1) `IsUserComposing` answered True for as long as any unsent draft Inspector was open, idle or not. A reply left open in the background deferred every `NudgeAsyncWork` from 09:00 to 09:30 (977 compose deferrals): no footer, no job timeout, 60 mails stuck in the queue. (2) A Short Scan rescan of the same mail then reconciled the 26-minute-old job for it straight after the new export. The commit hit `-2147221239` twice and Outlook's in-memory copy came back with no body properties at all (`HTMLBody` = a 241-char "Converted from text/plain format" shell, `Body` empty, `Saved = False`, AES mark + status PNGs attached). Any later Save of that object, including the next footer pass, would have made the loss permanent.
- **Repro:** Leave a reply draft open, let mail arrive, then Short Scan a message whose earlier compact job is still pending.
- **Fix:** `MSCANIdle.IsUserComposing` needs keyboard/mouse input in the last 4s (`GetLastInputInfo`) and a compose window in the foreground (caption match). `StartAsyncGeolocationJob` supersedes pending jobs for the same EntryID. `CommitMailHtml` refuses to save HTML with no sender content outside the AES blocks, and on a failed commit it discards the unsaved changes (`Close olDiscard`, skipped for open inspectors). `ApplyHighRiskTextOnlyMail` has the same empty-body guard. The footer paths neutralise beacons on the HTML string (`NeutralizeBeaconsInHtml`) instead of an extra write + Save.
- **Fixed in:** 1.4.10, with the dropped on-screen footer corrected in 1.4.11. Requires a VBA re-import.
- **Recovery for an affected mail:** select a different message (or restart Outlook) without saving the affected one so Outlook drops the in-memory copy and reloads the stored body.

### BUG-007 — Reading pane flickers and Outlook stops responding (footer commit retry storm)

- **Severity:** S1
- **Area:** VBA
- **Status:** fixed (pending operator verification after VBA re-import)
- **Reported:** 2026-09-29
- **Summary:** Outlook repeatedly redrew the same message and became unusable. When a mail would not accept a footer, `mail.Save` failed with `-2147221239` ("the operation cannot be performed because the message has been changed"). That error was treated as retryable at three nested levels: `CompleteAsyncFooter` retried the apply 3 times, each apply tried the banner then the no-banner body, and `CommitMailHtml` tried 5 times at 1s apart. One message could take up to 30 `HTMLBody` writes over ~30s, and every write repaints the item. `scripts/aes_log_health.py --session` on the live log showed 60 conflicts in 11 minutes, with a single message accounting for 6 full rounds.
- **Repro:** Let a mail arrive that the store keeps re-syncing (IMAP/Gmail), or restart Outlook with a queue backlog. The message list flickers continuously while the queue drains.
- **Fix:** `MSCANModule1` detects the conflict error explicitly (`IsMessageChangedError`) and records it in `m_LastCommitConflict`. `CommitMailHtml` is capped at 3 attempts 400ms apart and gives up after one re-resolve on a conflict; `InsertFooterIntoMail`, `ReplaceAesFooterInMail`, and `CompleteAsyncFooter` skip their own retries when the item is busy. The message is left unfooted for a later queue pass instead of being rewritten.
- **Fixed in:** 1.4.6. Requires a VBA re-import (new module `MSCANHealth.bas`).
- **Detection:** `MSCANHealth` now raises a `HEALTH` warning after 3 consecutive conflicts, on a queue backlog over 25, and on any operation holding the UI thread over 1.2s.

### BUG-006 — Typing in Outlook freezes and drops letters

- **Severity:** S1
- **Area:** VBA
- **Status:** fixed (pending operator verification after VBA re-import)
- **Reported:** 2026-09-25
- **Summary:** While writing a mail, the editor froze intermittently and roughly half the typed characters never appeared. AES waited on Outlook's UI thread with `DoEvents` spin loops: `PauseSeconds`, the `CompleteAsyncFooter` apply retry, `CommitMailHtml` (up to 5 x 1s), `RunCommandAndCaptureOutput` (spun for the entire Python run), `MSCANModWatchers.YieldBriefly`, and two send-conflict retries. A `Do While ... DoEvents ... Loop` re-enters Outlook's message pump at full CPU, so keystrokes were dispatched through VBA instead of the compose editor. Footer apply and queue drain also ran while a compose window was open, holding the UI thread during `HTMLBody` reads and writes.
- **Repro:** Open a new mail and type continuously while a scan is in flight (or within ~90s of the catch-up heartbeat). Characters are dropped and the window stalls for about a second at a time.
- **Fix:** New `VBA\MSCANIdle.bas` — `WaitMs` sleeps via kernel32 in 50ms slices without pumping messages, and `IsUserComposing` detects a compose Inspector or an inline reply. Every spin loop now sleeps, and `NudgeAsyncWork` / `NudgeQueueWork` / `ProcessQueue` defer while a message is being written. Queued mail drains when the window closes.
- **Fixed in:** 1.4.5. Requires a VBA re-import (new module `MSCANIdle.bas`).

### BUG-005 — Ribbon macro fallback can never work (Outlook has no `Application.Run`)

- **Severity:** S3
- **Area:** Ribbon
- **Status:** open
- **Reported:** 2026-09-24
- **Summary:** When the AES CommandBar button isn't found, `Connect.TryRunVbaMacro` calls `Application.Run`. The Outlook object model has no `Run`, so every attempt fails with "'ApplicationClass' does not contain a definition for 'Run'". Ribbon scan buttons therefore depend entirely on the VBA toolbar existing.
- **Repro:** Delete or skip the `AES` CommandBar, then click Short Scan on the ribbon. `AesRibbonHost.log` shows `TryRunVbaMacro miss ...` for every candidate.
- **Expected / actual:** The scan runs / nothing happens, and a "no matching button/macro" message appears.
- **Workaround:** Alt+F8 → `CreateToolbar` (or `RecoverAesUi`), or restart Outlook so `MSCANAppBootstrap.DeferredToolbar` rebuilds the bar.
- **Follow-up:** Remove the dead `Run` path for the scan buttons. Settings (1.4.2) no longer depends on it: if the toolbar button is missing, the ribbon starts `aes/settings_dialog.py` directly.
- **Update 2026-09-25:** Settings click logged `no matching button for tag=AES_SETTINGS` and four `Application.Run` misses. Direct launch is the Settings fallback. Short Scan and the other VBA buttons still need the toolbar.

### BUG-003 — Buttons do nothing / old dialogs appear: VBA import never saved to `VbaProject.OTM`

- **Severity:** S2
- **Area:** VBA / Install
- **Status:** investigating
- **Reported:** 2026-09-24
- **Summary:** After `Import_VBA_to_Outlook.bat`, the new modules only lived in memory. `%APPDATA%\Microsoft\Outlook\VbaProject.OTM` was still dated 2026-09-23 23:35, so each Outlook restart reloaded pre-1.2.1 code. That code looks for deleted files (`guri_gui.py`, `VBA\geolocate_headers.py`, `VBA\aes_settings_dialog.py`). As a result the Scan, GURI and Aura buttons fire but silently fail. Settings falls back to the legacy InputBox "Scan Accounts" dialog.
- **Repro:** Run the import bat, don't press Save in the VBA editor, restart Outlook, click GURI or Settings.
- **Evidence:** `VBA_Log.txt` shows `ShowGuriGui: python or guri_gui.py not found`, `GetPythonScript: Python script (geolocate_headers.py) not found`, `MSCANSettings: Python settings dialog unavailable; InputBox fallback`. `AesRibbonHost.log` shows the old DLL (`LaunchGuriGuiDirect: missing ... script=`).
- **Expected / actual:** Buttons launch the scan / GURI / Aura, and Settings opens the PySide6 dialog / nothing happens, or the old InputBox appears.
- **Workaround:** Re-run the import bat (1.2.4+ executes VBE Compile + Save and prints `File > Save: done`). Otherwise press Alt+F11 → Debug → Compile → Ctrl+S. Then fully quit Outlook, run `AesRibbonHost\install.ps1` with Outlook closed, and reopen.
- **Update 2026-09-24 18:08:** The external import can't reach VBE on this machine at all. `Application.VBE` is null from outside even with AccessVBOM=1 and a fresh Outlook start. GURI/Aura work only through the ribbon's direct launch; the scanner (pure VBA) stays broken.
- **Update 2026-09-24 18:30:** AccessVBOM=1 is confirmed and `Application.VBE` is still Nothing both externally and from `UpdateAesFromDisk`; workaround is manual File > Import (or updated SelfUpdate fallback in 1.2.8).
- **Fixed in:** 1.2.4 (`54f84b4`) import script saves. 1.2.6 adds the in-Outlook updater `MSCANSelfUpdate.UpdateAesFromDisk` (import that one file manually once). Pending operator verification.

### BUG-002 — VBA modules on disk are ahead of Outlook after 1.2.1 package move

- **Severity:** S2
- **Area:** VBA / Install
- **Status:** open
- **Reported:** 2026-09-24
- **Summary:** Outlook still runs the pre-restructure module set until the updated `.bas`/`.cls` files are imported. Scans may still look for old script paths (`VBA\geolocate_headers.py`, root `guri_gui.py`, etc.).
- **Repro:** Edit `VBA\MSCANPaths.bas` (or pull 1.2.1+); open Outlook without re-import; trigger a scan or Settings/GURI from the toolbar.
- **Workaround:** Run `scripts\Import_VBA_to_Outlook.bat` with Outlook open (use `-EnableAccessVBOM` once if needed, quit Outlook fully, then run again). Or re-import manually per `VBA\IMPORT.txt`. Then Debug → Compile VBAProject → Save → restart Outlook.
- **Follow-up:** Operator still needs to run the import on each machine after pull.

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
  
  Fully quit Outlook (tray too), reopen, then run `scripts\Import_VBA_to_Outlook.bat`
  (or `Import_VBA_to_Outlook.ps1 -EnableAccessVBOM` once to set the key).

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
