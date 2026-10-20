# GeoFooter / AES / GURI — Bug tracker

Living list of known defects and follow-ups. Newest open items at the top of each section.

Use this for **bugs and regressions**, not day-to-day build notes (`WORKLOG.md`) or shipped releases (`CHANGELOG.md`).

## How to use

1. Add a new entry under **Open** with the next free `BUG-NNN` id.
2. Fill Severity, Area, Status, and the short repro / impact.
3. When fixed, move the entry to **Closed** and set **Fixed in** (commit hash and/or version).
4. Prefer one bug per entry. Link related commits or PR URLs when useful.
5. **Log every bug seen, as soon as it is seen** — in a log, a screenshot, a test, or while reading code — even if it is not fixed in the same change, and even if it looks minor. When a fix ships, update the entry in the same commit. A recurring error line in `VBA_Log.txt` counts as a bug; a symptom with no known cause goes in `performance.md` first.

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

### BUG-019 — AES actions in text-only (mitigated) mail cannot be clicked

- **Severity:** S3
- **Area:** VBA
- **Status:** fixed for high-risk mitigation (pending VBA re-import); still open for Full No Trust
- **Reported:** 2026-10-20
- **Summary:** High-risk mitigation (score > 70) and Full No Trust rewrote the mail as plain text, listing actions as raw `aes://…` lines. Outlook only auto-links standard schemes (`file://`, `https://`) in plain text, so "Restore original HTML format" and every other action was dead text (seen on a Substack mail, 20 Oct).
- **Fixed in:** 1.5.1 for high-risk mitigation: the body is now AES-built HTML (restore link on top, actions as labelled links, sender text escaped). Full No Trust (`ApplyFullNoTrustTextMail`, `action_handler._apply_full_no_trust_text`) still writes plain text on purpose and keeps dead links.

### BUG-018 — Blocked tracking pixels show as a broken-image box

- **Severity:** S4
- **Area:** VBA
- **Status:** open
- **Reported:** 2026-10-20
- **Summary:** `MSCANModSenderRules.DefangBeaconTag` points a blocked beacon's `src` at `file:///<install root>/assets/pages/aes_blocked_pixel.png`. The file exists, but Outlook will not load local `file://` images in mail, so each blocked pixel renders as a red ✕ box (seen below the AES action buttons on a full scan).
- **Workaround:** none needed; nothing is fetched remotely.
- **Follow-up:** Swap `src` for an inline `data:` 1x1 GIF, or set `width`/`height` 0 and `display:none` on the defanged tag.

### BUG-017 — Bottom border of the AES action buttons cut off

- **Severity:** S4
- **Area:** AES
- **Status:** fixed (applies to mail scanned after 1.4.30)
- **Reported:** 2026-10-20
- **Summary:** In full scans the action buttons are the last row of the footer. Their cells set a 10px label with no line height, so Outlook's Word renderer sized them from the inherited line height and clipped the bottom border.
- **Fixed in:** 1.4.30. Fixed 14px line height (`mso-line-height-rule:exactly`), extra bottom padding, and an 8px spacer row under the buttons.

### BUG-014 — WHOIS lookups always fail: wrong `whois` package in the venv

- **Severity:** S3
- **Area:** AES / Install
- **Status:** open
- **Reported:** 2026-10-20
- **Summary:** Every WHOIS lookup in `aes/geolocate_headers.py` fails with "module 'whois' has no attribute 'whois'". The `.venv` has the `whois` package (which exposes `query()`) instead of `python-whois` (which exposes `whois.whois()`), and both install as the module `whois`. Domain age and registrant data are missing from every report; the scan carries on without them.
- **Repro:** `.venv\Scripts\python.exe -c "import whois; whois.whois('example.com')"` raises AttributeError.
- **Workaround:** `pip uninstall whois` then `pip install python-whois` in the venv.
- **Follow-up:** Guard the import so the wrong package is reported once instead of per lookup.

### BUG-013 — `StripSensitiveHeadersFromMail` fails on every footer apply

- **Severity:** S4
- **Area:** VBA
- **Status:** open
- **Reported:** 2026-10-20
- **Summary:** `MSCANModule1.StripSensitiveHeadersFromMail` tries to rewrite `PR_TRANSPORT_MESSAGE_HEADERS` (`0x007D001F`) and Outlook refuses with `-2147024891` (access denied, "does not support this operation"). Logged 1,391 times since 2026-10-01, roughly once per footer. The headers are never stripped, and each attempt parses the whole header block on the UI thread for nothing.
- **Repro:** Let any mail be footered; `VBA_Log.txt` logs the error straight after "Footer inserted successfully".
- **Workaround:** none needed; the footer still applies.
- **Follow-up:** Either drop the strip (the store treats the property as read-only) or stop after the first refusal in a session.

### BUG-011 — Ribbon shows "AES button '…' was not found" while the scan is running

- **Severity:** S3
- **Area:** Ribbon
- **Status:** fixed (pending operator verification after `install.ps1`)
- **Reported:** 2026-10-20
- **Summary:** `TryExecuteOnBar` wrapped `ctl.Execute()` in the same catch as the button match. When the VBA macro started and `Execute` then threw back, the host treated it as no match, fell through to the dead `Application.Run` path (BUG-005) and showed the popup. Seen on 24 Sep, 28 Sep, 16 Oct and 20 Oct.
- **Fixed in:** 1.4.25 (`AesRibbonHost/Connect.cs`). Needs Outlook fully quit and `AesRibbonHost\install.ps1`.

### BUG-010 — PC sleep logged as multi-hour UI stalls

- **Severity:** S4
- **Area:** VBA
- **Status:** fixed (pending VBA re-import)
- **Reported:** 2026-10-20
- **Summary:** `MSCANHealth.EndOp` measured with `GetTickCount`, which keeps counting while Windows sleeps. Operations that straddled a sleep were logged as stalls of 79,700,593 ms, 70,033,485 ms and 8,052,375 ms, which skewed the worst-stall stat and the HEALTH warnings.
- **Fixed in:** 1.4.26 (`c1d4f9c`). Readings over 300 s are logged once as a probable sleep and excluded.

### BUG-009 — Outlook frozen while AES Settings is open

- **Severity:** S1
- **Area:** VBA
- **Status:** fixed (pending VBA re-import)
- **Reported:** 2026-10-20
- **Summary:** `MSCANSettings` ran the Python settings dialog with `shell.Run(cmd, 1, True)`, so Outlook stopped responding (no mail, no typing) until the dialog was closed. The ribbon Settings button takes the same path through the VBA toolbar.
- **Fixed in:** 1.4.26 (`c1d4f9c`). The dialog runs through `aes_settings_launch.vbs`; the result is applied on the next ItemLoad. Re-import `MSCANSettings`, `MSCANModule1`, `MSCANHealth`.

### BUG-008 — Footer never applies while a draft is open, and a stale job empties the mail body

- **Severity:** S1
- **Area:** VBA
- **Status:** fixed (pending operator verification after VBA re-import)
- **Reported:** 2026-10-05
- **Summary:** Two faults in one incident. (1) `IsUserComposing` answered True for as long as any unsent draft Inspector was open, idle or not. A reply left open in the background deferred every `NudgeAsyncWork` from 09:00 to 09:30 (977 compose deferrals): no footer, no job timeout, 60 mails stuck in the queue. (2) A Short Scan rescan of the same mail then reconciled the 26-minute-old job for it straight after the new export. The commit hit `-2147221239` twice and Outlook's in-memory copy came back with no body properties at all (`HTMLBody` = a 241-char "Converted from text/plain format" shell, `Body` empty, `Saved = False`, AES mark + status PNGs attached). Any later Save of that object, including the next footer pass, would have made the loss permanent.
- **Repro:** Leave a reply draft open, let mail arrive, then Short Scan a message whose earlier compact job is still pending.
- **Fix:** `MSCANIdle.IsUserComposing` needs keyboard/mouse input in the last 4s (`GetLastInputInfo`) and a compose window in the foreground (caption match). `StartAsyncGeolocationJob` supersedes pending jobs for the same EntryID. `CommitMailHtml` refuses to save HTML with no sender content outside the AES blocks, and on a failed commit it discards the unsaved changes (`Close olDiscard`, skipped for open inspectors). `ApplyHighRiskTextOnlyMail` has the same empty-body guard. The footer paths neutralise beacons on the HTML string (`NeutralizeBeaconsInHtml`) instead of an extra write + Save.
- **Fixed in:** 1.4.10, with the dropped on-screen footer corrected in 1.4.11. 1.4.12 applies a finished footer to the message on screen instead of deferring it until the 90s deadline deletes the job. Requires a VBA re-import.
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
- **Update 2026-10-20:** The false popup when the toolbar button *was* found is BUG-011 (fixed 1.4.25). This entry stays open for the dead `Run` path itself.

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

### BUG-021 — Images under cloud/ were corrupted in git by the eol=lf rule

- **Severity:** S3
- **Area:** repo
- **Status:** fixed
- **Reported:** 2026-10-20
- **Summary:** 1.5.4 added `cloud/** text eol=lf` to `.gitattributes` so bash and systemd files keep LF. It also forced the text filter onto binary files, so committing `cloud/site/aliniant.net/aliniant-shire.jpg` and `favicon.ico` in 1.5.7 (`361bd27`) rewrote their CRLF byte pairs. The blobs in the repo no longer matched the files (git warned "CRLF will be replaced by LF"). The deployed copies were not affected, because they were copied from the working tree.
- **Repro:** `git hash-object --no-filters <file>` vs `git rev-parse HEAD:<file>` differed for both files.
- **Fixed in:** 1.5.8. `cloud/**/*.jpg|png|ico binary` comes after the LF rule, and both files were re-staged; the blobs now match the files on disk.

### BUG-020 — Pixel nginx site could have taken over i-txone.com's catch-all traffic

- **Severity:** S2
- **Area:** cloud
- **Status:** fixed (before deploy)
- **Reported:** 2026-10-20
- **Summary:** `workflow-api` on i-txone.com has no `default_server`, so nginx sends any request whose Host/SNI matches no `server_name` (such as `www.i-txone.com`, other names covered by the wildcard DNS, or requests by IP) to whichever server block comes first on the socket. If the pixel site had been installed under a name that sorts before `workflow-api` (e.g. `aes-pixel`), it would have become that catch-all and answered 404 to that production traffic. The site files also added `[::]:80` / `[::]:443` listeners, which opens new sockets that `workflow-api` does not have.
- **Repro:** Found by reading the stage-1/2 files against `nginx -T` on the server. Nothing was deployed.
- **Fixed in:** 1.5.3. The site files say to install as `sites-enabled/zz-aes-pixel` (included after `workflow-api`). The IPv6 listeners are gone, so the listens match `workflow-api` exactly, and a per-IP `limit_conn` of 10 is added to protect the shared worker connections (768).
- **Update 2026-10-20:** 1.5.4 moves the pixel to a dedicated VM (`track.aliniant.net`), so it no longer shares nginx with TXOne at all.
- **Update 2026-10-20:** the operator chose the existing VM after all. 1.5.4's `deploy.sh` had brought the problem back (it enabled the site as `aes-pixel`, and its `apt-get install nginx` could upgrade and restart a running nginx). Fixed in 1.5.5: the site is enabled as `zz-aes-pixel`, only missing packages are installed, and the default site is only removed on a fresh nginx.

### BUG-016 — Shodan shows "lookup failed (HTTPError)" although Settings says OK

- **Severity:** S4
- **Area:** AES
- **Status:** fixed
- **Reported:** 2026-10-20
- **Summary:** The Settings test calls `/api-info`, which succeeds on the free (oss) plan. Scans call `/shodan/host/{ip}`, which the free plan refuses with 403 "Requires membership or higher" for nearly every IP (it answers 8.8.8.8, not 54.240.4.6). `_shodan_ip` only handled 404, so the 403 surfaced as an error in the Reputation column.
- **Fixed in:** 1.4.29 (`5997c5d`). 403 and 429 read as no data with a plain reason; the Settings test flags the free plan.

### BUG-015 — Embedded full analysis takes on the sender's background colour

- **Severity:** S4
- **Area:** AES
- **Status:** fixed
- **Reported:** 2026-10-20
- **Summary:** On a mail with a coloured body (green on an eschools newsletter), the whole embedded "EMAIL SECURITY ANALYSIS" block showed green. The footer set `background:#ffffff` on a `div`, which Outlook's Word renderer ignores, so the sender's body colour showed through.
- **Fixed in:** 1.4.28. The analysis sits in a white `bgcolor` table cell. Rescan older mails to refresh them.

### BUG-012 — Hop ASN shows the receiving server's network (AS15169 instead of AS11377)

- **Severity:** S2
- **Area:** AES
- **Status:** fixed
- **Reported:** 2026-10-20
- **Summary:** In the Route Table, a SendGrid hop showed AS15169 (Google). `_fill_ip_registration` and `_apply_address_and_host_decoding` also resolved the hop's `by_host` (mx.google.com), which overwrote the sender's ASN. A wrong ASN feeds route risk scoring.
- **Fixed in:** 1.4.18. Only the hop's own FQDN and HELO are used, and an ASN is only replaced when blank or equal to the BGP value. Covered by `tests/test_hop_asn.py`.

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
