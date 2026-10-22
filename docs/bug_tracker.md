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

### BUG-035 — Outlook locks up for extended periods (hard UI freeze)

- **Severity:** S1
- **Area:** VBA / AES / Outlook
- **Status:** investigating
- **Reported:** 2026-10-22
- **Summary:** Outlook becomes unresponsive for extended periods again (user report). Prior PERF-002 sample showed idle queue, CatchUp heartbeats, duplicate GURI, and HEALTH uiStalls=0 because CatchUp was not timed. Hard freezes cannot be logged from VBA (UI thread stuck).
- **Repro:** Use Outlook with AES enabled until the UI stops responding for a long stretch.
- **Workaround / capture:** AES Diagnostics → Start lockup scanner (1.5.55+); reproduce; collect `%LOCALAPPDATA%\GeoFooter\Logs\lockup_scans\freeze_*.txt` and `outlook_lockup_events.log`. Re-import VBA so CatchUp stalls appear in HEALTH.
- **Related:** PERF-002, BUG-034

### BUG-034 — Two GURI GUI processes can run at once

- **Severity:** S3
- **Area:** GURI / Ribbon
- **Status:** open
- **Reported:** 2026-10-21
- **Summary:** After Aura/ribbon opens, two full `guri\gui.py --raise` processes can stay alive (each with a short parent launcher under Outlook). Seen at 19:35: PIDs 9684 and 33548 (parents 41552 and 34340), both started 18:44. AesRibbonHost.log at 18:44:21 shows repeated `BringGuriWindowToFront: no GURI window found` then both instances remain. Single-instance `QLocalServer` should have collapsed them; it did not (race on cold start, or second start before the first listened). Extra GURI burns RAM (~150–175 MB each) and can contend with Outlook.
- **Repro:** Cold-start GURI from the ribbon/Aura when no window is up yet; or click Aura/GURI twice while the first is still starting.
- **Workaround:** End the older `pythonw` `guri\gui.py` in Task Manager; keep one. Prefer opening Aura from the footer after GURI is already up.

### BUG-033 — Footer top and second rows centred left of the strip middle

- **Severity:** S4
- **Area:** AES
- **Status:** fixed (pending rescan)
- **Reported:** 2026-10-21
- **Summary:** In Outlook the top summary line and the sender line sat well left of the strip centre (about a third of the way across on a wide reading pane). All rows shared one column grid, and Word spread the table's spare width into the right-hand column (quip / AURA / copyright). The left padding that balanced the right side assumed that column's natural width, so the centre text drifted left. Row 2 was worse because its padding assumed the narrower AURA chip while the column was as wide as the quip.
- **Repro:** Rescan a message and view it in a wide reading pane.
- **Workaround:** None.
- **Fixed in:** 1.5.52. Each row is its own two-cell table with a fixed-width right cell; the centre cell is padded on the left by that width less the logo column.

### BUG-032 — Footer top-row title sits higher than underlined metrics

- **Severity:** S4
- **Area:** AES
- **Status:** fixed (pending rescan)
- **Reported:** 2026-10-21
- **Summary:** On the top summary strip, plain title text ("Aliniant AES Scan Result - Basic Scan") sat ~1–2px higher than the underlined Risk / Attachments / Beacons / Links / HRE runs. `text-decoration:underline` sits outside the line box, so neighbouring cells with `vertical-align:middle` did not share a baseline.
- **Repro:** Rescan a message and compare the title with the metric links on the same row.
- **Workaround:** None.
- **Fixed in:** 1.5.48. Underlines use a 1px `border-bottom` on links; plain title spans get a matching 1px strip-coloured bottom edge. Fragments use `vertical-align:middle`.

### BUG-031 — White hairline under footer letter E

- **Severity:** S4
- **Area:** AES
- **Status:** fixed (pending rescan)
- **Reported:** 2026-10-21
- **Summary:** A thin white horizontal sliver sat under / through the E in the A/E/S column. Measured as `#ffffff`. Per-row letter cells met the rule row and Outlook showed body white at that join.
- **Repro:** Rescan a message and look at the E letter.
- **Workaround:** None.
- **Fixed in:** 1.5.50. A/E/S live in one `rowspan=5` column with slate 1px spacers where the rules are; summary rules never share a cell edge with a letter. 1.5.48 `border-bottom` made it worse (Outlook painted white). 1.5.49 moved the rule out of the letter column but the 1px filler cell still flashed white.
- **Update 2026-10-21 (17:31):** user confirmed top-row text alignment (BUG-032) is good; white-through-E still present after 1.5.48.
- **Update 2026-10-21 (17:35):** 1.5.49 did not clear it; white stub still under E.

### BUG-030 — Rapid rescans flatten mail to text via dirty attachment churn

- **Severity:** S2
- **Area:** AES / VBA
- **Status:** fixed (pending VBA re-import of `MSCANModule1`)
- **Reported:** 2026-10-21
- **Summary:** Rescanning "Everything you need to know about Nvidia's RTX Spark" (and similar IMAP items) hit `CommitMailHtml` "message has been changed". `EmbedAesFooterMark` had deleted and re-added every `aes_mark_*` attachment on each apply, dirtying the item before Save. Failed Saves left unsaved half-written HTML on the reading-pane item; Outlook then shows that as text-only / "Converted from text/plain". Logs at 16:01–16:03: conflict storm, footer count 2, "unsaved AES changes stayed on the message".
- **Repro:** Short-scan an IMAP mail, then rescan it again within a few seconds while it stays selected.
- **Workaround:** Click another message without saving the damaged one, then rescan once. If the body is already plain, restore from webmail or re-download.
- **Fixed in:** 1.5.46. `EmbedAesFooterMark` skips delete/re-add when the mark and letter Content-IDs are already present. `ReplaceAesFooterInMail` refuses to write when the item is still dirty after `DiscardUnsavedChanges`. Re-import `MSCANModule1`.

### BUG-029 — White hairlines to the right of footer A and E letters

- **Severity:** S4
- **Area:** AES
- **Status:** fixed (pending rescan)
- **Reported:** 2026-10-21
- **Summary:** Outlook drew a 2px pure-white vertical sliver between the A/E letter cells and the summary cells. The S row was fine (its neighbour is a slate spacer table). Screenshot pixels at the join were `#ffffff`.
- **Repro:** Rescan a message and look at the A/E/S column.
- **Workaround:** None.
- **Fixed in:** 1.5.44. Letter cells get a 2px slate `border-right`; summary cells get a matching `border-left`. The summary nested table also sets the strip background.

### BUG-028 — Footer top summary row is about 5px taller than the other rows

- **Severity:** S4
- **Area:** AES
- **Status:** fixed (pending rescan)
- **Reported:** 2026-10-21
- **Summary:** After the compact strip work, the top summary row (title + risk chips + quip) still reads about 5px taller than the middle and bottom rows.
- **Repro:** Rescan a message and compare the three strip rows in Outlook.
- **Workaround:** None.
- **Fixed in:** 1.5.40. Top-row centre padding and quip side-cell padding go from 2px to 0 (line-height stays 16px).
- **Update 2026-10-21 (17:40):** text still sat high in each row because the mark `rowspan` stretched every content row to the tallest (chips) and Outlook top-aligned the nested tables. 1.5.41 puts the mark in a side column instead. Copyright uses `v.` before the version.
- **Update 2026-10-21 (18:00):** the side column made it worse in Word (an empty paragraph after the nested rows table, then centring, so row 1 sat ~6px low). Root cause in both layouts: the mark cell was taller than the rows, and Word gives the extra to row 1 (rowspan) or splits it around the rows (nested). 1.5.42 returns to rowspan with mark padding 3px above / 0 below. Word render: rows 20.7/20.7/21.3, strip 63.3px.

### BUG-027 — Blocked BB quick-action chip rendered wider than SL/BA/TS/NT

- **Severity:** S4
- **Area:** AES
- **Status:** fixed (pending rescan)
- **Reported:** 2026-10-21
- **Summary:** With beacons blocked, the red BB chip looked wider than the white chips. Screenshot measure: white chips 77px, BB 79px. The blocked chip used a border the same colour as its fill, so the fill read full-width while white chips had a dark inset border; letter-spacing also let Outlook widen "BB".
- **Repro:** Block beacons on a message and compare BB with BA/TS in the footer.
- **Workaround:** None.
- **Fixed in:** 1.5.39. Every chip uses the same dark border, fixed 78px cell, and no letter-spacing. Copyright line also shows the suite version.

### BUG-026 — Footer-mark follow-up job flattened a mail to plain text

- **Severity:** S2
- **Area:** AES
- **Status:** fixed (1.5.34; no VBA re-import needed for the fix itself)
- **Reported:** 2026-10-21
- **Summary:** In 1.5.32/1.5.33 the out-of-process `_embed_footer_mark_job` attached the A/E/S letter images, then rewrote `HTMLBody` and saved the mail. On "Agency Isn't Authorization" (julian.garrett@aliniant.com, a non-Exchange store) the job saved at 14:58:20. The user's second rescan then hit "the message has been changed" at 14:58:47, left unsaved AES changes, and the item ended up `BodyFormat=plain` with no footer. Before 1.5.32 the job always stopped at "mark already attached", so it never wrote to the mail.
- **Repro:** On 1.5.33, before re-importing `MSCANModule1`: Short Scan a mail, then rescan it within about 30 seconds.
- **Workaround:** None for the damaged item locally. Check the copy in webmail or on the server, or re-download it.
- **Fixed in:** 1.5.34. The job attaches only the mark again, and only when VBA has not embedded it. Letters are left to VBA `EmbedAesInlineImages` in the same commit as the footer.

### BUG-025 — Footer strip taller than its text; A/E/S could not sit close to the row edges

- **Severity:** S4
- **Area:** AES / VBA
- **Status:** fixed (pending VBA re-import of `MSCANModule1` + rescan)
- **Reported:** 2026-10-21
- **Summary:** The text-based A/E/S letters picked up Outlook's minimum line spacing. Making them bigger made the rows taller, and they could never get within a few px of the row edges. The strip also had a minimum height of about 84px from the 72px mark plus 12px padding, even though the text needs far less.
- **Repro:** Scan a message, then compare the footer strip height with its three text rows.
- **Workaround:** None.
- **Fixed in:** 1.5.32. The letters are now 12px images (`assets/icons/aes_mark_letter_*.png`, from `scripts/build_footer_strip_images.py`) with 4px padding in 20px rows. The rows have tighter padding (16px text plus 2px; chips 18px plus 1px), and the mark is 59x56, so the strip is 62px. VBA `EmbedAesInlineImages` attaches the letters from `AES-Inline-Img` markers. The Python follow-up job also attaches any that are missing, so they appear before the re-import.
- **Update 2026-10-21:** before the VBA re-import, the letters showed as broken-image boxes. The follow-up job read the mail before Outlook had written the new footer, saw the old footer (mark cid, no letters) and logged "already attached". 1.5.33 makes the job wait for a body that contains `cid:aesmarkletter`. Re-importing `MSCANModule1` avoids the race entirely.
- **Update 2026-10-21 (15:06):** after the re-import, VBA attached all three letters but their Content-IDs were garbage (for example `\u8d9c\u4c86...`). `EmbedAesInlineImages` passed a `String` variable to the late-bound `PropertyAccessor.SetProperty`, which received it ByRef and stored the reference bytes. The mark works because it passes a `Const`. 1.5.35 passes `CStr(cid)`. Re-import `MSCANModule1` again.
- **Update 2026-10-21 (16:40):** the letters rendered, but they sat high in each row with a gap below. A render through Word's own HTML engine showed why: the letter cell's 4px padding plus the 12px image made each row 24–27px instead of 21–24px, so the extra height collected under the letter. 1.5.36 removes the letter cell padding and the fixed line height and makes the letters 14px, centred in the row the text sets. In Word that gives about 4px above and below, and the strip is no taller than it is without letters (Python only; rescan).

### BUG-024 — Footer summary rows are not centred on the strip

- **Severity:** S4
- **Area:** AES
- **Status:** fixed (rescan to refresh footers)
- **Reported:** 2026-10-21
- **Summary:** In Outlook the top summary line sat left of the strip centre and the sender line sat right of it. The sender line was centred only on the space right of the logo and A/E/S column. The top line relied on an empty spacer table to balance the quip on the right, and Outlook did not keep that table at its width.
- **Repro:** Scan any message and compare the two summary lines with the centre of the strip. The quick-actions row is centred on its middle button, which is intended.
- **Workaround:** None.
- **Fixed in:** 1.5.16. Both rows now balance the side blocks with td padding: the top line is padded left by the quip width less the logo column, and the sender line is padded right by the logo column.

### BUG-023 — Footer strip shows banding after re-processing; A/E/S mark letters

- **Severity:** S3
- **Area:** AES
- **Status:** fixed (rescan to refresh footers)
- **Reported:** 2026-10-21
- **Summary:** After footer re-processing, the slate strip showed vertical smear / noise. Same-colour "ghost" balancing text (`left_top` / `left_copy`) was being painted by Outlook's Word renderer, and `aes_mark_footer.png` had non-solid background pixels. Operator also wanted A, E, S in Montserrat Bold beside the orange A, on the three content rows (just off the strip grey).
- **Fixed in:** 1.5.14. Ghost text replaced with empty spacers; mark PNG background solidified to `#46525c`; A/E/S column added (`font-family: Montserrat`, weight 700, colour `#55616b`).
- **Update 2026-10-21:** in Outlook the 1.5.14 letters came out black and sat left of the rules. 1.5.15 puts the colour on a `<font>` tag as well as the style, starts each rule at the letter column's left edge, and sets the letters to 17px bold `#5a6671`.

### BUG-022 — GURI window jumps to the front while other apps are in use

- **Severity:** S3
- **Area:** GURI
- **Status:** fixed (restart GURI to pick up)
- **Reported:** 2026-10-21
- **Summary:** The main window stole focus from Outlook and other apps. A second start of `guri\gui.py` without `--raise` always sent `RAISE` over the single-instance socket, and `expand_to_screen` / `_force_window_foreground` called `ShowWindow(SW_SHOW)`, `BringWindowToTop` and `SetForegroundWindow` even when the window was already visible.
- **Repro:** Leave GURI open in the background, work in Outlook; a launcher, script or second `gui.py` start without `--raise` pulls GURI forward.
- **Fixed in:** 1.5.13. Focus is taken only for tray Open, ribbon `--raise` / Aura, or the first process start. Bare second starts send `PING` and leave focus alone. `ShowWindow(SW_SHOW)` is no longer called on an already-visible window.
- **Update 2026-10-21:** still reported at 14:59. Both running `guri\gui.py` processes started at 09:26, before the 12:49 fix, so they are running the old code. A code check found no other focus-taking path: only the ribbon GURI/Aura buttons send `--raise`. GURI still needs a restart.

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
