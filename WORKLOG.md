# GeoFooter / AES / GURI — Work log

Append an entry **whenever meaningful code or project-doc changes are made**. Newest entries at the top.

## Format

```
### YYYY-MM-DD — short title
- What changed (files / modules)
- Why (user goal or bug)
- Follow-up (re-import VBA, restart Outlook, etc.) if any
```

---

### 2026-09-24 — Add docs/bug_tracker.md
- New living bug list under `docs\` (open/closed, severity, areas, template).
- Seeded BUG-001 (AccessVBOM missing from Trust Center UI) and BUG-002 (Outlook VBA still needs re-import after 1.2.1 package move).
- Follow-up: keep updating this file when bugs open/close; finish VBA auto-import script docs when that lands.

### 2026-09-24 — Reorganise Python into aes/ guri/ aura/ geofooter/ packages (1.2.1)
- **Moves:**
  - Root and `VBA\` Python → `aes/` (engine, checks, dialogs, `action_handler`, `secret_store`) and `aes/scanners/`.
  - `guri*.py` → `guri/` (`core`, `gui`, `gui_service`, …).
  - `broker_removal/` → `aura/`.
  - `geofooter_paths` / `version` / `aes_crashlog` → `geofooter/`.
- **Duplicates removed:** the root copies of the scanners and `aes_secrets` are gone; the `VBA\` copies were canonical, so the newer `link_scanner` now loads. `VBA\icons` and `VBA\pages` were merged into `assets\`.
- **Imports:** all imports are now absolute (`from guri.core import …`). Each entry script has a `sys.path` bootstrap because it is launched by file path.
- **GURI data:** recipients, important domains and email status now read from `datastore\`. The user's lists were migrated there after the previous tidy left GURI looking in the wrong folder.
- **Callers updated:**
  - VBA: `MSCANPaths` (`GetAesScript`, new script paths, `aes\` marker), Module1, Settings, Toolbar, ClassificationDialog, Diagnostics, ModSenderRules (blocked pixel via install root), AppBootstrap, ReadNotify comment.
  - Ribbon host `Connect.cs` (`guri\gui.py`; working directory is the install root).
  - `scripts\*.bat/.ps1`, `tests`, `diagnostics_dialog`, `geofooter.paths` discovery.
- **Verified:** every file compiles; every module imports; the geolocate optional modules are non-None; the entry scripts run by path from `VBA\` as working directory.
- **Follow-up:**
  - Re-import VBA: MSCANPaths, MSCANModule1, MSCANSettings, MSCANToolbar, MSCANClassificationDialog, MSCANDiagnostics, MSCANModSenderRules, MSCANAppBootstrap, MSCANReadNotify.
  - Re-run `AesRibbonHost\install.ps1` with Outlook closed.
  - Re-register `aes://` with `aes\action_handler.py --register`.

### 2026-09-24 — Tidy root: scripts/tests/docs/assets; drop Old/_archive
- Moved migrators/launchers → `scripts/`, ad-hoc tests → `tests/`, feature docs → `docs/`, brand/icons/forms/data → `assets/`.
- Removed tracked `Old/` and `_archive/` from the repo (gitignored going forward).
- Removed stale root `aes_classify_dialog.py` / `aes_diagnostics_dialog.py` (canonical under `VBA\`).
- Follow-up: none for layout; VBA re-import still needed for `MSCANPaths` if not done yet.

### 2026-09-24 — Move suite icons to assets/icons
- Corrected accidental `assets/icons_tmp` path from prior commit; icons live at `assets/icons/` (`geofooter_paths.icons_dir`).

### 2026-09-24 — AesRibbonHost verify: no false CoCreate FAIL
- `verify32.ps1` / `install.ps1`: success is CLSID + InprocServer32 present; CoCreate while Outlook holds the DLL is reported as skipped, not FAIL.
- Explicit CLSID fallback now uses the built DLL’s FileVersion (was hard-coded 1.0.0.0).
- Follow-up: none required if registry OK and Outlook already loads the add-in.

### 2026-09-24 — Configurable install root (relative paths)
- Added `geofooter_paths.py` + `VBA\MSCANPaths.bas`; GURI Database tab can set Install root.
- Pointer: `%LOCALAPPDATA%\GeoFooter\install_root.txt` (optional `GEOFOOTER_ROOT`).
- Suite dirs relative to root: `assets/`, `datastore/`, `debuglog/`, `crashlogs/`.
- Removed hard-coded `C:\GeoFooter` from live VBA/Python/ribbon host where practical.
- Follow-up: re-import `MSCANPaths.bas` first, then Module1/Toolbar/Settings; Save root once in GURI.

### 2026-09-24 — Correct suite version to 1.2.0 (patch-first SemVer)
- Suite is **1.2.0**, not 2.0.0 — SemVer MAJOR.MINOR.PATCH with ~90% patches.
- Commit messages must open with a caption of what was done (`AGENTS.md`).
- Retagged GitHub release `v1.2.0`; removed mistaken `v2.0.0`.

### 2026-09-24 — Suite versioning + public GitHub
- SemVer source of truth: `VERSION`, `version.py`, `CHANGELOG.md`, AesRibbonHost assembly version.
- Hardened `.gitignore` (DB configs, secrets, build junk). Published clean tree to https://github.com/youcangetjules/geofooter (no live credentials).
- Follow-up: rotate Postgres/MySQL passwords that existed in older local commits.

### 2026-09-24 — Create GURI must persist to database
- Root cause: `guri_records.avg_risk` was `VARCHAR(50)` but Create-GURI stores document file paths there — inserts failed silently while the UI still said success.
- Widened `avg_risk` to TEXT (Postgres migration on connect); `generate_guri` / `insert_guri` now raise on save failure.
- Accept common date formats (`YYYY/MM/DD …`); document-type radios share one exclusive group.
- Restart GURI GUI and Generate again to confirm the row appears under View Records.

### 2026-09-24 — VT multi-engine RiskTable scoring
- VirusTotal no longer treats a lone engine hit (e.g. 1/91 on Substack) as suspicious.
- Consensus thresholds: absolute count + % of engines; weak hits stay clean with “below RiskTable threshold” in the summary.
- Cached VT rows are re-scored on read so old 1/N verdicts do not stick.
- Files: `VBA\aes_threat_intel.py`. Rescan to refresh.

### 2026-09-24 — Threat intel OK hint
- All-OK Threat intelligence line now ends with: “Run Full or Deep scan to see these in full.”
- File: `VBA\geolocate_headers.py`. Rescan to refresh the footer.

### 2026-09-24 — RiskTable banner wording
- Banner note is now `{n} found on RiskTable` (was `{n} links disabled - found on RiskTable`).
- File: `VBA\geolocate_headers.py`. Rescan an email to refresh the strip.

### 2026-09-24 — Aura ribbon button (data-broker removal)
- New Home ribbon / Add-ins toolbar **Aura** (Aliniant Universal Removal Application).
- Opens GURI on the Aura tab (`guri_gui.py --raise --aura`); IPC `AURA` selects the tab on an existing instance.
- Icon `aes_aura.*` via `make_aes_icons.py`; VBA `ShowAuraGui` / tag `AES_AURA`.
- Follow-up: fully quit Outlook (tray too) → `AesRibbonHost\install.ps1` (DLL was locked mid-build) → re-import `MSCANToolbar.bas` → restart Outlook.

### 2026-09-24 — Clickable Attachments / Beacons / Links breakdowns
- Summary strip metrics are underlined links opening per-item HTML reports.
- New `aes://open-report?path=…` (Outlook-safe; `file://` is often blocked) in `aes_action_handler.py`.
- Beacons: None still gets a report page; Links summary now links to the links report.
- Follow-up: Short Scan to refresh the footer. `aes://` must stay registered (`python aes_action_handler.py --register` if clicks do nothing).

### 2026-09-24 — Rescan restore actually works (IMAP save conflicts)
- Root cause: rescan hit `#-2147221239 message has been changed` after banner `Attachments.Add`; IMAP also strips `<!-- AES -->` comments so replace fell into a failing append path — HTML never committed / looked like plain text.
- `CommitMailHtml`: HTMLBody+Save with re-resolve retries; re-resolve after banner embed before write.
- Footer bounds: durable `aes-footer-start/end` anchors + “Aliniant AES Scan Result” (not comments only).
- `TryRestoreMitigatedHtml`: detect mitigated body by content / `restore-html?id=` / subject; log skips.
- Never force `BodyFormat=HTML` on empty HTMLBody (that flattens Gmail/IMAP to text).
- Follow-up: **re-import `MSCANModule1.bas`**, close the reading pane on the mail, Short Scan again. Mitigated mails need a matching file under `%LOCALAPPDATA%\GeoFooter\mitigated_html\`.

### 2026-09-24 — Stop converting trusted mail to plain text
- Cause: score >70 mitigation (`AES-Mitigate: text-strip`) flattened HTML; false-positive link scores triggered it; rescans left mail as text.
- Trusted senders never get text-strip (Python marker + VBA `IsSenderTrusted` guard).
- Before a normal HTML footer apply, `TryRestoreMitigatedHtml` restores the saved backup when BodyFormat is still plain.
- Follow-up: **re-import** `MSCANModule1.bas`, then Short/Full rescan the affected mail (or use Restore original HTML if the backup id is in the plain body).

### 2026-09-24 — Status strip: “Links appear authentic”
- Banner note wording: `Links are authentic…` → `Links appear authentic - not found on RiskTable`.
- Follow-up: rescan to refresh the status strip.

### 2026-09-24 — Decode public hop IPs properly (allocation > BGP)
- Bug: RDAP/ipinfo BGP ASN (e.g. AS14618 Amazon) overwrote Sailthru allocation; Country/City/PTR stayed Unknown.
- RDAP now prefers registrant org + address (Sailthru, INC / Nashville US), allocation prefix (`192.64.236.0/22`), netname SAILTHRU; hostname map → AS18877.
- BGP transit kept as a decode note (`BGP transit AS14618 Amazon…`), not as Organisation.
- Fill registration **before** geo enrich; PTR falls back to Received FQDN; no Sailthru ASN on RFC1918 internal hops.
- Follow-up: rescan — exit hop should show AS18877 / Sailthru / US / mta237-62.sailthru.com.

### 2026-09-24 — Clean TI is not suspicious (scoring + UI)
- Root cause: newsletter ESP/Google redirects and real `google.com` hosts were scored medium/high; TI “OK” never cleared them → +122 High-risk elements on trusted mail.
- `link_scanner`: sender/ESP/official-brand redirects stay **low**; official brand domains are not lookalikes of themselves.
- TI annotation: clean URL/domain **demotes** heuristic medium/high → low before scoring.
- TI table: collapse OK wall to “All N indicators OK” (or “N OK hidden”); only show flagged/unknown rows.
- Follow-up: Short/Full rescan of the Telecoms newsletter — High-risk elements should drop sharply.

### 2026-09-24 — Threat intel: clean is not "suspicious"
- Verdict badge **clean** now displays as **OK** (not a threat tier next to suspicious).
- VirusTotal clean summary: `not flagged (N engines)` instead of `0 malicious, 0 suspicious…`.
- OTX below-threshold pulses: `weak community mention (below threshold)` instead of `in N threat pulse`.
- Follow-up: Short/Full/Deep rescan to refresh the Threat intelligence table.

### 2026-09-24 — Data broker removal (Incogni-style) v1
- New `broker_removal/` package: catalog (~40 brokers), local identity profile, SQLite request store, GDPR/CCPA/generic templates, AES pending queue.
- GURI **Data Brokers** tab: Identity / Brokers / Requests / AES inbox (Accept → draft; templates copy/export only — no auto-send).
- AES: catalog domain match adds a Data broker risk factor; footer **Queue data removal** → `aes://broker-removal` → `broker_pending_from_aes.json`.
- Follow-up: open GURI → Data Brokers; rescan broker mail to get the new footer button. No VBA re-import.

### 2026-09-23 — Diagnostics Live Controls + Outlook freeze relief
- Freeze: CatchUp no longer runs on every 2.5s queue tick (only heartbeat/quiet); skip CatchUp while queue/in-flight busy; VBS ItemLoad nudges cut 6→2 (queue) and 12→4 (geo job); diagnostics dialog launches **non-blocking** so Outlook stays responsive.
- `aes_diagnostics_dialog.py`: **Live Controls** — enact Queue tick / Catch-up / Heartbeat / Post-send; Info/Warning/Audit/Debug + Apply; Start/Stop **capture** (sidecar under Logs), live log tail.
- VBA: `aes_diag_commands.json` drained on every ItemLoad nudge; `MSCANModLogging` Save/Reload + capture APIs.
- Follow-up: **re-import** `MSCANModQueueManager`, `MSCANModLogging`, `MSCANDiagnostics`, `MSCANModule1` then Compile.

### 2026-09-23 — Auto-scan broken: NameError crashing every compact job
- Symptom: Outlook queued mail (ItemAdd/NewMailEx/CatchUp) but every Python job wrote `.fail` — no footers, CatchUp kept re-queueing.
- Cause: BGP prefix backfill called `HeaderAnalyzer._blank_reg` (class does not exist); `NameError` aborted report generation.
- Fix: use `self._blank_reg` on `HTMLReportGenerator`.
- Follow-up: next incoming mail (or CatchUp within ~90s) should apply footers again. No VBA re-import needed.

### 2026-09-23 — Trust button label: "Sender Trusted" (green)
- `VBA\geolocate_headers.py` / `aes_action_handler.py`: when the sender is on the trusted list, the footer button is solid green **"Sender Trusted ✓"** (was "Sender TRUSTED"). Live click refresh uses the same wording.
- Follow-up: click Trust on an open mail (or Short Scan) to refresh an old footer.

### 2026-09-23 — Route table: full IP registration (no more Prefix=NA)
- Cause: RDAP/`ipwhois` often returns the literal **"NA"** for Microsoft 365 IPv6 (`asn_cidr`), and that was stored as Prefix; compact also skipped geo/WHOIS on shared M365 hops; exit PTR only ran on deep.
- `VBA\geolocate_headers.py`: treat NA/Unknown as blank; RDAP → **RIPEstat network-info** → bounded WHOIS (8s) via `_fill_ip_registration`; exit hop always force-WHOIS + PTR; full/deep fill every thin public hop; BGP prefix back-filled after route checks; WHOIS only when prefix/ASN still empty (or exit), so shared relays stay fast.
- Follow-up: re-run Full/Deep (or Short) on the mail — Prefix should show real CIDRs (e.g. `2603:10b6::/32`), LinkedIn exit country/org from WHOIS/geo when available.

- `VBA\geolocate_headers.py` / `aes_action_handler.py`: footer action is always "Trust sender" (or "Sender TRUSTED" when already trusted).
- Follow-up: Short Scan to refresh existing footers.

### 2026-09-23 — Trust sender now clears beacon blocks (and refreshes footer)
- Bug: Trust removed only the exact `who` string from block lists. A domain-level (or mismatched) beacon entry could survive while attachments cleared; the footer is also baked at scan time so Beacons BLOCKED stayed red until rescan.
- `aes_action_handler.py`: Trust clears **email and domain** from both `block_attachments` and `block_beacons`; after save, rewrites the open/selected mail's AES action buttons (Beacons/Attachments back to outline, Trust → green TRUSTED).
- `VBA\geolocate_headers.py`: `_sender_block_state` — trusted wins (beacons/attachments forced off for footer); banner “B Blocked” ignores beacon hits when trusted.
- `VBA\aes_settings_panels.py`: same identity purge when trusting from Settings.
- Follow-up: click Trust again on the stuck mail (or Short Scan after Trust). No VBA re-import required for the handler; next scan picks up Python footer logic automatically.

### 2026-09-23 — Auto-scan catch-up + richer diagnostics
- Root cause: auto-scan works when `ItemAdd`/`NewMailEx` fire (log shows successful queue → ProcessQueue through 21:25), but some deliveries never raise those events; catch-up only ran after startup quiet and only looked at the 2 newest items / 20 min / 5 min if already read — so misses (e.g. LinkedIn invite scanned manually at 21:35) stayed unscanned.
- `VBA\MSCANModQueueManager.bas`: stronger `CatchUpMissedInboxMail` (15 newest/inbox, 90 min lookback, unread ≤120 min / read ≤20 min, skip reasons logged); periodic 90s catch-up heartbeat VBS (`aes_catchup_heartbeat.vbs`); catch-up also on due queue ticks; `GetAutoScanDiagnostics` / `GetQueueSize` / ItemAdd+NewMailEx timestamps.
- `VBA\MSCANEmailWatcher.cls` / `MSCANEventHandlers.bas`: richer NewMailEx skip logging (account OFF, non-scannable, GetItemFromID fail).
- Diagnostics: `MSCANDiagnostics.bas` embeds auto-scan pipeline + last ~40 auto-scan log lines (tail-only read); `aes_diagnostics_dialog.py` adds Auto-scan pipeline / Account SCAN-SKIP / Recent log checks.
- Follow-up: **re-import** `MSCANModQueueManager.bas`, `MSCANEmailWatcher.cls`, `MSCANEventHandlers.bas`, `MSCANDiagnostics.bas` (see `VBA\IMPORT.txt`); reopen Diagnostics → Run all.

### 2026-09-23 — Route maps plot every server; tromboning + BGP origin checks
- `VBA\geolocate_headers.py`: new `_route_map_points` feeds both deep-scan maps (Leaflet + offline SVG): each hop's sender **and** its receiving server (so the final mailbox, e.g. Dublin DU0, now appears), merged within 25 km, numbered stops, green origin / red relays / blue delivery, dashed marker border for approximate locations, popups name the decoded server role. `fitBounds` capped at zoom 7. Added `import math`.
- New `VBA\aes_route_checks.py`: (1) mail-path tromboning — path goes ≥3,000 km away and returns within 1,500 km of an earlier stop (+5, flagged on the far hop + journey); (2) BGP via RIPEstat (no key): announcing prefix/origin AS, RPKI status (invalid → +20 "possible hijack"), MOAS (+5), expected-vs-announced ASN (informational), origin's top upstreams → BGP Peers column (previously always empty). Compact: exit hop only, no upstream lookup; full/deep: up to 6 public hops. Scored under "Routing / BGP", capped at 25.
- Why: user asked whether we detect tromboning / hijack signals and whether deep-scan maps show every hop.
- Follow-up: rescan to verify; RIPEstat must be reachable (fails quietly if not).

### 2026-09-23 — Geo fallback fix: approximate IP-database city for exit hop
- `VBA\geolocate_headers.py`: `_try_ipinfo` no longer marks success on error / rate-limit / empty replies (that suppressed the ip-api geolocation fallback, leaving e.g. Sailthru 163.47.180.68 as Unknown instead of ~Nashville). ip-api fills a missing city when ipinfo gave only a country. New `GeoLocationResult.geo_source`.
- `_enrich_public_hop`: IP-database locations are tagged `location_source = "IP geolocation via … — approximate"` and `geo_confidence = low`; journey shows "(approx.)". A high/medium Microsoft-style hostname location now overrides an approximate IP-database city even within the same country.
- Follow-up: rescan to verify.

### 2026-09-23 — Route Journey narrative + EPF edge hosts
- `VBA\geolocate_headers.py`: new "Route Journey" block under the route table (`_build_route_journey_html`): one plain-English line per hop (who hands to whom, decoded role + city), flags waits ≥10s, backwards clocks, and country/continent crossings; closing line with final mailbox server, hop count and end-to-end time.
- `VBA\aes_addr_decode.py`: `decode_m365_host` handles `BL02EPF0002992C` (EPF = Edge Protection front end, 1–2 char DC number); tenant-MX rule now only applies when the first label is not a server name (EPF hosts live under `mail.protection.outlook.com`).
- Follow-up: rescan to verify.

### 2026-09-23 — Route table: full Microsoft 365 hostname breakdown (from + by)
- `VBA\aes_addr_decode.py`: new `decode_m365_host` / `describe_m365_host` — parses datacenter + DC number (AM7 → Amsterdam DC 7), server role + serial, environment, forest/pod label (EURP250, EURPRD08, eop-nam12, GBRP265…), region. New forms: `…PPF<hex>` (e.g. AM7PPF30B196146 — previously undecoded), EOP filtering nodes (`BN8NAM12FT012`), outbound EOP relays (`mail-am6eur05on2078`), tenant MX (`contoso-com.mail.protection.outlook.com`). `decode_host_location` uses it, so PPF hosts now fill Country/City. Added `2603:10a6::/32` (Exchange Online Europe forests).
- `VBA\geolocate_headers.py`: route-table note now decodes both hosts ("from: … → by: …"); dropped the duplicated raw "by HOST" note (already shown in the host cell); `NA`/`N/A` treated as blank so hostname org/location hints can fill them.
- Why: hop rows like `AM7PPF… by DU0P250MB0793…` showed only "Europe Pod 250" and Unknown location.
- Follow-up: rescan to verify (shell unavailable for tests).

### 2026-09-23 — GURI "AES Scores" tab: view / set / reset sender threat scores
- New `aes_score_history.py` (root): reads/edits `%LOCALAPPDATA%\GeoFooter\aes_sender_history.json` using the scanner's `.lock` protocol; set score, reset score (drop `score_ema`), forget sender, list senders at a domain, clear `threat_intel_cache.json` entries for a domain (+ subdomains / URLs on it).
- `guri_gui.py`: new "AES Scores" tab (after Inbox Search): filterable, numerically sortable table (sender, domain, stored score colour-banded, emails, first/last seen) with Set score… (also double-click), Reset score, Reset domain… (scores + threat-intel cache), Forget sender….
- Manually reset `reply@foreignpolicy.com` stored score 62 → 6 (inflated by the pre-fix Deep Scan of tracking links).
- Why: trialling/tuning needs a way to clear score memory after scoring bugs.
- Follow-up: restart GURI; not yet run-tested (shell unavailable).

### 2026-09-23 — Links: newsletter click-tracking to sender's own site is low risk
- `VBA\link_scanner.py` (+ root copy) — decodes redirect destinations (base64 path segments like Sailthru `/click/.../<b64>/...`, or `url=`/`u=`/`redirect=` params). Redirect is **low** ("Tracking redirect to sender's own site") only when the tracking host is on the sender's domain or a known ESP (Sailthru, Mailchimp, SendGrid…) **and** the destination is on the sender's domain. Unrelated destination → medium "Redirects to X (not the sender's domain)"; undecodable → medium as before. New `destination` field on findings
- `VBA\geolocate_headers.py` — threat intel also checks each redirect's decoded destination domain
- Why: Foreign Policy newsletter links (`link.foreignpolicy.com/click/...`) were all "medium", inflating risk score and HRE
- Follow-up: tests in agent store `test_link_redirects.py` — could not run (terminal hung); re-scan the FP mail to confirm links show low

### 2026-09-23 — Deep Scan: threat intelligence shown once
- `VBA\geolocate_headers.py` — the deep report showed the threat-intel table twice (appended under the route table and again in its own card). `_build_full_details_html(include_threat_intel=False)` for the deep report; full-scan footers/reports unchanged
- Follow-up: run a new Deep Scan

### 2026-09-23 — Report route map: replace blocked OSM tiles
- `VBA\geolocate_headers.py` — Leaflet map used `tile.openstreetmap.org`, which returns 403 "Access blocked" for `file://` reports (no Referer, per the OSM tile policy). Now uses CARTO Voyager tiles, falling back to Esri World Street Map if CARTO fails
- Follow-up: re-scan (or re-open a newly generated report); existing reports on disk keep the old tile URL

### 2026-09-23 — Text-only mitigation: all AES actions as plain links + restore HTML
- High-risk (>70) text-only mails now include the same footer actions as plain links (Show links, Block attachments/beacons, Trust / Not trusted, Restore original HTML)
- `VBA\geolocate_headers.py` — shared action specs; plain-text block in `AES-Plain-Summary`; RiskTable banner note when links checked
- `VBA\MSCANModule1.bas` — saves HTML backup under `mitigated_html\`, fills `aes://restore-html?id=…`, defangs RiskTable hrefs on HTML footers
- `aes_action_handler.py` — handles `restore-html` via Outlook COM
- Follow-up: re-import `MSCANModule1.bas`; ensure `aes://` protocol still registered (`python aes_action_handler.py --register`)

### 2026-09-23 — Compact scan: count Links (body was never exported)
- Root cause: compact auto-scan deferred body/beacon export (`Links: 0` / `Beacon: 0` even when the message had many hrefs)
- `VBA\MSCANModule1.bas` — compact now writes a capped (~400KB) body sidecar + beacon count from that file (no neutralize on auto-scan); full/deep still export full body
- `VBA\link_scanner.py` — also match unquoted `href=`
- Follow-up: re-import `MSCANModule1.bas` into Outlook (`VBA\IMPORT.txt`), then rescan — banner should show real Links / Beacon counts

### 2026-09-23 — Route Table: always WHOIS sending exit (Sailthru AS18877)
- `VBA\aes_addr_decode.py` — Sailthru `163.47.180.0/23` → AS18877 / Sailthru; `*.sailthru.com` host map
- `VBA\geolocate_headers.py` — compact geo budget no longer spent on Microsoft/Google inbound hops; first non-shared public hop (the sending exit, e.g. 163.47.180.68) is always geo+RDAP+WHOIS'd and tagged `[sending exit]`
- Follow-up: re-scan — that exit row should show AS18877 / Sailthru, Prefix 163.47.180.0/23, RIR ARIN

### 2026-09-23 — Route Table: RDAP for Prefix / RIR / ASN on every public hop
- `VBA\geolocate_headers.py` — RDAP (`ipwhois` depth=1) now runs for every public hop in compact as well as full/deep, filling ASN, organisation, Prefix (asn_cidr) and RIR; process-level cache avoids repeat lookups. Compact still skips reverse-DNS and the secondary AbuseIPDB call; geo stays capped at the first few public hops
- Follow-up: re-scan — Prefix/RIR should populate (e.g. ARIN + a CIDR for Microsoft hops)

### 2026-09-23 — Route Table: origin-first order + offline ASNs
- `VBA\geolocate_headers.py` — Route Table / maps now list hops oldest→newest (hop 1 = origin, last hop = delivery); delays recalculated for that order
- `VBA\aes_time_checks.py` — ordering / future-stamp / stamping-server checks updated for oldest-first
- `VBA\aes_addr_decode.py` — provider nets carry ASN (AS8075 Microsoft, AS15169 Google, AS16509 Amazon SES); hostname suffixes fill ASN when geo/RDAP leave it Unknown or Skipped (compact)
- Follow-up: re-scan to see reversed Route Table

### 2026-09-23 — Route Table: wider IP column, faint cell borders, military Created DTG
- `VBA\geolocate_headers.py` — IP column ~16% (was 9%); faint `#e4eaee` borders on every Route Table cell; Created stamp is now military Zulu `YYYYMMDDHHmmz` (e.g. `202609231812z`) in the Route Table footer, summary line 2, and sidecar reports; Sent on line 2 / Sent (local) also uses the same DTG (sender wall-clock kept as a grey note in the details row)
- Follow-up: re-scan to see the new table layout

### 2026-09-23 — Settings: dropdown chevron + urlscan row alignment
- `VBA\aes_settings_dialog.py` — all QComboBoxes use a flat chevron (`VBA\icons\chevron_down.svg`) instead of the native drop-down button
- `VBA\aes_settings_panels.py` — urlscan visibility row moved into the key grid so the dropdown lines up with the key fields
- Follow-up: reopen AES Settings

### 2026-09-23 — Threat intel wired into every scan
- `VBA\geolocate_headers.py` — `_run_threat_intel()` runs after link extraction for compact / full / deep: origin IP + up to 3 non-provider public hop IPs, sender + header domains (From/Reply-To/Return-Path/DKIM d=/Message-ID), link domains and URLs (compact: List-Unsubscribe etc. from headers; full/deep: body links), attachment SHA-256s. Findings scored under "Threat intel" (cap +60); flagged links get provider reasons and are raised to high/medium; Route Table Reputation cell lists provider verdicts and a malicious hop turns red; new indicator × provider table under the Route Table (full scan) and a "Threat intelligence" card (deep report)
- Time budgets: compact 8s, full 20s, deep 75s (urlscan submissions only in deep, using the Settings visibility)
- Why: user added keys for all providers and wants links / IPs checked against them
- Follow-up: NOT yet run — agent terminal wedged; run one Full Scan to verify

### 2026-09-23 — Settings: threat-intel API keys + trusted-sender table
- NEW `VBA\aes_settings_panels.py` — External APIs tab gains a key row per provider (Shodan, VirusTotal, Google Safe Browsing, GreyNoise, AlienVault OTX, abuse.ch, urlscan.io): paste → "Save & test" (DPAPI via `aes_secrets.set_secret`), Clear, live key probe; keyless InternetDB / domain-blocklist rows get a Test button; urlscan visibility (default Private) saved to `aes_api.json`
- Sender Status tab gains a table of trusted / untrusted senders (email count, first/last seen from sender history) with Move to Untrusted / Move to Trusted / Remove / Add as Trusted; writes `aes_sender_rules.json` immediately, same consistency rules as `aes_action_handler.py`
- `VBA\aes_settings_dialog.py` — mounts both panels; External APIs tab now scrolls
- Why: user wants all intel sources configurable, and a way to review/undo trust decisions
- Follow-up: NOT yet run/previewed — process creation stalled on the machine during testing; open AES Settings to verify

### 2026-09-23 — Threat intel module, timestamp checks, IP/host decoding
- NEW `VBA\aes_threat_intel.py` — IP / domain / URL / hash lookups across InternetDB (keyless), DNS blocklists DBL/SURBL/URIBL (keyless, queried at the lists' authoritative servers with a canary check), Shodan, VirusTotal (free-tier 4/min enforced across processes), Safe Browsing (batched), GreyNoise, OTX, URLhaus/ThreatFox/MalwareBazaar, urlscan (deep only). Per-mode time budget, disk cache `threat_intel_cache.json`, `score_reports()` → risk findings (cap +60). No uploads; hashes only. NOT yet called from the scan engine
- NEW `VBA\aes_time_checks.py` — Received ordering (older hop stamped later: +5, >1h +15), final hop in the future, hop UTC offset vs the stamping server's country (+10 decoded host / 2h+ gap, +3 GeoIP-only 1h), Date header after first hop (+10) or >72h before (+5), Date offset vs origin country (+10); cap +25; UTC and Google Pacific stamps exempt. Wired into `_prepare_display_data` as "Timestamps" risk category; flagged hops turn red with ⚠ notes in the Timestamp cell
- NEW `VBA\aes_addr_decode.py` wired into `_extract_hop_details`: bare IP literals (Gmail `by 2002:a05:…`) and Exchange `(2603:…)` parenthesised IPs now captured; 6to4/Teredo/NAT64/ULA/link-local/EUI-64/provider ranges decoded under the IP cell; 6to4-wrapped public IPv4 used for geo; hostname-derived location (M365 region-aware DC codes, AWS SES regions, Mimecast, IATA) fills Country/City and overrides provider-HQ registry geo (Redmond → Cardiff/London)
- `SecurityAssessor.apply_extra_factors()` — generic scored-findings hook
- FOUND: local DNS resolvers NXDOMAIN every blocklist query (even ZEN's 127.0.0.2 test entry), so the existing origin-IP DNSBL check has never listed anything. `_check_blocklists` now goes through `aes_threat_intel.dnsbl_lookup` (authoritative servers, canary-verified) and reports "Unavailable" instead of a false "Not Listed" when no path works (not yet run — see stall note above)
- Follow-up: none for VBA; Python applies on next scan

### 2026-09-23 — Risk Score Compilation: row colour by contribution
- `VBA\geolocate_headers.py` — each factor row is shaded by its points: 0 or negative = no background; linear fade from white to solid red (#d32f2f) at +25 and above; white text once the fill is dark. Colour is set per cell (`bgcolor` + CSS) because Outlook's Word renderer ignores `<tr>` backgrounds. Negative points now show as "−5" instead of "+-5"
- Why: make the biggest score drivers visible at a glance
- Follow-up: none (Python only); applies to new scans and re-scans

### 2026-09-22 — GURI Welcome stats / refresh look stuck
- `guri_gui.py` — Email Statistics now auto-updates from Outlook scrape cache (all AES / learning "my" addresses, not only Aliniant); Welcome rebuild stamps Last refresh; AES `--raise` / expand triggers UI refresh + scrape if stale (>2 min)
- Why: sidebar showed Last refresh 2026-08-25 while scrape cache was current (15:21); stats only ran on manual "Refresh Statistics" (10k DB rows)
- Follow-up: restart GURI (old process must quit) so the new code loads; AES GURI button will raise + refresh

### 2026-09-22 — GURI ribbon: bring existing window to front
- `AesRibbonHost\Connect.cs` — after launching/`--raise`, Outlook (click owner) finds the "GURI Database Viewer" HWND and `SetForegroundWindow` / restore; polls briefly on cold start
- `guri_gui.py` — RAISE IPC waits for payload; stronger restore + taskbar flash if focus steal is blocked
- Why: button logged successful `LaunchGuriGuiDirect` but GURI was already running; pythonw `--raise` exits after IPC and cannot steal focus from Outlook
- Follow-up: fully quit Outlook; re-run `AesRibbonHost\install.ps1`; restart. Optional: restart existing GURI so Python RAISE improvements load

### 2026-09-22 — Ribbon GURI without VBA toolbar button
- `AesRibbonHost\Connect.cs` — Home **GURI** no longer hard-fails when `AES_GURI` is missing on the CommandBar: try button → `Application.Run(ShowGuriGui)` → launch `guri_gui.py --raise` directly (same Python paths as VBA)
- Other ribbon actions also try `Application.Run` before showing the “button not found” dialog
- Why: Outlook still had an older AES bar (7 controls, no GURI) after ribbon GURI was added; click logged `no matching button for tag=AES_GURI`
- Follow-up: fully quit Outlook; `powershell -ExecutionPolicy Bypass -File C:\GeoFooter\AesRibbonHost\install.ps1`; restart. Optionally re-import `MSCANToolbar.bas` so Add-ins bar also shows GURI

### 2026-09-22 — Smooth threat score (less jumpy)
- `VBA\geolocate_headers.py` — three de-jitter changes:
  - Transient lookup failures no longer add risk: "security flag data unavailable" capped at +2 total (was +2 per flag, up to +8); WHOIS lookup errors now informational (was +5)
  - AbuseIPDB graded band replaces hard cliff (origin + route hops): +8 in threshold..threshold+14, +15 at threshold+15 or listed ASN (was 0→+15 at the threshold)
  - Per-sender score smoothing via `aes_sender_history.json` (`score_ema`): blends raw score with sender history (rise 0.6 / fall 0.4, ±3 deadband); raw ≥ 70 and forced-high signals (blocklist, domain <10 days) bypass and show immediately; adjustment appears in the score breakdown as "Score smoothing"
- Why: score bounced scan-to-scan on the same sender from API wobble and threshold cliffs, not real risk changes
- Follow-up: none for VBA; Python-only change, takes effect on next scan

### 2026-09-22 — Outlook freeze on new mail (latch + install)
- `VBA\MSCANModQueueManager.bas` — refuse early ItemLoad drain of queue tick (must wait `QUEUE_TICK_MIN_SEC`); delay 2.5s; max 1 in-flight; remove `DoEvents` from ProcessQueue; quieter VBS nudge (3x)
- `VBA\MSCANModule1.bas` — compact export headers-only (no HTMLBody/beacon/body sidecars); geo job nudges 12x/3s
- `AesRibbonHost\install.ps1` — no longer abort on Wow6432Node mirror; dual-register CLSID; always write Outlook Add-in key
- Why: logs showed ProcessQueue in the same second as ItemAdd (early latch drain); install died before Add-in key when mirror failed
- Follow-up: re-import `MSCANModQueueManager.bas` + `MSCANModule1.bas`; Compile; fully quit Outlook; re-run `AesRibbonHost\install.ps1` (already OK this session); reopen Outlook

### 2026-09-17 — AES large GURI icon → foreground GURI GUI
- `VBA\MSCANToolbar.bas` — large **GURI** button (`ShowGuriGui`) launches `guri_gui.py --raise`
- `guri_gui.py` — single-instance IPC: second launch / `--raise` brings existing window to foreground
- `make_aes_icons.py` + icons — `aes_guri.bmp` / mask / png in `icons\` and `VBA\icons\`
- `AesRibbonHost\Ribbon.xml` + `Connect.cs` — large Home ribbon **GURI** button
- Why: full-size AES control to surface GURI as the main analysis UI
- Follow-up: re-import `MSCANToolbar.bas`; quit Outlook fully; `powershell -ExecutionPolicy Bypass -File C:\GeoFooter\AesRibbonHost\install.ps1`; restart Outlook

### 2026-09-17 — Project agent docs + work log
- Added `AGENTS.md`, `SKILLS.md`, `.cursor/rules/agents-md.mdc`, `WORKLOG.md`
- Why: give Cursor persistent suite context, persona/skills, always-load rule, and a running change log
- Follow-up: none

### 2026-09-17 — Outlook freeze on new mail
- `VBA\MSCANModQueueManager.bas` — removed immediate `DeferredProcessQueue` on high-priority enqueue; shorter queue tick; no full-inbox Sort on Restrict failure
- `VBA\MSCANEmailWatcher.cls` — removed `DrainPostSendSubjects` from `ItemAdd`
- `VBA\MSCANModule1.bas` — compact export skips attachment SaveAsFile + beacon neutralize; body export capped; neutralize on footer apply
- `VBA\geolocate_headers.py` — compact lite (skip crt.sh/AV; light hop enrichment)
- Why: Outlook froze when new mail arrived because sync export ran on the UI event thread
- Follow-up: re-import `MSCANModQueueManager`, `MSCANEmailWatcher`, `MSCANModule1`, and `MSCANModSenderRules` (for `QuarantineAllAttachments`); Compile; restart Outlook

### 2026-09-07 — AES status banner + action buttons
- Banner PNG: border on image, clip-guard, extra top text padding; HTML avoid tall cream cells
- Action buttons: nested tables for Outlook; restored `border-radius: 3px`
- Why: banner height/alignment/clipping and jammed/sharp action buttons in Outlook
- Follow-up: Short Scan to refresh banner/footer
