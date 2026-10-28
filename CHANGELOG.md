# Changelog

All notable changes to the **GeoFooter / AES / GURI / Aura** suite are documented here.

Format follows [Keep a Changelog](https://keepachangelog.com/en/1.1.0/).
Versioning follows [Semantic Versioning](https://semver.org/spec/v2.0.0.html):
**MAJOR.MINOR.PATCH** — roughly **90% of releases are patches**.

| Bump      | When                                                            |
| --------- | --------------------------------------------------------------- |
| **PATCH** | Default — fixes, wording, scoring tweaks, icons, docs           |
| **MINOR** | New backward-compatible features (ribbon buttons, tabs, modes)  |
| **MAJOR** | Rare — breaking Outlook/VBA contracts, DB schema, or public CLI |

The canonical version string lives in `geofooter/version.py` and the plain-text `VERSION` file.
Bump those together with this changelog and `AesRibbonHost/AesRibbonHost.csproj` `<Version>`.

Every git commit must start with a short **caption** (one line) stating what was done,
e.g. `Fix Create-GURI silent insert failure for long document paths`.

## [Unreleased]

## [1.5.59] — 2026-10-28

### Fixed

- `MSCANCore.GetVersion` reads install-root `VERSION` instead of a hard-coded `2.0.0` (BUG-039). Re-import `MSCANCore`.

## [1.5.58] — 2026-10-22

### Changed

- AES Diagnostics default window size is 1800×1000.

## [1.5.57] — 2026-10-22

### Changed

- AES Diagnostics is a two-column window: Live Controls (timers, logging, capture, lockup scanner) on the left; Run all / test list / result detail on the right.

## [1.5.56] — 2026-10-22

### Fixed

- CatchUp / queue skip paths use `IsAesScannedLite` (UserProperty or `aes_mark_*` attachment only) so idle sweeps never read `HTMLBody` on the UI thread (BUG-036).
- CatchUp and post-send inbox walks use `GetFirst`/`GetNext` instead of `Items.Count` (avoids full Restricted-set enumeration).
- Idle CatchUp heartbeat backs off to 240s when the last sweep queued nothing; new mail resets the streak.
- `StripSensitiveHeadersFromMail` stops retrying after one unsupported `SetProperty` on the store (BUG-037). Requires VBA re-import.

## [1.5.55] — 2026-10-22

### Added

- AES Diagnostics → Live Controls: **Start / Stop lockup scanner**. An out-of-process watcher (`aes/outlook_lockup_scanner.py`) probes Outlook with `SendMessageTimeout`; on hang it writes a freeze report under `%LOCALAPPDATA%\GeoFooter\Logs\lockup_scans\` (process snapshot + VBA / ribbon / geolocate tails) and appends `outlook_lockup_events.log`. Starting the scanner also kicks off VBA log capture.
- `CatchUpMissedInboxMail` is wrapped in `MSCANHealth.BeginOp` / `EndOp` so long Restrict walks show up as `HEALTH | UI stall` / uiStalls (PERF-002). Requires VBA re-import.

## [1.5.54] — 2026-10-21

### Added

- Footer Quick Action chips (SL / BA / BB / TS / NT) and the AURA chip show a hover explanation (`title=`). Outlook shows these as tooltips on the reading-pane links.
- AURA chip opens GURI on the Aura → Brokers tab and preselects the catalog broker for that mail's sender domain when known (`aes://open-aura?sender=…&domain=…&broker=…`). Restart GURI once so an already-running instance picks up the new IPC. Rescan to refresh footers.

## [1.5.53] — 2026-10-21

### Changed

- Footer quip (top row, right) is raised 2px so it sits in the middle of the row instead of on the rule below. Rescan to refresh footers.

## [1.5.52] — 2026-10-21

### Fixed

- Footer rows centre on the whole strip. Rows used to share one column grid, and Outlook spread the spare width into the right-hand column, which pushed the top and second rows left. Each row now has its own fixed-width right cell (BUG-033).
- AURA lettering is centred in its chip. The image is now the full 76x16 interior; Word had placed the shorter image on the text baseline, leaving a gap below. Rescan to refresh footers.

## [1.5.51] — 2026-10-21

### Changed

- Footer AURA chip: the image no longer draws its own frame inside the chip border, and the lettering is larger and spread across the chip, with the rainbow running red to violet across the letters. Rescan to refresh footers.

## [1.5.50] — 2026-10-21

### Added

- Footer second row far-right **AURA** chip (same 78×16 frame as Quick Actions). Rainbow-gradient lettering via `assets/icons/aes_aura_chip.png`. Click opens GURI on the Aura tab (`aes://open-aura`).

### Fixed

- A/E/S letters are one `rowspan` column with slate rule spacers, so Outlook no longer paints a white seam under the E (BUG-031). Rescan to refresh footers.

## [1.5.49] — 2026-10-21

### Fixed

- Footer strip rules stay out of the A/E/S column (solid slate cell there instead). Removes the white hairline through the E; 1.5.48's letter `border-bottom` had made Outlook paint it white (BUG-031). Rescan to refresh footers.

## [1.5.48] — 2026-10-21

### Fixed

- Footer top summary row: title and underlined metrics now share one vertical centre (border-bottom underlines instead of `text-decoration`, matching edge on plain title). White hairline under the E letter painted out with a strip-coloured `border-bottom` (BUG-031, BUG-032). Rescan to refresh footers.

## [1.5.47] — 2026-10-21

### Changed

- Footer `QUICK ACTIONS` label is beside the chips again (undoes the stacked label from 1.5.45). Rescan to refresh footers.

## [1.5.46] — 2026-10-21

### Fixed

- Rapid rescans could leave a mail as plain text. Each footer apply was deleting and re-adding the orange mark and A/E/S images, which dirtied the IMAP item and raced `Save`. Re-import `MSCANModule1` (BUG-030).

## [1.5.45] — 2026-10-21

### Changed

- Footer bottom row centres on the action chips alone. `QUICK ACTIONS` is back as its own centred line above the chips, so the label no longer pulls the buttons left. Rescan to refresh footers.

## [1.5.44] — 2026-10-21

### Fixed

- Footer showed thin white lines to the right of the A and E letters. Outlook was leaving a 2px gap between those letter cells and the summary cells; the join is now painted with the strip colour (BUG-029). Rescan to refresh footers.

## [1.5.43] — 2026-10-21

### Changed

- Footer bottom row (quick actions / copyright) has 2px more space under the content. Rescan to refresh footers.

## [1.5.42] — 2026-10-21

### Fixed

- Footer top row (text and the A of A/E/S) sat a few px low, and the orange A was not centred in the strip. In Outlook the mark column was taller than the rows, and Word pushed the extra height into row 1. The mark spans the rows again with 3px above and none below, so all three rows come out the same height and the orange A sits 4px from the top and 3px from the bottom (BUG-028).

### Changed

- Copyright version reads `v.1.5.42` (no space). Rescan to refresh footers.

## [1.5.41] — 2026-10-21

### Fixed

- Footer summary text sat at the top of each row. The orange mark no longer uses rowspan, so rows keep their own height and the text centres vertically.

### Changed

- Footer copyright shows `v.` before the suite version. Rescan to refresh footers.

## [1.5.40] — 2026-10-21

### Fixed

- Footer top summary row was about 5px taller than the others. Its vertical padding is now 0 (BUG-028). Rescan to refresh footers.

## [1.5.39] — 2026-10-21

### Fixed

- Footer BB (blocked beacons) chip matched the fill colour for its border, so it looked wider than SL/BA/TS/NT. All chips now share one dark border and a fixed 78px cell with no letter-spacing (BUG-027).

### Changed

- Footer copyright reads `© 2026 Aliniant Labs` plus the suite version. Rescan to refresh footers.

## [1.5.38] — 2026-10-21

### Changed

- Footer orange A sits 5px lower (8px from the top of the strip instead of 3px). Rescan to refresh footers.

## [1.5.37] — 2026-10-21

### Changed

- Footer A/E/S letters are centred on one another. A is wider than E and S, so left-aligning them made the column look ragged. Rescan to refresh footers.

## [1.5.36] — 2026-10-21

### Fixed

- Footer A/E/S letters sat high in each row with a gap below, and made the rows taller. The letter cells no longer have padding or a fixed line height; the letters are 14px and centred, about 4px from the top and bottom of each row in Outlook. Rescan to refresh footers (BUG-025).

## [1.5.35] — 2026-10-21

### Fixed

- Footer A/E/S letters still showed as empty boxes after the VBA re-import. The images were attached, but VBA gave them garbage Content-IDs, so the footer could not find them. Re-import MSCANModule1 (BUG-025).

## [1.5.34] — 2026-10-21

### Fixed

- A rescan could leave a mail as plain text with no footer. The follow-up step that attached the A/E/S letter images rewrote and saved the mail from outside Outlook, and that collided with the rescan. It now only attaches the orange mark, and only when Outlook hasn't (BUG-026). The letters come from the VBA step once MSCANModule1 is re-imported.

## [1.5.33] — 2026-10-21

### Fixed

- Footer A/E/S letters showed as broken-image boxes before the VBA re-import. The follow-up attach step now waits until Outlook has written the new footer before it checks which images are missing (BUG-025).

## [1.5.32] — 2026-10-21

### Changed

- Footer strip is as short as its content allows: 62px, down from about 84px. The text is unchanged. Rows are 20px each, and the orange mark is scaled to fit them.
- **A / E / S** are now crisp Montserrat Bold images, 4px from the top and bottom of each row. Outlook can no longer pad them or stretch the rows (BUG-025). Re-import MSCANModule1 so new footers attach them directly; until then, the scan's follow-up step attaches them.

## [1.5.31] — 2026-10-21

### Changed

- Footer **A / E / S** letters fill each summary row with a 3px top/bottom inset (20px glyphs in a locked 26px row). The strip no longer grows from an oversized line-height.

## [1.5.30] — 2026-10-21

### Fixed

- Hop timestamp parsing: only call `total_seconds` when `dt.utcoffset()` is not `None`.

## [1.5.29] — 2026-10-21

### Changed

- Footer **A / E / S** letters are much larger (32px / ~24pt, HTML `size=6`) so Outlook's Word renderer actually shows the bump.

## [1.5.28] — 2026-10-21

### Fixed

- Footer builder: initialise `block_state` as `{}` and read `trusted` only inside the `isinstance(..., dict)` branch, so nothing calls `.get` on a possibly-`None` value.

## [1.5.27] — 2026-10-21

### Fixed

- Footer builder: use an `if`/`else` to narrow `block_state` (the ternary form still left some checkers treating it as `None`).

## [1.5.26] — 2026-10-21

### Fixed

- Footer builder: narrow `block_state` to a dict before `.get("trusted")`.

## [1.5.25] — 2026-10-21

### Fixed

- RDAP registrant contact: assign once and `isinstance`-narrow so `contact.get` is never typed against `None`.

## [1.5.24] — 2026-10-21

### Fixed

- Footer mark embed job: `_com_get` is annotated to return `Any`, so Outlook COM attributes like `CurrentItem` type-check.

## [1.5.23] — 2026-10-21

### Fixed

- Hop address decode: after the `addr_decode is None` guard, bind a local `ad` and use it in the nested `_host_loc` helper so the type checker keeps the narrowing.

## [1.5.22] — 2026-10-21

### Fixed

- RDAP registrant geo: assign `objects` from a single `data.get` then narrow with `isinstance`, so it is never typed as `None` before `.get`.

## [1.5.21] — 2026-10-21

### Fixed

- Threat intel HTML builder: coerce `worst` / `kind` / `verdict` / `provider` to `str` before `dict.get`, fixing the type-checker overload error on the sort key.

## [1.5.20] — 2026-10-21

### Fixed

- Threat intel path: `_run_threat_intel` returns early when the module failed to import, so type checkers no longer treat `ThreatIntel` / `registrable_domain` as attributes of `None`.

## [1.5.19] — 2026-10-21

### Fixed

- Status banner PNG: the text box from Pillow is rounded up to whole pixels, so the image size passed to `Image.new` is always integers.

## [1.5.18] — 2026-10-21

### Fixed

- WHOIS lookup: the `concurrent.futures` import now sits before the `try`, so the `except FuturesTimeout` clause can never reference an unbound name.

## [1.5.17] — 2026-10-21

### Changed

- Footer **A / E / S** letters are larger (22px, up from 17px). Colour and position are unchanged.

## [1.5.16] — 2026-10-21

### Fixed

- The two summary lines in the footer are now centred on the strip. Before, the top line sat a little left of centre and the sender line a little right. The quick-actions row is unchanged (BUG-024).

## [1.5.15] — 2026-10-21

### Changed

- Footer **A / E / S** letters now start exactly where the strip's horizontal lines start, left-aligned to them. They are 4pt bigger (17px), bold, and a little lighter than the strip background. Outlook had been showing them black; the colour is now set in a way Outlook keeps (BUG-023).

## [1.5.14] — 2026-10-21

### Fixed

- Footer strip banding from re-processing: same-colour ghost spacer text is gone, and the orange A mark tile has a solid slate background (BUG-023).

### Added

- **A / E / S** beside the footer mark, one letter per content row, in Montserrat Bold at a grey just off the strip background.

## [1.5.13] — 2026-10-21

### Fixed

- **GURI no longer jumps to the front** while you are in Outlook or another app. Focus is only taken when you open it from the tray, the ribbon (`--raise` / Aura), or on first launch. A bare second start of `gui.py` leaves the running window alone (BUG-022).

## [1.5.12] — 2026-10-20

### Changed

- The aliniant.net holding image now scales to the window height (leaving room for the welcome text), keeping its aspect ratio and never overflowing the width.

## [1.5.11] — 2026-10-20

### Fixed

- The aliniant.net holding image was already updated on the server, but browsers kept the old one for up to 7 days. The file is now `aliniant-village.jpg` so a refresh loads the new picture.

## [1.5.10] — 2026-10-20

### Changed

- **aliniant.net holding page** uses the new solarpunk Shire image, with dark-grey side panels and cream text so the paw trails stay readable.

## [1.5.9] — 2026-10-20

### Changed

- **aliniant.net holding page:** under the picture it now says "You've reached www.aliniant.net", points to www.aliniant.com for details and says engineering tools are coming soon. It is signed with a dragon's paw print, and a trail of paw prints runs up the left of the picture and down the right (smaller on phones).

## [1.5.8] — 2026-10-20

### Fixed

- Images under `cloud/` are stored as binary in git. The `cloud/** eol=lf` rule had altered the holding page's JPEG and favicon when they were committed (BUG-021).

## [1.5.7] — 2026-10-20

### Added

- **aliniant.net holding page.** A single static page with the "Aliniant — the wise company with clever ideas" watercolour. `www.aliniant.net` and plain HTTP redirect to `https://aliniant.net/`. Source and setup script are in `cloud/site/aliniant.net/` (strict Content-Security-Policy, only the page, the image and the favicon are served, rate limited).

## [1.5.6] — 2026-10-20

### Changed

- **Read tracking is live.** The pixel endpoint runs at `https://track.aliniant.net` and tracking is switched on. Outgoing Aliniant mail carries the pixel once `MSCANPixel`, `MSCANEventHandlers` and `MSCANModule1` are re-imported.

## [1.5.5] — 2026-10-20

### Fixed

- `cloud/pixel/deploy.sh` is now safe on a VM that already serves other sites. It installs only missing packages (so it never upgrades or restarts a running nginx), enables the site as `zz-aes-pixel` so it can't become the catch-all for other names, and removes nginx's default site only on a fresh install (BUG-020).

## [1.5.4] — 2026-10-20

### Changed

- Read-tracking pixels are served from `https://track.aliniant.net` on a dedicated Aliniant VM, which keeps them off the TXOne server entirely. The nginx site files are renamed to `cloud/pixel/nginx-track.aliniant.net*.conf`.

### Added

- `cloud/pixel/deploy.sh`: sets up the pixel endpoint on a fresh Ubuntu 24.04 VM (packages, service, nginx, Let's Encrypt certificate, health check). The key is read from stdin. Safe to re-run.
- `.gitattributes`: `cloud/**` is checked out with LF line endings so bash and systemd can read the files.

## [1.5.3] — 2026-10-20

### Fixed

- Pixel nginx site files (not deployed yet) can no longer take over i-txone.com traffic that matches no site name. They are installed as `zz-aes-pixel` so `workflow-api` stays the catch-all, the IPv6 listeners are gone so no new sockets open, and a per-IP connection limit is added (BUG-020).

## [1.5.2] — 2026-10-20

### Changed

- Read-tracking pixels are served from `https://track.i-txone.com` instead of `track.aliniant.com` (aliniant.com is a separate WordPress site). The i-txone.com wildcard DNS already covers the name, so no DNS change is needed. The nginx site files are renamed to `cloud/pixel/nginx-track.i-txone.com*.conf`. `base_url` in `aes_pixel.ini` still overrides the default.

## [1.5.1] — 2026-10-20

### Added

- High-risk mitigated mail (score above 70) opens with a link, **"Not Dangerous you big scaredy cat, turn me back into a proper email"**, which restores the original HTML.

### Fixed

- The AES actions in that mail (Show links, Restore, Block, Trust, the full report) are now clickable. They were raw `aes://` lines in a plain-text body, which Outlook never turns into links. The body is now HTML written entirely by AES, with the sender's text escaped and none of their markup, images or links kept (BUG-019). Full No Trust still converts to plain text.

## [1.5.0] — 2026-10-20

### Added

- **Read tracking for outgoing mail.** HTML mail sent from the Aliniant account gets an invisible AES pixel (`https://track.aliniant.com/p/<id>.gif`), added at send time after classification and ReadNotify. Earlier AES pixels quoted in replies and forwards are removed first.
- **GURI Tracking tab.** Lists each tracked send with opens, likely reads, first/last open and last client, plus a per-open detail list. Opens are collected every 5 minutes in the background while GURI is open. Each tracked send gets a GURI record (risk field `OUTGOING (tracked)`). Your own views are left out (opens from the IP that polls), and Gmail/Yahoo proxy fetches and Apple Mail pre-fetches are labelled.
- **Pixel endpoint** in `cloud/pixel/`: a stdlib-only Python service (loopback port 8791, SQLite, 400-day retention, per-pixel cap), a hardened systemd unit (64 MB / 10% CPU) and a self-contained nginx site for `track.aliniant.com`.

### Changed

- Our own pixel quoted back in an incoming reply no longer counts as a third-party beacon.

### Notes

- Off by default. Set `enabled=1` in `%LOCALAPPDATA%\GeoFooter\aes_pixel.ini` once the endpoint is deployed. Re-import `MSCANPixel` (new), `MSCANEventHandlers` and `MSCANModule1`.

## [1.4.30] — 2026-10-20

### Fixed

- The AES action buttons at the bottom of a full scan no longer lose their bottom edge. The button cells now have a fixed 14px line height and extra padding, with an 8px spacer under the row, so Outlook has nothing to clip.

## [1.4.29] — 2026-10-20

### Fixed

- Shodan in the hop Reputation column no longer reads "lookup failed (HTTPError)" on the free (oss) plan. That plan refuses host lookups for most IPs with 403 "Requires membership or higher"; the scan now shows "free plan can't look up this IP (needs Membership)" and treats it as no data. A 429 reads "rate limited".
- The Settings Shodan test notes when the key is on the free plan, since the key itself checks out fine there.

## [1.4.28] — 2026-10-20

### Fixed

- The full security analysis embedded in the footer no longer takes on the sender's page colour (a green newsletter body turned the whole analysis green). Outlook ignores `div` backgrounds, so the analysis now sits in a white table cell.

## [1.4.27] — 2026-10-20

### Documentation

- `docs/bug_tracker.md` now records the bugs found this week: Settings freezing Outlook (BUG-009), sleep logged as UI stalls (BUG-010), the false ribbon "button not found" popup (BUG-011), the hop ASN mix-up (BUG-012, closed), the header strip that fails on every footer (BUG-013) and the wrong `whois` package in the venv (BUG-014).
- New always-on rule `.cursor/rules/bug-tracker.mdc`: every bug seen is logged in the tracker in the same turn, and updated in the same commit as its fix.

## [1.4.26] — 2026-10-20

### Fixed

- Opening AES Settings no longer freezes Outlook until the dialog is closed. VBA starts the dialog through a small WScript launcher and applies the result on the next item load after the dialog closes. Clicking Settings again while it is open says so instead of opening a second copy.
- A footer save or other Outlook-thread operation that straddles a PC sleep is no longer logged as a UI stall of several hours (seen as 79,700,593 ms and 70,033,485 ms). Readings over five minutes are logged once as a probable sleep and left out of the stall counters.

## [1.4.25] — 2026-10-20

### Fixed

- Clicking Short Scan (or another AES button) on the ribbon no longer shows "AES button 'ShortScanEmail' was not found on the VBA toolbar" while the scan is actually running. The toolbar button's ``Execute`` can throw back after the VBA macro has started, and the ribbon host took that as "not found".

## [1.4.24] — 2026-10-20

### Fixed

- Outlook startup no longer yields with ``DoEvents`` before the queue manager is ready. A queued ``ItemLoad`` could reach ``ProcessQueue`` before initialisation and the startup quiet period.
- The AES log is size-checked every 200 lines, not on every line, which saves a FileSystemObject create and a file stat per line on the UI thread. A failed log write closes its file handle. A ``max_mib`` of 0 (from older settings, which disabled trimming) now means 16 MiB, values under 4 MiB are raised to 4, and ``max_mib`` is saved with the logging settings.
- Restoring the persistent status no longer reschedules itself on every ``ItemLoad``, which launched a ``wscript`` for each message clicked.

## [1.4.23] — 2026-10-20

### Added

- Conversation-hijack check. A first-time sender whose "Re:" message cites In-Reply-To/References IDs that AES has never seen scores +15 ("Claims to continue a thread we never saw"). A first-time "Re:" with no thread headers at all scores +10. Senders in GURI's recipients list or your own domains, bulk or automated mail, and newsletters and read receipts that only carry References are not scored. Scoring needs at least 14 days of sender history; with less, the Thread row is report-only. Run against 221 recent mails, it flagged none.

## [1.4.22] — 2026-10-20

### Added

- TLS version and cipher are read from each Received line and shown under the hop in the Route Table. Outdated TLS (1.0/1.1, SSL) on an external hop scores +5, and plain SMTP on an external hop scores +3.
- Sending client row in the Full Scan report: Message-ID host, X-Mailer, User-Agent and X-Originating-IP. Old PHPMailer or SwiftMailer, raw PHP `mail()`, or a script library scores +10. A Message-ID host that names a Windows PC (`DESKTOP-…`, `WIN-…`) scores +5.
- Mail type row (bulk/list, automated, transactional through an ESP, or personal) with List-Unsubscribe and RFC 8058 one-click. Unsubscribe links go through the link checker with the body links.

## [1.4.21] — 2026-10-20

### Added

- Signing / alignment row in the Full Scan report. Every DKIM signature is listed with its domain, selector and algorithm, so dual signing shows as "Signed by carto.com and sendgrid.info". Each DKIM and SPF domain is marked strict, relaxed or not aligned with From. A signature that leaves From or Subject unsigned (or Reply-To, when one is present) scores +5.
- On full and deep scans AES looks up the DKIM key and the published DMARC record, cached for 24 hours. An RSA key under 1024 bits scores +5. `p=none` is reported as monitoring only. Compact scans use the policy Gmail writes into Authentication-Results and make no DNS calls.
- `cryptography` added to `requirements.txt` for reading DKIM key sizes.

## [1.4.20] — 2026-10-20

### Added

- Sending-platform fingerprint for SendGrid, Mailgun, Amazon SES, Postmark, Mailchimp/Mandrill and Salesforce, read from their X- headers, Feedback-ID, bounce address and hop names. The SendGrid HTTP API hop's base64 name is decoded to the account ID, so hop 1 reads "SendGrid API, account 21732389" instead of Unknown. The Full Scan report has a Sending platform row with the platform, account and campaign ID.
- Tenant lists in `%LOCALAPPDATA%\GeoFooter\aes_esp_tenants.json` (`{"trusted": ["sendgrid:21732389"], "blocked": []}`). A blocked tenant scores +25. A trusted tenant's shared-IP hops are left out of route/ASN risk.

## [1.4.19] — 2026-10-20

### Added

- Identity checks on every scan, compact included. Reply-To on another domain scores +10, or +15 when it is free webmail behind a corporate From. A display name that shows another address or domain scores +15. A brand name (PayPal, Microsoft, DocuSign and others) or your own organisation's name on an unrelated sending domain scores +10. A Return-Path on another domain scores +5, but only when DMARC did not pass. The Full Scan report has an Identity row listing From, Reply-To, Return-Path and Sender.

## [1.4.18] — 2026-10-20

### Fixed

- The Route Table no longer shows the receiving server's ASN on a hop. A SendGrid IP delivered to Gmail showed AS15169 (Google) because the `by` host was run through the hostname-to-ASN map. Only the sending host names are used now, and a registry ASN is replaced only when it is BGP transit. `*.sendgrid.net` maps to AS11377.

## [1.4.17] — 2026-10-20

### Fixed

- Create GURI no longer asks to add a recipient on every keystroke. The prompt runs when you leave the Recipients field or press Enter, and only when the entry looks finished (for example `name@domain.com`).

## [1.4.16] — 2026-10-19

### Changed

- The footer strip is a slightly lighter slate (`#46525c`, was `#39444d`). The row lines and the orange A tile use the same step so they still match.

## [1.4.15] — 2026-10-19

### Fixed

- A message that stays inside one domain shows the hop address and the location "Internal within same domain". The summary no longer says Sender IP Unknown.
- That mail is no longer scored for missing SPF, DKIM, DMARC, or origin geo. Those checks apply to mail that arrived from outside. A Coleago-to-Coleago message that scored about 19 now scores 0 unless something else on the message (a link, an attachment) adds risk.

## [1.4.14] — 2026-10-08

### Fixed

- GURI no longer forces itself maximized. Launch, the tray, a second start, and the Aura tab bring the existing window forward at the size you left. The first open is a normal window, and that size is remembered.

## [1.4.13] — 2026-10-05

### Fixed

- Quick Actions rescan the open message. SL, BA, BB, TS, and NT each start AES Short Scan. The previous call used `FindControl(Tag=...)`, which from Python returns the Task Pane instead of the Short Scan button, so the click saved the rule and did not scan.

## [1.4.12] — 2026-10-05

### Fixed

- A finished footer is written onto the message that is open in the reading pane. Leaving it until the user clicked away meant the footer never appeared, and the job script kept polling.
- A scan that already has a footer is no longer failed at the 90 second deadline. That deadline deleted the Kenable parcel footer and the Product Lead footer while those messages stayed selected.

## [1.4.11] — 2026-10-05

### Fixed

- A finished footer for the message on screen was removed from the pending list and never applied. The job stays pending and is applied once the message is no longer selected.
- That waiting footer no longer occupies the only scan slot, so other mail keeps scanning while one message is open.
- A failed footer save discards the unsaved edit before the retry, and warns when Outlook keeps the unsaved copy anyway (the reading pane does).
- Compose detection still counts a draft as active when the window title has a suffix the caption does not.

## [1.4.10] — 2026-10-05

### Fixed

- An idle draft no longer stops AES. `IsUserComposing` counted any open unsent Inspector as typing, so a reply left open in the background deferred every footer, queue tick and job timeout. It now needs input in the last 4 seconds and a compose window in the foreground.
- A rescan supersedes any pending job for the same mail. The stale job used to commit right after the rescan touched the item, fail with "message has been changed", and leave Outlook's copy without a body.
- `CommitMailHtml` refuses to save a body that is empty apart from the AES footer, and discards the unsaved changes after a failed commit so a later Save cannot persist an emptied copy. The high-risk text-only path has the same guard.
- Beacon blocking in the footer paths edits the HTML string instead of making its own `HTMLBody` write and Save before the footer commit.
- `ReconcileAsyncJobs` checks the job deadline before looking for output, and failed scan starts release the body payload directory as well as the attachment one.

## [1.4.9] — 2026-09-30

### Fixed

- A ribbon button repaint could rewrite a message body. `PendingAsyncJobCount` reconciles pending jobs first, so the chain `GetServiceToggleLabel` -> `IsScanBusy` -> `PendingAsyncJobCount` -> `ReconcileAsyncJobs` -> `CompleteAsyncFooter` ended in an `HTMLBody` write. Outlook calls ribbon getters often, unpredictably, and re-entrantly during `Invalidate`, so mail was being rewritten at moments unrelated to anything the user did. Added a read-only `PendingAsyncJobCountFast` and pointed every ribbon getter, toolbar state write, and diagnostic read at it; reconciliation now happens only on deliberate nudge and queue-tick paths.
- Automatic scans no longer rewrite the message the user is reading. `MSCANIdle.IsItemOnScreen` checks the reading-pane selection and open inspectors, and `CompleteAsyncFooter` defers the footer to the next nudge when the target is on screen. Explicit ribbon scans are exempt.
- Beacon blocking rewrote the body on every scan of the same message. In the default `defang` mode the tag keeps its dimensions and only its `src` changes, so `IsLikelyTrackingBeacon` matched it again forever. Defanged tags are now stamped with `data-aes-defanged` and skipped, and an unchanged body is never written.
- `IsAesScanned` treated any failed property read as "not scanned", sending an already-footered message through a full rescan and another rewrite. It now falls back to the body markers, and an unreadable body answers "scanned" — a missing footer is cheaper than a rewrite on every pass.
- `MarkAesScanned` was a `Sub` wrapped in `On Error Resume Next`, so a failed save left the message footered but looking unscanned, guaranteeing a rescan. It is now a `Function` that reports failure and logs it.
- `CompleteAsyncFooter` reported success to `NoteScanOutcome` even when it had abandoned the footer, so the failure streak reset every time and the circuit breaker never tripped on a store refusing writes.
- A synthesised `<!-- AES Start -->` marker was added without its matching end marker. `StripBlockBetween` only removes bounded blocks, so such a footer could never be stripped and every later scan would append another one.

### Added

- Health counters for on-screen deferrals and scanned-mark failures, both reported in the session snapshot and by `scripts/aes_log_health.py`.

## [1.4.8] — 2026-09-30

### Fixed

- A message could end up with more than one scan footer. `StripAllAesScanResults` only removed a block whose **end** marker survived, and markers do not always survive: IMAP strips HTML comments and Outlook or Word can rewrite the `aes-footer-start` anchor. A footer quoted from an earlier email in a reply chain therefore stayed in the body and the new footer was appended beside it. Any block still carrying the rendered "Aliniant AES Scan Result" heading is now removed by its enclosing `<div>` or `<table>`, matching nested tags so only the footer is cut. If no container can be identified the body is left untouched rather than guessed at.
- `CountAesScanResults` returned after the first marker style that matched, so a body holding one comment-marked footer and one stripped footer reported 1. It now counts each style and takes the highest.
- A message already stamped as scanned was skipped without checking how many footers it carried. It is now rewritten whenever it holds more than one.

### Added

- The finished body is checked before it is written and anything other than exactly one footer is logged as a `HEALTH` warning naming the message, counted in the session snapshot as `dupFooters` / `missingFooters`, and reported by `scripts/aes_log_health.py`.

## [1.4.7] — 2026-09-29

### Fixed

- Every scan wrote a one-shot `aes_geo_job_*.vbs` into `%LOCALAPPDATA%\GeoFooter` and nothing ever removed it. 3,290 had accumulated, dating back to 8 July. AES writes the delayed queue tick script into that same folder, and the live log shows one `LaunchDelayedQueueTick: could not write tick script` failure. `MSCANModule1.PurgeOldJobScripts` now runs once at startup and deletes job scripts older than two days.
- `scripts/aes_log_health.py` reports the job script count so the build-up is visible before it causes a failure.

## [1.4.6] — 2026-09-29

### Fixed

- The message list flickered and Outlook became unusable when a message would not accept a footer. Outlook returned `-2147221239` ("the message has been changed") and AES retried the whole write up to 30 times for a single mail — 3 outer passes x 2 body variants x 5 commits — each one rewriting `HTMLBody` and repainting the item. A conflicting item is now retried once after being re-resolved and then left for a later queue pass, capping it at a few attempts.
- `VBA_Log.txt` had no default size cap and had grown to 37 MiB. It now defaults to 16 MiB; `aes_logging.json` still overrides it, and 0 is still unlimited.

### Added

- `VBA\MSCANHealth.bas` — session counters for commit conflicts, footers applied and abandoned, queue depth, compose deferrals, and time spent holding the Outlook UI thread. Anything past a threshold logs a `HEALTH` warning as it happens: a conflict storm, a queue backlog over 25 items, or an operation holding the UI thread longer than 1.2s (the point where typing starts dropping letters). A `HEALTH` summary line is written periodically and at shutdown.
- `scripts/aes_log_health.py` — parses the log tail and reports counts, queue peaks, stall durations, the messages that would not accept a footer, and a plain-language verdict. `--session` limits it to the current Outlook session.

## [1.4.5] — 2026-09-25

### Fixed

- Typing in Outlook froze and dropped letters. AES waited on the UI thread with `DoEvents` spin loops (footer apply retries, HTML commit retries, waiting for the Python process), which re-entered Outlook's message pump and swallowed keystrokes. All of them now sleep instead.
- AES no longer scans mail or applies footers while a message is being written. Compose windows and inline replies are both detected; queued mail drains once the window closes.

## [1.4.4] — 2026-09-25

### Added

- The AES Link Safety page has a Your rating column. Low, Medium, or High is stored for that URL and used on later scans. Use scanner clears it.

## [1.4.3] — 2026-09-25

### Changed

- Clicking a witticism loads it into the line box, and Add becomes Edit so that line can be updated.

## [1.4.2] — 2026-09-25

### Changed

- The Witticisms model is a dropdown, and Test Connection checks the Poe key and that the selected model is listed.

### Fixed

- The ribbon Settings button opens the AES dialog when the VBA toolbar button is missing. Outlook has no `Application.Run`, so the host starts the Python dialog itself.

## [1.4.1] — 2026-09-25

### Fixed

- “1 NOK” is raised one pixel so it lines up with the rest of the top row.
- “Links:” and “HRE:” have a small gap before their values.

## [1.4.0] — 2026-09-25

### Added

- AES Settings has a Witticisms tab. Add or remove the footer one-liners for each risk band, and ask Poe for a new line. The Poe API key is stored with Windows DPAPI.

## [1.3.44] — 2026-09-25

### Fixed

- The footer A spans all three rows of the scan strip, centred against the whole block.

## [1.3.43] — 2026-09-25

### Fixed

- “Links”, “OK”, and “NOK” are separate cells on the top row, so NOK no longer sits below the rest of that line.

## [1.3.42] — 2026-09-25

### Fixed

- The Links count, including NOK, sits on the same baseline as the rest of the top row.
- The scan line, the sender line, and the Quick Actions are centred on the full footer. The logo and the side text no longer shift that centre.
- The bottom-right notice uses the © symbol.

## [1.3.41] — 2026-09-25

### Fixed

- The second footer line was printing the outbound-relay note as raw HTML. That note now renders, and the line is centred.

### Added

- The Quick Actions row shows “Copyright 2026 Aliniant Labs” on the far right.

## [1.3.40] — 2026-09-25

### Fixed

- The footer A was only sought in a separate message window. It now attaches to the mail selected in the reading pane.

## [1.3.39] — 2026-09-25

### Changed

- The cheeky footer lines are longer and funnier, and they still stay on one line.

## [1.3.38] — 2026-09-25

### Changed

- The cheeky footer line stays on one line.
- Top-row counts, including NOK, share the same baseline.
- The orange A is the logo image again, attached to the message after the footer is applied.

## [1.3.37] — 2026-09-25

### Changed

- Quick Action buttons are 16px tall. The fill and the border are the same box, so the colour no longer spills past the line.
- The orange A is drawn in the footer itself, so Outlook does not need a separate image.
- The top-right of the footer shows a random cheeky line for the risk band.

## [1.3.36] — 2026-09-25

### Fixed

- Compact scans crashed before writing a footer, so Outlook held AES PROC for five minutes per mail.
- A scan that overruns its budget, or dies before writing a footer, leaves a failure marker. DNS and WHOIS lookups are time-limited. Three failures in a row pause the queue for two minutes.

## [1.3.35] — 2026-09-25

### Changed

- Actions to take opens in time order, newest first. A column header sorts by that column; a second click reverses it.
- Timeline day numbers are paler and carry the day suffix (24th, 25th).

## [1.3.34] — 2026-09-25

### Changed

- Footer top-row labels, counts, and the green dots share one vertical centre.

## [1.3.33] — 2026-09-24

### Fixed

- The footer A was a blocked file link, which Outlook drew as an empty bar. It is now embedded in the message.

## [1.3.32] — 2026-09-24

### Changed

- The orange A is on the left of the footer, with the white background removed.
- A faint line separates each footer row.

## [1.3.31] — 2026-09-24

### Fixed

- A mail already converted to text stays text. A footer button no longer writes HTML into that message and starts a scan at the same time, which was flipping it back and forth.

## [1.3.30] — 2026-09-24

### Changed

- NT has three states. The chip stays the neutral colour. The first click removes Trust and blocks beacons. The second click is Full No Trust: attachments are blocked and that mail becomes plain text, with a restore link that is not opened until you click it.
- BA, BB, TS, and the first NT click replace the footer on the open message.

## [1.3.29] — 2026-09-24

### Changed

- Beacon Blocking in AES Settings is the default for every sender. The BB button on a mail overrides that default for that sender, including turning blocking off when the global rule is on.

## [1.3.28] — 2026-09-24

### Fixed

- With beacon blocking on for the sender, a detected beacon is counted as blocked. The line no longer says “1 detected 0 Blocked” when those beacons are blocked.

## [1.3.27] — 2026-09-24

### Changed

- The whole **HRE** label links to the risk breakdown. The count keeps its green, yellow, or red colour.

## [1.3.26] — 2026-09-24

### Changed

- **HRE** is a link to the risk-score breakdown, which lists the elements behind the count.
- Vertical padding on each footer row is 5px, half of the previous 10px.

## [1.3.25] — 2026-09-24

### Fixed

- A mail converted to text was being turned back into HTML on the next scan, then converted to text again. Scans no longer restore that HTML. Use the restore link in the text to put the HTML back.

## [1.3.24] — 2026-09-24

### Added

- **HRE** (high-risk elements) is shown to the right of Links. The count is green at 0, yellow at 1–3, and red above that.

## [1.3.23] — 2026-09-24

### Changed

- The three footer rows use the same vertical padding. The hairline between the detail line and Quick Actions is repeated between the title line and the detail line, at the same width.

## [1.3.22] — 2026-09-24

### Changed

- Attachment counts sit inside the status colour: green with **OK**, red with **NOK**.

## [1.3.21] — 2026-09-24

### Fixed

- **BB** stays white unless the sender is on the per-sender beacon list. With “block all scanned mail” on, and the sender not whitelisted, **BB** is now red.

## [1.3.20] — 2026-09-24

### Changed

- The risk words and score are colour-coded: **LOW** light green, **RAISED** (26–50) amber, **HIGH** (51–75) red.
- A score above 75 still strips attachments and converts the message to text. That text includes a link to restore the original HTML.

## [1.3.19] — 2026-09-24

### Changed

- Attachments still show a green **OK** when there are none or every file is fine.
- Beacons show **0 Blocked** in that same green. If any beacon is blocked, that count is **Blocked** in red.
- Links show the clear count as **OK** in green. Dangerous links add **NOK** in red beside it.

## [1.3.18] — 2026-09-24

### Fixed

- The footer no longer sits flush against the message. Outlook was dropping the div padding, so a 12px white spacer row is now part of the footer HTML.

## [1.3.17] — 2026-09-24

### Changed

- Attachments, beacons, and links on the footer top row use the same 12px bold type as the rest of that line.

## [1.3.16] — 2026-09-24

### Changed

- Quick Action chips show status: blocked attachments or beacons are red, a trusted sender is a lighter solid green. The 5px gap above the footer is the email background (white), not the slate of the strip.

## [1.3.15] — 2026-09-24

### Changed

- The footer is set in Segoe UI. Each green dot has a 5px-wide cell on either side, which Outlook keeps. The strip is 10% darker (`#39444d`). The second line is 1pt smaller. The action buttons are 4px taller.

## [1.3.14] — 2026-09-24

### Changed

- The footer background is a grey slate (`#3f4c55`, the previous slate at 80% over white). Facts are separated by a small green dot with 5px on each side. The second line is 2pt larger. The button row has 5px less space above and below the chips. There is 5px between the bottom of the message and the top of the footer.

## [1.3.13] — 2026-09-24

### Added

- AES Settings → Logging has **Max log size** in MiB. The default is 0, which means unlimited. When `VBA_Log.txt` grows past the limit, the oldest lines are removed.

## [1.3.12] — 2026-09-24

### Changed

- The AES footer strip is redesigned: one flat dark-slate background (`#0f1f2a`), white heading, muted detail line, and hairline `|` dividers instead of the green-to-red gradient. Risk shows as a small coloured dot beside the label.
- Quick Action chips are white pills with slate letters so they lift off the strip. Blocked attachments/beacons fill red; a trusted sender fills green.

## [1.3.11] — 2026-09-24

### Changed

- The selected tab is a shade darker than the window background.

## [1.3.10] — 2026-09-24

### Changed

- The email timeline draws the day of the month behind the hours at 20% opacity and 72pt.
- The grid line between 23:00 and 00:00 is a solid dark bar.

## [1.3.9] — 2026-09-24

### Changed

- **Quick Actions:** uses near-black or white, whichever contrasts with the scan-strip colour. White on the low-risk green was washing out.

## [1.3.8] — 2026-09-24

### Changed

- Footer Quick Action chips (SL, BA, BB, TS, NT) sit together in the centre of the strip, with 20px between each chip.

## [1.3.7] — 2026-09-24

### Changed

- The top bar shows a centred live clock, for example **Now: 24 September 2026; 20:08:56**.
- **PostgreSQL** and **Database…** moved from the top bar into the Settings tab.

## [1.3.6] — 2026-09-24

### Changed

- Welcome no longer has its own Refresh buttons. One **Refresh** sits at the top right. The pill beside it is **Push** (AES tells GURI as soon as a scan finishes) or **Pull** (GURI polls Outlook on a timer).
- **Settings** is the rightmost tab, next to Ollama.

## [1.3.5] — 2026-09-24

### Fixed

- The suite tray icon was registered as GURI under Python, so it did not appear as AGS in Settings > Personalisation > Taskbar > Other system tray icons. It now registers as **AGS (Aliniant Geosense Suite)** and turns that entry On.

## [1.3.4] — 2026-09-24

### Fixed

- Welcome Refresh (timeline, matching emails, actions, statistics) ignored the click: Qt passed a bool into handlers that did not accept it, so the dashboard never reloaded. Refresh now redraws and pulls a fresh Outlook scrape.

## [1.3.3] — 2026-09-24

### Changed

- Quick Action chips are about three times as wide. "Quick Actions:" sits on the left of that row; the label and the chips share the footer width in equal slots so the row stays centred.

## [1.3.2] — 2026-09-24

### Added

- Footer third row, **Quick Actions**: Show Links `<SL>`, Block Attachments `<BA>`, Block Beacons `<BB>`, Trust Sender `<TS>`, Mark Not Trusted `<NT>`. Chips are white on green. `<BA>` / `<BB>` turn red when that block is already on. A trusted sender shows `<ST>` in green instead of `<TS>`.

### Fixed

- A reply chain no longer keeps every earlier AES scan. Previous results are removed and one scan result is placed at the bottom of the message.

## [1.3.1] — 2026-09-24

### Added

- Master installer `scripts/Install_GeoFooter.bat` (runs `scripts/Install-GeoFooter.ps1`) installs both the AES email tool (venv, `aes://`, Outlook VBA, ribbon) and the GURI GUI (same venv, desktop shortcut, `Launch_GURI_GUI.bat` uses `.venv`). It tells you PostgreSQL is required and copies `guri_postgres_config.example.json` when the real config is missing.

## [1.3.0] — 2026-09-24

### Added

- Aura automation (`aura/automation.py`): creates an Outlook draft removal email for every catalogue broker with an opt-out address, plus AES-detected brokers. Drafts are saved to Drafts (category "Aura") and are **never sent automatically**.
- Aura tab **Review** sub-tab: Open draft, Regenerate, Skip, Prepare all now. Aura runs on open and every 5 minutes, and marks requests as sent (30-day follow-up) when the email shows up in Sent Items.
- "Send from" Outlook account in the Aura identity form (`aura/settings.py`, `aura_settings.json`).
- AES scans auto-queue detected broker mail to Aura, but only mail dated on or after the Aura start date (default 2026-09-24).
- `draft_entry_id` column on Aura requests (auto-migrated).
- `tests/test_aura_automation.py` (fake draft backend, no Outlook needed); runs in Conda CI.

### Changed

- Web-form-only brokers are tracked as "Manual web form (not automated)" and are not submitted.

## [1.2.10] — 2026-09-24

### Fixed

- Conda CI smoke tests failed with `ModuleNotFoundError` for `geofooter` / `aes` / `guri` / `aura`: added `tests/conftest.py` to put the install root on `sys.path`.

## [1.2.9] — 2026-09-24

### Fixed

- Conda CI flake8 F821: `scripts/fix_duplicate_guris.py` used `os` without importing it (blocked every push since the workflow started linting `scripts/`). Path now uses `geofooter.paths.datastore_path` instead of a hard-coded `C:\GeoFooter`.

## [1.2.8] — 2026-09-24

### Fixed

- `MSCANSelfUpdate`: no longer tells you to set AccessVBOM when it is already 1. Nudges Alt+F11 and retries; if `Application.VBE` is still Nothing, opens the VBA folder and pastes a remove/import checklist (File > Import does not need VBOM access).

## [1.2.7] — 2026-09-24

### Fixed

- VBA modules MSCANModule1.bas and MSCANToolbar.bas were LF-only; Outlook File > Import needs CRLF. Added `.gitattributes` (`*.bas` / `*.cls` eol=crlf).

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
