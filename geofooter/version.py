#!/usr/bin/env python3
"""GeoFooter / AES / GURI suite version and release metadata.

Single source of truth for About UI, window titles, footers, ribbon host,
and packaging. Bump VERSION + VERSION_STAMP + RELEASE_NOTES together when
shipping a meaningful change; mirror the same version in:
  - VERSION (plain text, install root)
  - CHANGELOG.md
  - AesRibbonHost/AesRibbonHost.csproj <Version>

SemVer (MAJOR.MINOR.PATCH) — almost everything is a PATCH:
  PATCH (~90%)  bugfixes, wording, icons, scoring tweaks, docs
  MINOR         new backward-compatible features (tabs, ribbon buttons, modes)
  MAJOR         rare — breaking Outlook/VBA contracts, schema, or public CLI
"""

from __future__ import annotations

# Semver: MAJOR.MINOR.PATCH
VERSION = "1.5.57"
VERSION_TUPLE = (1, 5, 57)
# Local release / build timestamp for this VERSION (YYYY-MM-DD HH:MM)
VERSION_STAMP = "2026-10-22 19:51"

APP_NAME = "GURI"
APP_FULL_NAME = "GURI - Aliniant Smart Ass Email"
# The "Smart Ass" joke: Smart Ass Email ≈ Aliniant Smart Scraper
APP_SILLY_QUOTE = '"Aliniant Smart Scraper"'
APP_ORG = "Aliniant"
SUITE_NAME = "GeoFooter"
SUITE_PRODUCTS = ("AES", "GURI", "Aura", "AesRibbonHost")
COPYRIGHT_YEAR = "2026"

# Short release notes for the About tab (newest first).
# Each entry: (version, timestamp, notes)
RELEASE_NOTES = (
    (
        "1.5.57",
        "2026-10-22 19:51",
        "AES Diagnostics uses a two-column layout: Live Controls | tests and detail.",
    ),
    (
        "1.5.56",
        "2026-10-22 18:22",
        "CatchUp no longer reads HTMLBody; idle sweeps back off; header-strip spam stopped.",
    ),
    (
        "1.5.55",
        "2026-10-22 18:13",
        "Diagnostics can start an Outlook lockup scanner that auto-captures freeze reports.",
    ),
    (
        "1.5.54",
        "2026-10-21 18:40",
        "Footer chips have hover tips; AURA opens GURI on that sender's broker.",
    ),
    (
        "1.5.53",
        "2026-10-21 18:23",
        "Footer quip is vertically centred in the top row.",
    ),
    (
        "1.5.52",
        "2026-10-21 18:09",
        "Footer rows centre on the whole strip; AURA lettering centred in its chip.",
    ),
    (
        "1.5.51",
        "2026-10-21 17:50",
        "AURA chip has a single border and larger rainbow lettering.",
    ),
    (
        "1.5.50",
        "2026-10-21 17:37",
        "Footer row 2 has an AURA chip; A/E/S sit in one column so no white seam under E.",
    ),
    (
        "1.5.49",
        "2026-10-21 17:32",
        "Footer rule no longer draws a white hairline through the E letter.",
    ),
    (
        "1.5.48",
        "2026-10-21 17:16",
        "Footer top-row title and metrics share one vertical centre; white line under E gone.",
    ),
    (
        "1.5.47",
        "2026-10-21 19:00",
        "QUICK ACTIONS is beside the chips again (undo stacked label).",
    ),
    (
        "1.5.46",
        "2026-10-21 18:45",
        "Rescans no longer re-attach footer images when they are already present (stops text-only flatten).",
    ),
    (
        "1.5.45",
        "2026-10-21 18:30",
        "Footer centres each row; Quick Actions sits above the centred chips.",
    ),
    (
        "1.5.44",
        "2026-10-21 18:20",
        "Footer A/E letter cells no longer show white hairlines.",
    ),
    (
        "1.5.43",
        "2026-10-21 18:10",
        "Footer bottom row is 2px taller under the quick actions.",
    ),
    (
        "1.5.42",
        "2026-10-21 18:00",
        "Footer top row and orange A vertically centred; version reads v.X.Y.Z.",
    ),
    (
        "1.5.41",
        "2026-10-21 17:40",
        "Footer row text is vertically centred; copyright shows v. version.",
    ),
    (
        "1.5.40",
        "2026-10-21 17:25",
        "Footer top summary row is 5px shorter.",
    ),
    (
        "1.5.39",
        "2026-10-21 17:20",
        "Footer BB chip matches the others; copyright shows the suite version.",
    ),
    (
        "1.5.38",
        "2026-10-21 17:10",
        "Footer orange A moved 5px down.",
    ),
    (
        "1.5.37",
        "2026-10-21 17:00",
        "Footer A/E/S letters are centred on one line instead of left-aligned.",
    ),
    (
        "1.5.36",
        "2026-10-21 16:45",
        "Footer A/E/S letters are 14px and centred in their rows; the strip no longer grows taller around them.",
    ),
    (
        "1.5.35",
        "2026-10-21 16:15",
        "Footer letters: correct Content-ID on the VBA attach.",
    ),
    (
        "1.5.34",
        "2026-10-21 16:10",
        "Footer-mark job no longer rewrites mail to attach letters.",
    ),
    (
        "1.5.33",
        "2026-10-21 15:58",
        "Footer letters: follow-up attach waits for the new footer.",
    ),
    (
        "1.5.32",
        "2026-10-21 15:40",
        "Compact 62px footer strip; A/E/S letters as images, 4px from row edges.",
    ),
    (
        "1.5.31",
        "2026-10-21 15:10",
        "A/E/S fill the row with a 3px inset; strip height unchanged.",
    ),
    (
        "1.5.30",
        "2026-10-21 15:05",
        "Guard None utcoffset before total_seconds.",
    ),
    (
        "1.5.29",
        "2026-10-21 15:00",
        "Much larger A/E/S letters in the footer strip.",
    ),
    (
        "1.5.28",
        "2026-10-21 14:55",
        "Footer: read trusted inside isinstance branch.",
    ),
    (
        "1.5.27",
        "2026-10-21 14:50",
        "Footer: if/else narrow for block_state.",
    ),
    (
        "1.5.26",
        "2026-10-21 14:45",
        "Footer build: narrow block_state to a dict.",
    ),
    (
        "1.5.25",
        "2026-10-21 14:40",
        "RDAP contact: narrow to a dict before .get.",
    ),
    (
        "1.5.24",
        "2026-10-21 14:35",
        "Footer mark job: type COM helpers as Any.",
    ),
    (
        "1.5.23",
        "2026-10-21 14:30",
        "Hop decode: keep addr_decode narrowing in nested helper.",
    ),
    (
        "1.5.22",
        "2026-10-21 14:25",
        "RDAP registrant geo: narrow objects to a dict.",
    ),
    (
        "1.5.21",
        "2026-10-21 14:20",
        "Threat intel HTML: coerce dict keys to str for sort.",
    ),
    (
        "1.5.20",
        "2026-10-21 14:15",
        "Threat intel: guard None import before attribute use.",
    ),
    (
        "1.5.19",
        "2026-10-21 14:12",
        "Status banner: integer image size from text bbox.",
    ),
    (
        "1.5.18",
        "2026-10-21 14:10",
        "WHOIS lookup: fix possibly-unbound FuturesTimeout.",
    ),
    (
        "1.5.17",
        "2026-10-21 14:05",
        "Larger A/E/S letters in the footer.",
    ),
    (
        "1.5.16",
        "2026-10-21 13:50",
        "Footer summary lines are centred on the strip.",
    ),
    (
        "1.5.15",
        "2026-10-21 13:45",
        "Footer A/E/S letters line up with the strip lines, bigger and bolder.",
    ),
    (
        "1.5.14",
        "2026-10-21 13:35",
        "Footer strip banding fixed; A/E/S beside the mark in Montserrat Bold.",
    ),
    (
        "1.5.13",
        "2026-10-21 12:50",
        "GURI no longer jumps in front of Outlook unless you open it on purpose.",
    ),
    (
        "1.5.12",
        "2026-10-20 22:35",
        "aliniant.net holding image scales to the window height, keeping its proportions.",
    ),
    (
        "1.5.11",
        "2026-10-20 22:25",
        "aliniant.net holding image filename changed so browsers pick up the new picture.",
    ),
    (
        "1.5.10",
        "2026-10-20 22:20",
        "aliniant.net holding page gets the solarpunk Shire image and dark-grey panels.",
    ),
    (
        "1.5.9",
        "2026-10-20 21:55",
        "aliniant.net holding page gets its welcome text and dragon paw prints.",
    ),
    (
        "1.5.8",
        "2026-10-20 21:50",
        "Holding-page images are stored intact in the repository.",
    ),
    (
        "1.5.7",
        "2026-10-20 21:45",
        "aliniant.net now shows the Aliniant holding page.",
    ),
    (
        "1.5.6",
        "2026-10-20 21:30",
        "Read tracking is live at track.aliniant.net.",
    ),
    (
        "1.5.5",
        "2026-10-20 21:15",
        "Pixel setup script is safe to run on a server that already hosts other sites.",
    ),
    (
        "1.5.4",
        "2026-10-20 21:10",
        "Read-tracking pixels move to track.aliniant.net on a dedicated server, with a one-command setup script.",
    ),
    (
        "1.5.3",
        "2026-10-20 20:50",
        "Tracking-pixel web config can't intercept other traffic on the i-txone.com server.",
    ),
    (
        "1.5.2",
        "2026-10-20 20:40",
        "Read-tracking pixels now use track.i-txone.com instead of track.aliniant.com.",
    ),
    (
        "1.5.1",
        "2026-10-20 18:20",
        "High-risk mitigated mail has a 'Not Dangerous you big scaredy cat' link at the top that restores the original formatting, and its AES actions are clickable.",
    ),
    (
        "1.5.0",
        "2026-10-20 18:05",
        "AES read-tracking pixel on outgoing Aliniant mail, with opens shown in a new GURI Tracking tab (off until the track.aliniant.com endpoint is live).",
    ),
    (
        "1.4.30",
        "2026-10-20 17:45",
        "AES action buttons at the bottom of a full scan are no longer cut off.",
    ),
    (
        "1.4.29",
        "2026-10-20 17:35",
        "Shodan on the free plan shows 'free plan can't look up this IP' instead of 'lookup failed (HTTPError)'.",
    ),
    (
        "1.4.28",
        "2026-10-20 17:25",
        "The full analysis in the footer is always on white, even when the sender's mail has a coloured background.",
    ),
    (
        "1.4.27",
        "2026-10-20 17:00",
        "Bug tracker brought up to date (BUG-009 to BUG-014), with a standing rule to log every bug seen.",
    ),
    (
        "1.4.26",
        "2026-10-20 16:30",
        "Outlook stays usable while AES Settings is open, and a PC sleep is no longer logged as an hours-long UI stall.",
    ),
    (
        "1.4.25",
        "2026-10-20 16:15",
        "Ribbon Short Scan (and other AES buttons) no longer shows a false 'button not found' popup when the VBA macro is busy starting.",
    ),
    (
        "1.4.24",
        "2026-10-20 16:05",
        'Outlook VBA: no startup race into the queue, log trimming checked every 200 lines, and no wscript launched per clicked message.',
    ),
    (
        "1.4.23",
        "2026-10-20 17:35",
        'Flags a first-time sender claiming to reply to a thread AES never saw.',
    ),
    (
        "1.4.22",
        "2026-10-20 17:10",
        'Route Table shows TLS per hop; Full Scan adds sending client and mail type, and unsubscribe links are link-checked.',
    ),
    (
        "1.4.21",
        "2026-10-20 16:50",
        'Full Scan lists every DKIM signature, SPF and DKIM alignment with From, the DKIM key size and the published DMARC policy.',
    ),
    (
        "1.4.20",
        "2026-10-20 16:25",
        'Sending platform and tenant ID (SendGrid account, Mailgun, SES and others) in the Full Scan report, with trusted and blocked tenant lists.',
    ),
    (
        "1.4.19",
        "2026-10-20 16:05",
        'Identity checks: Reply-To, display-name and brand spoofing, and Return-Path mismatches now score and show in the Full Scan report.',
    ),
    (
        "1.4.18",
        "2026-10-20 15:45",
        "Route Table hops keep the sending IP's ASN; the receiving server no longer overwrites it.",
    ),
    (
        "1.4.17",
        "2026-10-20 00:35",
        "Create GURI waits until Recipients is finished before offering to add a new address.",
    ),
    (
        "1.4.16",
        "2026-10-19 13:10",
        "The footer strip is a slightly lighter slate.",
    ),

    (
        "1.4.15",
        "2026-10-19 13:05",
        "Mail that stays inside one domain shows the hop address as internal within that domain, "
        "and that no longer adds the outside-sender risk score.",
    ),
    (
        "1.4.14",
        "2026-10-08 10:10",
        "GURI stays at the window size you leave it. Opening it, the tray, and Outlook no longer force it full screen.",
    ),
    (
        "1.4.13",
        "2026-10-05 10:30",
        "SL, BA, BB, TS, and NT each start a Short Scan of the open message.",
    ),
    (
        "1.4.12",
        "2026-10-05 10:25",
        "A finished scan footer is written onto the message you are reading. It used to wait until you "
        "clicked away, then give up after 90 seconds.",
    ),
    (
        "1.4.11",
        "2026-10-05 10:00",
        "A footer that waits because you are reading the message is applied once you move on, "
        "and it no longer holds up every other scan.",
    ),
    (
        "1.4.10",
        "2026-10-05 10:15",
        "An open draft no longer stops AES applying footers. A failed footer save can no longer leave a mail "
        "without its body, and a rescan replaces the older pending scan of the same mail.",
    ),
    (
        "1.4.9",
        "2026-09-30 16:25",
        "AES leaves the message you are reading alone, and a ribbon button repaint can no longer rewrite a mail. "
        "Beacon blocking stops rewriting bodies it has already cleaned.",
    ),
    (
        "1.4.8",
        "2026-09-30 16:20",
        "A message carries exactly one scan footer. Older footers quoted from previous emails are removed even "
        "when their markers were stripped in transit.",
    ),
    (
        "1.4.7",
        "2026-09-29 10:12",
        "Old scan job scripts are deleted at startup. They had built up to 3290 files in the folder AES also "
        "writes the queue tick script to.",
    ),
    (
        "1.4.6",
        "2026-09-29 10:05",
        "A message that will not accept a footer is left for later instead of being rewritten up to 30 times, "
        "which is what made the mail list flicker. Health counters and scripts/aes_log_health.py report stalls, "
        "conflicts and backlogs.",
    ),
    (
        "1.4.5",
        "2026-09-25 13:15",
        "Typing in Outlook no longer freezes or drops letters: AES stops busy-waiting on the UI thread and pauses scan work while you write.",
    ),
    (
        "1.4.4",
        "2026-09-25 12:45",
        "Each link on the Link Safety page can be set to Low, Medium, or High. That rating is kept and used on later scans.",
    ),
    (
        "1.4.3",
        "2026-09-25 12:32",
        "Click a witticism to load it into the line box. Add becomes Edit and updates that line.",
    ),
    (
        "1.4.2",
        "2026-09-25 12:35",
        "Witticisms picks a Poe model from a list and can test the connection. Settings opens even when the VBA toolbar button is missing.",
    ),
    (
        "1.4.1",
        "2026-09-25 12:45",
        "The footer NOK count sits one pixel higher, and labels have a gap before their values.",
    ),
    (
        "1.4.0",
        "2026-09-25 12:40",
        "AES Settings has a Witticisms tab to add or remove footer lines, and to ask Poe for a new one.",
    ),
    (
        "1.3.44",
        "2026-09-25 12:30",
        "The footer A spans all three rows of the scan strip.",
    ),
    (
        "1.3.43",
        "2026-09-25 12:25",
        "Links, OK, and NOK each sit in their own footer cell so they share the top-row baseline.",
    ),
    (
        "1.3.42",
        "2026-09-25 12:20",
        "Footer lines are centred on the full strip. Links counts share one baseline, and the copyright uses the © symbol.",
    ),
    (
        "1.3.41",
        "2026-09-25 12:05",
        "The second footer line renders the relay note instead of raw HTML, and the Quick Actions row shows the copyright on the right.",
    ),
    (
        "1.3.40",
        "2026-09-25 11:58",
        "The footer A attaches to the message you are reading, not only to a popped-out window.",
    ),
    (
        "1.3.39",
        "2026-09-25 11:55",
        "Footer asides are longer and funnier, and they still stay on one line.",
    ),
    (
        "1.3.38",
        "2026-09-25 11:55",
        "Footer asides stay on one line. The top-row counts share one baseline. The orange A is the logo image again.",
    ),
    (
        "1.3.37",
        "2026-09-25 11:50",
        "Footer buttons are shorter and their fill stays inside the border. The orange A is drawn in the footer. A random cheeky line sits on the top right.",
    ),
    (
        "1.3.36",
        "2026-09-25 11:40",
        "A scan that dies or overruns writes a failure marker, so Outlook does not stay on AES PROC. Repeated failures pause the queue.",
    ),
    (
        "1.3.35",
        "2026-09-25 10:30",
        "Actions to take opens newest first. Column headers sort, and a second click reverses. Timeline days read 24th and 25th, paler.",
    ),
    (
        "1.3.34",
        "2026-09-25 00:20",
        "Footer top-row labels and dots share one vertical centre.",
    ),
    (
        "1.3.33",
        "2026-09-24 22:55",
        "The footer A is embedded in the message so Outlook can show it.",
    ),
    (
        "1.3.32",
        "2026-09-24 22:50",
        "The orange A sits on the left of the footer. A faint line separates each row.",
    ),
    (
        "1.3.31",
        "2026-09-24 22:45",
        "A message already converted to text stays text. Button clicks no longer write HTML and a scan at the same time.",
    ),
    (
        "1.3.30",
        "2026-09-24 22:40",
        "NT stays neutral: first click drops trust and blocks beacons. Second click is Full No Trust and converts that mail to text. Footer buttons replace the footer.",
    ),
    (
        "1.3.29",
        "2026-09-24 22:30",
        "Beacon Blocking settings are the default. A footer button on one mail overrides them for that sender.",
    ),
    (
        "1.3.28",
        "2026-09-24 22:20",
        "When beacon blocking covers the sender, detected beacons count as blocked.",
    ),
    (
        "1.3.27",
        "2026-09-24 22:18",
        "The whole HRE label is the link, not only the count.",
    ),
    (
        "1.3.26",
        "2026-09-24 22:15",
        "HRE opens the risk breakdown. Footer row padding is half of what it was.",
    ),
    (
        "1.3.25",
        "2026-09-24 22:12",
        "Scans no longer restore HTML automatically, so a mitigated mail stops flipping back to HTML.",
    ),
    (
        "1.3.24",
        "2026-09-24 22:10",
        "HRE sits to the right of Links: green at 0, yellow at 1–3, red above that.",
    ),
    (
        "1.3.23",
        "2026-09-24 22:08",
        "Footer rows are evenly spaced, with the same hairline between row 1 and 2 as between 2 and 3.",
    ),
    (
        "1.3.22",
        "2026-09-24 22:05",
        "Attachment counts use the same green or red as OK and NOK.",
    ),
    (
        "1.3.21",
        "2026-09-24 22:00",
        "BB is red when global beacon blocking covers the sender, not only the per-sender list.",
    ),
    (
        "1.3.20",
        "2026-09-24 21:55",
        "Risk label is light green, amber, or red by band. Scores above 75 become text with a restore-HTML link.",
    ),
    (
        "1.3.19",
        "2026-09-24 21:50",
        "Beacon 0 Blocked and link OK counts are green; blocked beacons and link NOK counts are red.",
    ),
    (
        "1.3.18",
        "2026-09-24 21:45",
        "A 12px white spacer row sits above the footer so Outlook cannot attach it to the message.",
    ),
    (
        "1.3.17",
        "2026-09-24 21:42",
        "Every item on the footer top row is 12px bold.",
    ),
    (
        "1.3.16",
        "2026-09-24 21:40",
        "Blocked chips are red, trusted is a lighter solid green, and the gap above the footer is email white.",
    ),
    (
        "1.3.15",
        "2026-09-24 21:35",
        "Footer uses Segoe UI, 5px cells beside each green dot, a darker slate, a smaller second line, and taller buttons.",
    ),
    (
        "1.3.14",
        "2026-09-24 21:30",
        "Footer uses a grey slate, green dots between facts, a larger second line, tighter buttons, and 5px above the strip.",
    ),
    (
        "1.3.13",
        "2026-09-24 21:25",
        "Logging settings include a max log size in MiB. 0 is unlimited; over the cap, the oldest lines are removed.",
    ),
    (
        "1.3.12",
        "2026-09-24 21:20",
        "Footer strip redesigned: one dark slate background, risk dot, hairline dividers, white pill actions.",
    ),
    (
        "1.3.11",
        "2026-09-24 21:12",
        "The selected tab is a shade darker than the window background.",
    ),
    (
        "1.3.10",
        "2026-09-24 21:10",
        "Email timeline shows the day of the month at 20% / 72pt, and a solid bar between 23:00 and 00:00.",
    ),
    (
        "1.3.9",
        "2026-09-24 21:05",
        "Quick Actions label uses dark or white text so it contrasts with the scan strip.",
    ),
    (
        "1.3.8",
        "2026-09-24 21:00",
        "Quick Action chips sit together with 20px between them.",
    ),
    (
        "1.3.7",
        "2026-09-24 20:10",
        "Centred top-bar clock. PostgreSQL and Database live on the Settings tab.",
    ),
    (
        "1.3.6",
        "2026-09-24 20:05",
        "One Refresh at the top right, a Push/Pull pill beside it, and a Settings tab after Ollama.",
    ),
    (
        "1.3.5",
        "2026-09-24 20:00",
        "Tray icon registers as AGS (Aliniant Geosense Suite) and is turned on in Windows taskbar settings.",
    ),
    (
        "1.3.4",
        "2026-09-24 19:35",
        "Welcome Refresh reloads Outlook. The click was ignored because Qt passed a bool the handler rejected.",
    ),
    (
        "1.3.3",
        "2026-09-24 19:30",
        "Quick Action chips are wider and spaced evenly, with Quick Actions: on the left of that centred row.",
    ),
    (
        "1.3.2",
        "2026-09-24 19:25",
        "Footer Quick Actions row (SL, BA, BB, TS/ST, NT). Reply chains keep one scan result at the bottom.",
    ),
    (
        "1.3.1",
        "2026-09-24 19:45",
        "Master installer scripts/Install_GeoFooter.bat. It states that PostgreSQL is required.",
    ),
    (
        "1.3.0",
        "2026-09-24 19:30",
        "Aura automation: removal emails for the whole catalogue drafted in Outlook for "
        "review (never auto-sent); AES auto-queues broker mail dated 2026-09-24 onwards; "
        "sent drafts tracked with 30-day follow-up.",
    ),
    (
        "1.2.10",
        "2026-09-24 18:45",
        "CI: tests/conftest.py puts the install root on sys.path for pytest.",
    ),
    (
        "1.2.9",
        "2026-09-24 18:40",
        "Unblock Conda CI: fix missing import os in scripts/fix_duplicate_guris.py.",
    ),
    (
        "1.2.8",
        "2026-09-24 18:30",
        "MSCANSelfUpdate: wake VBE with Alt+F11; manual import fallback when Application.VBE stays Nothing despite AccessVBOM=1.",
    ),
    (
        "1.2.7",
        "2026-09-24 18:25",
        "Force CRLF on all VBA modules; .gitattributes locks *.bas/*.cls.",
    ),
    (
        "1.2.6",
        "2026-09-24 18:20",
        "In-Outlook VBA updater (Alt+F8 UpdateAesFromDisk) for machines where external import can't reach VBE.",
    ),
    (
        "1.2.5",
        "2026-09-24 17:40",
        "VBA modules are pure ASCII (no more garbled dashes in Outlook dialogs); CI guard.",
    ),
    (
        "1.2.4",
        "2026-09-24 17:30",
        "VBA import script compiles and saves the project so Outlook restarts keep the new code.",
    ),
    (
        "1.2.3",
        "2026-09-24 17:10",
        "Conda environment.yml + CI; SemVer bump on every patch; GURI/Aura cold-start "
        "and AES toolbar recovery after VBA re-import.",
    ),
    (
        "1.2.2",
        "2026-09-24 16:35",
        "Automated Outlook VBA sync (scripts/Import_VBA_to_Outlook); docs bug tracker.",
    ),
    (
        "1.2.1",
        "2026-09-24 18:00",
        "Suite reorganised into aes/, guri/, aura/, geofooter/ packages; VBA\\ "
        "holds only Outlook modules; GURI lists live in datastore/.",
    ),
    (
        "1.2.0",
        "2026-09-24 15:25",
        "Aura (data-broker removal) ribbon + GURI tab; threat-intel RiskTable "
        "consensus scoring; Create-GURI DB persist (avg_risk TEXT); clickable "
        "footer metrics; IP allocation decode; suite versioning + public GitHub.",
    ),
    (
        "1.1.0",
        "2026-08-10 14:16",
        "User rules with compile + local enforcement; Locations file scrape; "
        "identity / not-for-me deadlines; timeline importance colours; "
        "scrape lookback months; About + version.py.",
    ),
    (
        "1.0.0",
        "2026-01-02 09:00",
        "GURI database viewer, Outlook scrape, deadlines, learning loop, Ollama.",
    ),
)


def version_string() -> str:
    """Human-readable version with timestamp, e.g. 1.2.0 (2026-09-24 15:25)."""
    stamp = (VERSION_STAMP or "").strip()
    if stamp:
        return f"{VERSION} ({stamp})"
    return VERSION


def suite_banner() -> str:
    """One-line product + version for logs / footers."""
    return f"{SUITE_NAME} / {'+'.join(SUITE_PRODUCTS)} v{version_string()}"


def about_summary() -> str:
    return (
        f"{APP_FULL_NAME}  {APP_SILLY_QUOTE}\n"
        f"Suite {SUITE_NAME}  v{version_string()}\n"
        f"Products: {', '.join(SUITE_PRODUCTS)}\n"
        f"© {COPYRIGHT_YEAR} {APP_ORG}"
    )
