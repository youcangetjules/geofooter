# GeoFooter / AES / GURI — Performance log

Raw observations of **glitches, slowdowns, freezes and stutters** — recorded
**before** anyone decides whether they are bugs.

This is the front door. An entry here is a *symptom*, not a diagnosis. Most
entries will never become bugs: some are one-off machine load, some are Outlook
or Windows doing something unrelated, some turn out to be expected cost. The
point is to capture what was felt, while it is still fresh, so that a pattern
can be spotted across many small annoyances that individually look like nothing.

| File | Holds |
| ---- | ----- |
| **`performance.md`** (this file) | Symptoms as observed. No cause required. |
| `docs/bug_tracker.md` | Confirmed defects with a repro and an owner. |
| `WORKLOG.md` | What was changed and why. |
| `CHANGELOG.md` | What shipped, per release. |

## How to use

1. Add a new entry at the top of **Observations** with the next `PERF-NNN` id.
2. Record it even if you are not sure it is real. Uncertainty is expected here.
3. Do **not** try to diagnose while logging. Write what happened, not why.
4. Attach evidence if it is cheap to get (see **Evidence** below). If not, log
   the observation anyway — a vague entry beats a missing one.
5. Review the open entries periodically. When several describe the same thing,
   or one reproduces on demand, promote it to `docs/bug_tracker.md` as a
   `BUG-NNN` and set this entry's status to `promoted` with the bug id.

### Status

`new` · `watching` · `promoted` · `explained` · `not-reproducible` · `wontfix`

### Impact

| Level | Meaning |
| ----- | ------- |
| **P1** | Outlook or GURI unusable — hang, freeze, dropped keystrokes |
| **P2** | Clearly disruptive — long stall, flicker, must wait or click away |
| **P3** | Noticeable but tolerable — sluggish, brief stutter |
| **P4** | Barely perceptible, or only visible in logs / timings |

### Evidence

Cheap things worth grabbing while the symptom is fresh:

- `python scripts/aes_log_health.py --session` — session counters, UI stalls,
  commit conflicts, queue depth, worst stall in ms.
- `HEALTH` lines in `%LOCALAPPDATA%\GeoFooter\Logs\VBA_Log.txt`.
- A screen recording, if it is visual (flicker, repaint, cursor jump).
- What else was running: a scan in flight, Outlook starting, a big send/receive,
  GURI refreshing, another application pinning the CPU.

### Entry template

```text
### PERF-NNN — one-line symptom as experienced

- **Impact:** P1-P4
- **Area:** AES | VBA | GURI | Aura | Ribbon | Install
- **Status:** new
- **Observed:** YYYY-MM-DD (and how often: once / daily / every scan)
- **What happened:** plain description from the user's point of view.
- **What else was going on:** concurrent activity, if known.
- **Evidence:** log excerpt, analyzer output, recording, or "none".
- **Outcome:** filled in later — promoted to BUG-NNN, explained, or dropped.
```

---

## Observations

### PERF-002 — Outlook unresponsive / very slow (evening session)

- **Impact:** P2
- **Area:** AES / VBA / Outlook
- **Status:** watching
- **Observed:** 2026-10-21 ~19:35
- **What happened:** Outlook felt unresponsive or very slow again.
- **What else was going on:** AES queue idle (pending 0); last scans finished ~19:13. Catch-up heartbeat every ~90s across 5 inboxes (~58 CatchUp log lines in the prior hour). Two GURI `gui.py --raise` processes (PIDs 9684 and 33548, both started 18:44, parents under Outlook). Claude Desktop and Cursor also on CPU. Workflow Manager `main.py` ~577 MB but quiet in the sample.
- **Evidence:** At check time Outlook was mostly idle (~0.8% of one core) then spiked to ~16% of one core in a 5s sample; ~12,010 COM/UI handles, 152 threads, 450–480 MB. HEALTH: conflicts=0, uiStalls=0, worstStall=0ms, uiBusy=13s since 16:11, queuePeak=5, 30 commits. Disk idle; 31 GB RAM free. No AES conflict storm.
- **Outcome:** Not an AES scan backlog. Likely mix of (1) Outlook UI-thread CatchUp Restrict walks, (2) two GURI instances, (3) other desktop apps (Claude/Cursor). Watch whether closing the duplicate GURI and/or restarting Outlook clears the feel; if uiBusy climbs without uiStalls, CatchUp cost may need a bug.
- **Update 2026-10-22:** 1.5.55 instruments CatchUp with BeginOp/EndOp (uiStalls will count long Restrict walks after VBA re-import). Diagnostics can start an out-of-process lockup scanner that writes freeze reports when Outlook stops answering WM_NULL — use that for the next hard freeze (VBA cannot log while hung).
- **Update 2026-10-22 (cleanse):** VBA_Log showed ~9k CatchUp sweeps (~every 90s, 5 inboxes) and 8 `CompleteAsyncFooter` UI stalls (1.4–7.2s). Promoted CatchUp cost to BUG-036 (fixed 1.5.56: no HTMLBody on CatchUp, idle backoff 240s); footer stalls logged as BUG-038.
- **Update 2026-10-28:** Freeze still happens with VBA project empty (no MSCAN modules). AES VBA is not required for the current lockup. Check COM add-ins (AesRibbonHost, Google Sync/Drive, Acrobat) and other desktop Python load.
- **Update 2026-10-29 09:38:** Soft lock again while typing felt frozen (`freeze_live_20261029_093846.txt`). Windows still says not hung. Same long-lived Outlook (~25h, ~18k handles, many open compose/Word editors). Scanner separately logged hard hung streaks earlier the same morning (~08:34–08:42). Promoted pattern remains under BUG-035.

### PERF-001 — Machine slowed to a crawl; GURI suspected

- **Impact:** P2
- **Area:** GURI
- **Status:** explained (not GURI)
- **Observed:** 2026-10-21 ~17:44 (once so far)
- **What happened:** whole machine felt very slow; user suspected GURI (`guri\gui.py --raise`, PID 5220, running since 15:05, window title still v1.5.34).
- **What else was going on:** Cursor (2.1 GB, up to ~176% of one core), Lumen Workflow Manager `main.py` (PID 62812, 17–54% of one core, ~475 MB), an AES compact scan in flight (~38% of one core for ~40 s), Outlook ~22% of one core.
- **Evidence:** 20 s sample on 12 cores: GURI 0.6% of the machine (7.6% of one core), 149 MB working set, 6 threads, window responding. 430 CPU-seconds total over 2h40m (about 4.5% of one core on average). Total CPU 52–61%; disk idle; 30 GB RAM free. AES scans every few minutes, no storm.
- **Outcome:** GURI ruled out for this episode. The load was Cursor plus Workflow Manager plus one AES scan. Watch for repeats; if GURI ever shows sustained CPU, capture the same sample.

---

## Closed

Entries that were explained, dropped, or promoted to a bug. Keep them — a
symptom that recurs after being closed is a strong signal.

*No entries yet.*
