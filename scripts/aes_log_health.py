"""Summarise the AES VBA log into a health report.

VBA_Log.txt is a flat INFO stream that grows to tens of megabytes, so faults
like a commit-conflict storm or a queue that never drains are invisible unless
you know the exact phrase to search for. This walks the log backwards from the
end and reports the patterns that matter.

    python scripts\\aes_log_health.py                # today
    python scripts\\aes_log_health.py --session      # since the last Outlook start
    python scripts\\aes_log_health.py --hours 4
    python scripts\\aes_log_health.py --log path\\to\\VBA_Log.txt
"""

from __future__ import annotations

import argparse
import os
import re
import sys
from collections import Counter
from datetime import datetime, timedelta
from pathlib import Path

LINE_RE = re.compile(r"^(\d{4}-\d{2}-\d{2} \d{2}:\d{2}:\d{2}) \| (\w+) \| (.*)$")
SESSION_START = "MSCAN Started - Outlook Session Begin"

# Only the tail is ever interesting, and reading 39 MiB to answer "what
# happened this morning" is what made the log unusable in the first place.
DEFAULT_TAIL_BYTES = 12 * 1024 * 1024

PATTERNS: list[tuple[str, str, str]] = [
    # key, substring to match, human label
    ("conflict", "The operation cannot be performed because the message has been changed", "Commit conflicts (message changed)"),
    ("commit_fail", "CommitMailHtml attempt", "CommitMailHtml attempts that failed"),
    ("commit_giveup", "item keeps changing", "Commits abandoned (item never settled)"),
    ("stall", "HEALTH | UI stall", "UI thread stalls"),
    ("dup_footer", "HEALTH | footer count is", "Messages not carrying exactly one footer"),
    ("collapse", "footers found; collapsing", "Duplicate footers collapsed"),
    ("storm", "HEALTH | commit conflict storm", "Conflict storms detected"),
    ("backlog", "HEALTH | queue backlog", "Queue backlog warnings"),
    ("process", "ProcessQueue: Started", "Queue drain passes"),
    ("compose", "compose", "Compose-related deferrals"),
    ("footer_ok", "AES footer replaced", "Footers applied"),
    ("footer_fail", "footer NOT applied", "Footers abandoned"),
    ("tick_fail", "could not write tick script", "Tick script write failures"),
    ("error", "error:", "Logged errors"),
]

def safe(text: str) -> str:
    """Log lines carry subjects in whatever encoding the store used, so drop
    anything the console cannot render rather than dying on it."""
    encoding = getattr(sys.stdout, "encoding", None) or "ascii"
    return text.encode(encoding, errors="replace").decode(encoding, errors="replace")


QUEUE_SIZE_RE = re.compile(r"Queue size: (\d+)")
STALL_MS_RE = re.compile(r"held the Outlook thread for (\d+)ms")
PROCESS_RE = re.compile(r"ProcessQueue: Started with (\d+) items")


def default_log_path() -> Path:
    local = os.environ.get("LOCALAPPDATA", "")
    candidates = [
        Path(local) / "GeoFooter" / "Logs" / "VBA_Log.txt",
        Path(os.environ.get("GEOFOOTER_ROOT", "")) / "Logs" / "VBA_Log.txt",
    ]
    for path in candidates:
        if path.is_file():
            return path
    return candidates[0]


def read_tail(path: Path, max_bytes: int) -> list[str]:
    size = path.stat().st_size
    with path.open("r", encoding="utf-8", errors="replace") as handle:
        if size > max_bytes:
            handle.seek(size - max_bytes)
            handle.readline()  # discard the partial line
        return handle.read().splitlines()


def parse(lines: list[str]) -> list[tuple[datetime, str, str]]:
    parsed: list[tuple[datetime, str, str]] = []
    for line in lines:
        match = LINE_RE.match(line)
        if not match:
            continue
        try:
            stamp = datetime.strptime(match.group(1), "%Y-%m-%d %H:%M:%S")
        except ValueError:
            continue
        parsed.append((stamp, match.group(2), match.group(3)))
    return parsed


def select_window(rows: list[tuple[datetime, str, str]], args) -> tuple[list, str]:
    if args.session:
        for index in range(len(rows) - 1, -1, -1):
            if SESSION_START in rows[index][2]:
                return rows[index:], f"since Outlook started at {rows[index][0]:%Y-%m-%d %H:%M:%S}"
        return rows, "whole log (no session start found in the tail)"
    if args.hours:
        cutoff = rows[-1][0] - timedelta(hours=args.hours) if rows else datetime.now()
        return [r for r in rows if r[0] >= cutoff], f"last {args.hours}h"
    if args.all:
        return rows, "whole log tail"
    today = rows[-1][0].date() if rows else datetime.now().date()
    return [r for r in rows if r[0].date() == today], f"{today:%Y-%m-%d}"


def summarise(rows: list[tuple[datetime, str, str]]) -> None:
    counts: Counter[str] = Counter()
    stall_ms: list[int] = []
    queue_sizes: list[int] = []
    drain_sizes: list[int] = []
    subjects: Counter[str] = Counter()
    warnings: list[str] = []

    for stamp, level, msg in rows:
        for key, needle, _label in PATTERNS:
            if needle in msg:
                counts[key] += 1
        match = STALL_MS_RE.search(msg)
        if match:
            stall_ms.append(int(match.group(1)))
        match = QUEUE_SIZE_RE.search(msg)
        if match:
            queue_sizes.append(int(match.group(1)))
        match = PROCESS_RE.search(msg)
        if match:
            drain_sizes.append(int(match.group(1)))
        if msg.startswith("HEALTH |") or level in {"WARN", "ERROR"}:
            warnings.append(f"  {stamp:%H:%M:%S} {level:<5} {safe(msg[:150])}")
        if "CommitMailHtml failed for subject:" in msg:
            subjects[msg.split("subject:", 1)[1].strip()[:70]] += 1

    print(f"Lines examined : {len(rows)}")
    if rows:
        print(f"Time range     : {rows[0][0]:%H:%M:%S} -> {rows[-1][0]:%H:%M:%S}")
    print()

    print("Counts")
    for key, _needle, label in PATTERNS:
        if counts[key]:
            print(f"  {label:<40} {counts[key]}")
    if not any(counts.values()):
        print("  (nothing notable)")
    print()

    if queue_sizes or drain_sizes:
        print("Queue")
        if queue_sizes:
            print(f"  Peak queue size observed               {max(queue_sizes)}")
        if drain_sizes:
            print(f"  Drain passes                           {len(drain_sizes)} (largest {max(drain_sizes)} items)")
        print()

    if stall_ms:
        stall_ms.sort()
        print("UI thread stalls")
        print(f"  Count                                  {len(stall_ms)}")
        print(f"  Worst                                  {stall_ms[-1]}ms")
        print(f"  Median                                 {stall_ms[len(stall_ms) // 2]}ms")
        print(f"  Total time the UI was blocked          {sum(stall_ms) / 1000:.1f}s")
        print()

    if subjects:
        print("Messages that would not accept a footer")
        for subject, hits in subjects.most_common(10):
            print(f"  {hits:>3}x  {safe(subject)}")
        print()

    verdict(counts, stall_ms, queue_sizes)

    if warnings:
        print()
        print(f"Health and warning lines (last {min(len(warnings), 40)} of {len(warnings)})")
        for line in warnings[-40:]:
            print(line)


def job_script_report() -> tuple[int, int]:
    """Count the one-shot aes_geo_job_*.vbs scripts. They are written per scan
    into the same folder as the queue tick script and were never cleaned up."""
    folder = Path(os.environ.get("LOCALAPPDATA", "")) / "GeoFooter"
    if not folder.is_dir():
        return 0, 0
    cutoff = datetime.now() - timedelta(days=2)
    total = 0
    stale = 0
    for path in folder.glob("aes_geo_job_*.vbs"):
        total += 1
        try:
            if datetime.fromtimestamp(path.stat().st_mtime) < cutoff:
                stale += 1
        except OSError:
            continue
    return total, stale


def verdict(counts: Counter[str], stall_ms: list[int], queue_sizes: list[int]) -> None:
    findings: list[str] = []
    if counts["conflict"] >= 5:
        findings.append(
            f"{counts['conflict']} commit conflicts. Each failed attempt rewrites HTMLBody and "
            "repaints the mail, which is seen as the reading pane flickering."
        )
    if queue_sizes and max(queue_sizes) >= 25:
        findings.append(
            f"Queue reached {max(queue_sizes)} items. A backlog that deep keeps rewriting mail "
            "for several minutes after Outlook starts."
        )
    if stall_ms and max(stall_ms) >= 1200:
        findings.append(
            f"Longest UI stall was {max(stall_ms)}ms. Anything over ~1s on the Outlook thread "
            "drops keystrokes while composing."
        )
    if counts["tick_fail"]:
        findings.append(f"{counts['tick_fail']} tick scripts could not be written; delayed work may never run.")
    if counts["dup_footer"] or counts["collapse"]:
        findings.append(
            f"{counts['dup_footer'] + counts['collapse']} messages did not carry exactly one footer. "
            "Old footers quoted in a reply chain are surviving removal."
        )
    if counts["footer_fail"]:
        findings.append(f"{counts['footer_fail']} messages ended up with no footer.")

    total_jobs, stale_jobs = job_script_report()
    if stale_jobs >= 200:
        findings.append(
            f"{total_jobs} job scripts in %LOCALAPPDATA%\\GeoFooter ({stale_jobs} older than 2 days). "
            "The queue tick script is written to the same folder, so this can make that write fail."
        )

    print("Verdict")
    if findings:
        for item in findings:
            print(f"  - {item}")
    else:
        print("  Nothing above the warning thresholds in this window.")


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description="Summarise the AES VBA log.")
    parser.add_argument("--log", type=Path, default=None, help="Path to VBA_Log.txt")
    parser.add_argument("--session", action="store_true", help="Only since the last Outlook start")
    parser.add_argument("--hours", type=float, default=None, help="Only the last N hours")
    parser.add_argument("--all", action="store_true", help="The whole tail that was read")
    parser.add_argument("--tail-mib", type=float, default=DEFAULT_TAIL_BYTES / 1048576,
                        help="How much of the end of the file to read")
    args = parser.parse_args(argv)

    path = args.log or default_log_path()
    if not path.is_file():
        print(f"Log not found: {path}", file=sys.stderr)
        return 1

    size = path.stat().st_size
    print(f"Log            : {path}")
    print(f"Size           : {size / 1048576:.1f} MiB")

    rows = parse(read_tail(path, int(args.tail_mib * 1048576)))
    if not rows:
        print("No parseable log lines found.", file=sys.stderr)
        return 1

    window, label = select_window(rows, args)
    print(f"Window         : {label}")
    print()
    if not window:
        print("No lines in that window.")
        return 0
    summarise(window)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
