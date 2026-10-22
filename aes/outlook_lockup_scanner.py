#!/usr/bin/env python3
"""Out-of-process Outlook UI lockup scanner.

Outlook freezes on its UI thread, so VBA cannot log them. This watcher runs
outside Outlook and probes the main window with SendMessageTimeout. On hang it
writes a freeze report (process snapshot + log tails) under
%LOCALAPPDATA%\\GeoFooter\\Logs\\lockup_scans\\.

Start / stop from AES Diagnostics, or:

  pythonw aes\\outlook_lockup_scanner.py --start
  pythonw aes\\outlook_lockup_scanner.py --stop
  pythonw aes\\outlook_lockup_scanner.py --status
"""

from __future__ import annotations

import argparse
import json
import os
import sys
import time
from datetime import datetime, timezone
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple

_PKG_DIR = Path(__file__).resolve().parent
sys.path[:] = [p for p in sys.path if Path(p or ".").resolve() != _PKG_DIR]
if str(_PKG_DIR.parent) not in sys.path:
    sys.path.insert(0, str(_PKG_DIR.parent))

# Defaults: probe every 2s; treat no reply within 2.5s as hung; require two
# consecutive hung probes before declaring a freeze (avoids one-off spikes).
DEFAULT_INTERVAL_SEC = 2.0
DEFAULT_TIMEOUT_MS = 2500
DEFAULT_CONFIRM = 2

WM_NULL = 0x0000
SMTO_ABORTIFHUNG = 0x0002
SMTO_BLOCK = 0x0001


def _user_logs() -> Path:
    try:
        from geofooter.paths import user_data_dir

        root = user_data_dir() / "Logs"
    except Exception:
        root = Path(os.environ.get("LOCALAPPDATA", "")) / "GeoFooter" / "Logs"
    root.mkdir(parents=True, exist_ok=True)
    return root


def _scan_dir() -> Path:
    d = _user_logs() / "lockup_scans"
    d.mkdir(parents=True, exist_ok=True)
    return d


def _status_path() -> Path:
    return _user_logs() / "outlook_lockup_scanner_status.json"


def _pid_path() -> Path:
    return _user_logs() / "outlook_lockup_scanner.pid"


def _event_log_path() -> Path:
    return _user_logs() / "outlook_lockup_events.log"


def _write_status(payload: Dict[str, Any]) -> None:
    path = _status_path()
    tmp = path.with_suffix(".tmp")
    tmp.write_text(json.dumps(payload, indent=2), encoding="utf-8")
    os.replace(tmp, path)


def _read_status() -> Dict[str, Any]:
    path = _status_path()
    if not path.is_file():
        return {"running": False}
    try:
        data = json.loads(path.read_text(encoding="utf-8"))
        return data if isinstance(data, dict) else {"running": False}
    except Exception:
        return {"running": False}


def _append_event(line: str) -> None:
    path = _event_log_path()
    with path.open("a", encoding="utf-8") as fh:
        fh.write(line.rstrip() + "\n")


def _tail(path: Path, max_bytes: int = 48_000) -> str:
    if not path.is_file():
        return f"(missing {path})"
    try:
        size = path.stat().st_size
        with path.open("rb") as fh:
            if size > max_bytes:
                fh.seek(size - max_bytes)
                data = fh.read()
            else:
                data = fh.read()
        text = data.decode("utf-8", errors="replace")
        if size > max_bytes:
            nl = text.find("\n")
            if nl >= 0:
                text = text[nl + 1 :]
        return text
    except Exception as exc:
        return f"(read failed: {exc})"


def _outlook_pids() -> List[int]:
    try:
        import psutil  # type: ignore
    except Exception:
        psutil = None  # type: ignore
    pids: List[int] = []
    if psutil is not None:
        for proc in psutil.process_iter(["pid", "name"]):
            try:
                name = (proc.info.get("name") or "").lower()
            except Exception:
                continue
            if name in {"outlook.exe", "outlook"}:
                pids.append(int(proc.info["pid"]))
        return pids
    # Fallback without psutil: tasklist
    import subprocess

    try:
        out = subprocess.check_output(
            ["tasklist", "/FI", "IMAGENAME eq OUTLOOK.EXE", "/FO", "CSV", "/NH"],
            creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0),
            text=True,
            errors="replace",
        )
    except Exception:
        return []
    for line in out.splitlines():
        parts = [p.strip().strip('"') for p in line.split(",")]
        if len(parts) >= 2 and parts[0].upper().startswith("OUTLOOK"):
            try:
                pids.append(int(parts[1]))
            except ValueError:
                pass
    return pids


def _find_outlook_hwnd(pid: int) -> int:
    """Largest visible top-level window owned by Outlook."""
    import ctypes
    from ctypes import wintypes

    user32 = ctypes.windll.user32
    best = 0
    best_area = -1

    @ctypes.WINFUNCTYPE(wintypes.BOOL, wintypes.HWND, wintypes.LPARAM)
    def enum_proc(hwnd: int, _lparam: int) -> bool:
        nonlocal best, best_area
        if not user32.IsWindowVisible(hwnd):
            return True
        proc_id = wintypes.DWORD()
        user32.GetWindowThreadProcessId(hwnd, ctypes.byref(proc_id))
        if int(proc_id.value) != pid:
            return True
        # Skip tiny tool windows / owned popups without a title when possible.
        rect = wintypes.RECT()
        if not user32.GetWindowRect(hwnd, ctypes.byref(rect)):
            return True
        w = max(0, rect.right - rect.left)
        h = max(0, rect.bottom - rect.top)
        area = w * h
        if area < 20_000:
            return True
        if area > best_area:
            best_area = area
            best = int(hwnd)
        return True

    user32.EnumWindows(enum_proc, 0)
    return best


def _probe_hwnd(hwnd: int, timeout_ms: int) -> Tuple[bool, int]:
    """Return (responsive, elapsed_ms). False means hung or no reply."""
    import ctypes
    from ctypes import wintypes

    user32 = ctypes.windll.user32
    result = wintypes.DWORD()
    t0 = time.perf_counter()
    ok = user32.SendMessageTimeoutW(
        hwnd,
        WM_NULL,
        0,
        0,
        SMTO_ABORTIFHUNG | SMTO_BLOCK,
        int(timeout_ms),
        ctypes.byref(result),
    )
    elapsed = int((time.perf_counter() - t0) * 1000)
    return bool(ok), elapsed


def _process_snapshot() -> List[Dict[str, Any]]:
    rows: List[Dict[str, Any]] = []
    try:
        import psutil  # type: ignore
    except Exception:
        return rows
    watch = {
        "outlook",
        "python",
        "pythonw",
        "wscript",
        "cscript",
        "winword",
        "cursor",
        "claude",
    }
    for proc in psutil.process_iter(
        ["pid", "name", "cpu_percent", "memory_info", "cmdline", "num_threads"]
    ):
        try:
            name = (proc.info.get("name") or "").lower()
            base = name.replace(".exe", "")
            if base not in watch and "geofooter" not in " ".join(
                proc.info.get("cmdline") or []
            ).lower():
                # Still keep high-CPU processes.
                cpu = float(proc.info.get("cpu_percent") or 0)
                if cpu < 5:
                    continue
            mi = proc.info.get("memory_info")
            ws = int(getattr(mi, "rss", 0) or 0) // (1024 * 1024)
            cmd = " ".join(proc.info.get("cmdline") or [])[:220]
            rows.append(
                {
                    "pid": proc.info.get("pid"),
                    "name": proc.info.get("name"),
                    "cpu_pct": proc.info.get("cpu_percent"),
                    "ws_mb": ws,
                    "threads": proc.info.get("num_threads"),
                    "cmd": cmd,
                }
            )
        except Exception:
            continue
    rows.sort(key=lambda r: float(r.get("cpu_pct") or 0), reverse=True)
    return rows[:40]


def _write_freeze_report(
    *,
    pid: int,
    hwnd: int,
    probe_ms: int,
    hung_streak: int,
) -> Path:
    stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    out = _scan_dir() / f"freeze_{stamp}_pid{pid}.txt"
    logs = _user_logs()
    try:
        from geofooter.paths import get_install_root

        root = get_install_root()
    except Exception:
        root = Path()
    sections = [
        f"Outlook UI freeze detected at {datetime.now().isoformat(timespec='seconds')}",
        f"pid={pid} hwnd={hwnd} probe_ms={probe_ms} hung_streak={hung_streak}",
        "",
        "=== process snapshot (top) ===",
        json.dumps(_process_snapshot(), indent=2),
        "",
        "=== VBA_Log.txt (tail) ===",
        _tail(logs / "VBA_Log.txt"),
        "",
        "=== AesRibbonHost.log (tail) ===",
        _tail(logs / "AesRibbonHost.log"),
        "",
        "=== aes_footer_mark.log (tail) ===",
        _tail(logs / "aes_footer_mark.log"),
        "",
        "=== geolocate_debug.log (tail) ===",
        _tail(root / "debuglog" / "geolocate_debug.log") if root else "(no install root)",
        "",
        "=== outlook_lockup_events.log (tail) ===",
        _tail(_event_log_path(), max_bytes=16_000),
    ]
    out.write_text("\n".join(sections), encoding="utf-8")
    return out


def _is_pid_alive(pid: int) -> bool:
    if pid <= 0:
        return False
    try:
        import psutil  # type: ignore

        return psutil.pid_exists(pid)
    except Exception:
        pass
    try:
        os.kill(pid, 0)
        return True
    except OSError:
        return False


def cmd_stop() -> int:
    path = _pid_path()
    if not path.is_file():
        _write_status({"running": False, "stopped_at": _now()})
        print("lockup scanner: not running")
        return 0
    try:
        pid = int(path.read_text(encoding="utf-8").strip())
    except Exception:
        pid = 0
    if pid and _is_pid_alive(pid):
        try:
            import signal

            os.kill(pid, signal.SIGTERM)
        except Exception:
            try:
                import subprocess

                subprocess.run(
                    ["taskkill", "/PID", str(pid), "/F"],
                    capture_output=True,
                    creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0),
                )
            except Exception:
                pass
    try:
        path.unlink(missing_ok=True)  # type: ignore[call-arg]
    except TypeError:
        if path.is_file():
            path.unlink()
    except Exception:
        pass
    _write_status({"running": False, "stopped_at": _now(), "stopped_pid": pid})
    _append_event(f"{_now()} | STOP | scanner stopped")
    print(f"lockup scanner: stop requested (pid={pid})")
    return 0


def cmd_status() -> int:
    st = _read_status()
    print(json.dumps(st, indent=2))
    return 0


def _now() -> str:
    return datetime.now().strftime("%Y-%m-%d %H:%M:%S")


def run_loop(
    *,
    interval_sec: float,
    timeout_ms: int,
    confirm: int,
) -> int:
    _pid_path().write_text(str(os.getpid()), encoding="utf-8")
    freezes = 0
    probes = 0
    hung_streak = 0
    last_report = ""
    _append_event(
        f"{_now()} | START | pid={os.getpid()} interval={interval_sec}s "
        f"timeout={timeout_ms}ms confirm={confirm}"
    )
    _write_status(
        {
            "running": True,
            "pid": os.getpid(),
            "started_at": _now(),
            "interval_sec": interval_sec,
            "timeout_ms": timeout_ms,
            "confirm": confirm,
            "probes": 0,
            "freezes": 0,
            "last_probe_at": "",
            "last_probe_ms": 0,
            "last_state": "starting",
            "last_report": "",
            "outlook_pid": 0,
        }
    )
    try:
        while True:
            pids = _outlook_pids()
            if not pids:
                hung_streak = 0
                _write_status(
                    {
                        "running": True,
                        "pid": os.getpid(),
                        "started_at": _read_status().get("started_at"),
                        "interval_sec": interval_sec,
                        "timeout_ms": timeout_ms,
                        "confirm": confirm,
                        "probes": probes,
                        "freezes": freezes,
                        "last_probe_at": _now(),
                        "last_probe_ms": 0,
                        "last_state": "no_outlook",
                        "last_report": last_report,
                        "outlook_pid": 0,
                    }
                )
                time.sleep(interval_sec)
                continue

            pid = pids[0]
            hwnd = _find_outlook_hwnd(pid)
            if not hwnd:
                hung_streak = 0
                state = "no_hwnd"
                probe_ms = 0
                ok = True
            else:
                ok, probe_ms = _probe_hwnd(hwnd, timeout_ms)
                probes += 1
                if ok:
                    hung_streak = 0
                    state = "ok"
                else:
                    hung_streak += 1
                    state = f"hung_streak_{hung_streak}"
                    _append_event(
                        f"{_now()} | HUNG_PROBE | pid={pid} hwnd={hwnd} "
                        f"ms={probe_ms} streak={hung_streak}"
                    )

            if hung_streak >= confirm and hwnd:
                report = _write_freeze_report(
                    pid=pid,
                    hwnd=hwnd,
                    probe_ms=probe_ms,
                    hung_streak=hung_streak,
                )
                freezes += 1
                last_report = str(report)
                _append_event(
                    f"{_now()} | FREEZE | pid={pid} hwnd={hwnd} ms={probe_ms} "
                    f"report={report.name}"
                )
                state = "freeze_logged"
                hung_streak = 0  # cooldown; next freeze needs another streak

            _write_status(
                {
                    "running": True,
                    "pid": os.getpid(),
                    "started_at": _read_status().get("started_at"),
                    "interval_sec": interval_sec,
                    "timeout_ms": timeout_ms,
                    "confirm": confirm,
                    "probes": probes,
                    "freezes": freezes,
                    "last_probe_at": _now(),
                    "last_probe_ms": probe_ms,
                    "last_state": state,
                    "last_report": last_report,
                    "outlook_pid": pid,
                    "outlook_hwnd": hwnd,
                }
            )
            time.sleep(interval_sec)
    except KeyboardInterrupt:
        _append_event(f"{_now()} | STOP | keyboard interrupt")
    finally:
        try:
            _pid_path().unlink(missing_ok=True)  # type: ignore[call-arg]
        except TypeError:
            if _pid_path().is_file():
                _pid_path().unlink()
        except Exception:
            pass
        _write_status(
            {
                "running": False,
                "stopped_at": _now(),
                "probes": probes,
                "freezes": freezes,
                "last_report": last_report,
            }
        )
    return 0


def cmd_start(
    *,
    interval_sec: float,
    timeout_ms: int,
    confirm: int,
    foreground: bool,
) -> int:
    # Replace a stale PID file.
    existing = _pid_path()
    if existing.is_file():
        try:
            old = int(existing.read_text(encoding="utf-8").strip())
        except Exception:
            old = 0
        if old and _is_pid_alive(old):
            print(f"lockup scanner already running (pid={old})")
            return 0
        try:
            existing.unlink()
        except Exception:
            pass

    if foreground:
        return run_loop(
            interval_sec=interval_sec, timeout_ms=timeout_ms, confirm=confirm
        )

    py = Path(sys.executable)
    pythonw = py.with_name("pythonw.exe")
    exe = pythonw if pythonw.is_file() else py
    import subprocess

    creation = 0
    if sys.platform == "win32":
        creation = int(getattr(subprocess, "CREATE_NO_WINDOW", 0x08000000))
    proc = subprocess.Popen(
        [
            str(exe),
            str(Path(__file__).resolve()),
            "--run",
            f"--interval={interval_sec}",
            f"--timeout-ms={timeout_ms}",
            f"--confirm={confirm}",
        ],
        cwd=str(_PKG_DIR.parent),
        creationflags=creation,
        close_fds=True,
    )
    # Wait briefly for status file.
    for _ in range(20):
        time.sleep(0.1)
        st = _read_status()
        if st.get("running"):
            break
    print(f"lockup scanner started (launcher spawned pid={proc.pid})")
    return 0


def main(argv: Optional[List[str]] = None) -> int:
    parser = argparse.ArgumentParser(description="Outlook UI lockup scanner")
    g = parser.add_mutually_exclusive_group(required=True)
    g.add_argument("--start", action="store_true", help="Start in background")
    g.add_argument("--run", action="store_true", help="Run in this process")
    g.add_argument("--stop", action="store_true", help="Stop background scanner")
    g.add_argument("--status", action="store_true", help="Print status JSON")
    parser.add_argument("--interval", type=float, default=DEFAULT_INTERVAL_SEC)
    parser.add_argument("--timeout-ms", type=int, default=DEFAULT_TIMEOUT_MS)
    parser.add_argument("--confirm", type=int, default=DEFAULT_CONFIRM)
    parser.add_argument(
        "--foreground",
        action="store_true",
        help="With --start, run in this console instead of spawning",
    )
    args = parser.parse_args(argv)
    if args.stop:
        return cmd_stop()
    if args.status:
        return cmd_status()
    if args.run:
        return run_loop(
            interval_sec=args.interval,
            timeout_ms=args.timeout_ms,
            confirm=args.confirm,
        )
    return cmd_start(
        interval_sec=args.interval,
        timeout_ms=args.timeout_ms,
        confirm=args.confirm,
        foreground=args.foreground,
    )


if __name__ == "__main__":
    raise SystemExit(main())
