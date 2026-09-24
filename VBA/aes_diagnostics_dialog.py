#!/usr/bin/env python3
"""AES interactive diagnostics dialog.

CLI:
  pythonw aes_diagnostics_dialog.py --context context.json

Runs individual checks in-process. Outlook/COM tests use the running Outlook
instance when pywin32 is available. Classification tests mirror MSCANCore.
"""

from __future__ import annotations

import argparse
import json
import os
import platform
import subprocess
import sys
import time
from dataclasses import dataclass, field
from datetime import datetime
from pathlib import Path
from typing import Any, Callable, Dict, List, Optional, Tuple

STYLESHEET = """
QDialog { background: #f5f8fa; }
QLabel#title { font-size: 16px; font-weight: 700; color: #0f6b7c; }
QLabel#subtitle { color: #4a6570; font-size: 12px; }
QLabel#section { font-size: 12px; font-weight: 700; color: #0f6b7c; }
QLabel#hint { color: #5a7280; font-size: 11px; }
QFrame#row {
    background: #ffffff;
    border: 1px solid #d5dee3;
    border-radius: 6px;
}
QFrame#row:hover { border-color: #0f6b7c; }
QPushButton {
    background: #0f6b7c;
    color: white;
    border: none;
    border-radius: 5px;
    padding: 6px 12px;
    font-size: 12px;
}
QPushButton:hover { background: #0c5663; }
QPushButton:disabled { background: #9bb4bb; }
QPushButton#secondary {
    background: #ffffff;
    color: #0f6b7c;
    border: 1px solid #0f6b7c;
}
QPushButton#secondary:hover { background: #e8f3f5; }
QTextEdit {
    background: #ffffff;
    border: 1px solid #d5dee3;
    border-radius: 6px;
    font-family: Consolas, "Courier New", monospace;
    font-size: 11px;
    color: #1a2e35;
}
QLabel#statusPass { color: #1b7a3d; font-weight: 700; }
QLabel#statusFail { color: #b42318; font-weight: 700; }
QLabel#statusWarn { color: #b54708; font-weight: 700; }
QLabel#statusInfo { color: #175cd3; font-weight: 700; }
QLabel#statusIdle { color: #667085; font-weight: 600; }
QProgressBar {
    border: 1px solid #d5dee3;
    border-radius: 4px;
    text-align: center;
    background: #ffffff;
    height: 14px;
}
QProgressBar::chunk { background: #0f6b7c; border-radius: 3px; }
"""


@dataclass
class TestResult:
    status: str  # pass | fail | warn | info
    summary: str
    detail: str = ""


@dataclass
class TestDef:
    id: str
    name: str
    group: str
    description: str
    runner: Callable[["DiagContext"], TestResult]


@dataclass
class DiagContext:
    raw: Dict[str, Any] = field(default_factory=dict)

    @property
    def tags(self) -> List[str]:
        tags = self.raw.get("tags") or []
        return [str(t) for t in tags if str(t).strip()]

    @property
    def service_enabled(self) -> bool:
        return bool(self.raw.get("service_enabled", False))

    @property
    def watcher_count(self) -> int:
        try:
            return int(self.raw.get("watcher_count", 0))
        except (TypeError, ValueError):
            return 0

    @property
    def log_path(self) -> str:
        return str(self.raw.get("log_path") or "")

    @property
    def logging_cfg(self) -> Dict[str, Any]:
        raw = self.raw.get("logging")
        if isinstance(raw, dict):
            levels = raw.get("levels") if isinstance(raw.get("levels"), dict) else {}
            return {
                "enabled": bool(raw.get("enabled", True)),
                "levels": {
                    "info": bool(levels.get("info", True)),
                    "audit": bool(levels.get("audit", True)),
                    "warn": bool(levels.get("warn", True)),
                    "debug": bool(levels.get("debug", False)),
                },
                "capture_active": bool(raw.get("capture_active", False)),
                "capture_path": str(raw.get("capture_path") or ""),
            }
        return {
            "enabled": True,
            "levels": {"info": True, "audit": True, "warn": True, "debug": False},
            "capture_active": False,
            "capture_path": "",
        }

    @property
    def accounts_summary(self) -> str:
        return str(self.raw.get("accounts_summary") or "")

    @property
    def auto_scan_summary(self) -> str:
        return str(self.raw.get("auto_scan_summary") or "")

    @property
    def recent_auto_scan_log(self) -> str:
        return str(self.raw.get("recent_auto_scan_log") or "")

    @property
    def queue_size(self) -> int:
        try:
            return int(self.raw.get("queue_size", 0))
        except (TypeError, ValueError):
            return 0

    @property
    def inflight_scans(self) -> int:
        try:
            return int(self.raw.get("inflight_scans", 0))
        except (TypeError, ValueError):
            return 0

    @property
    def startup_quiet(self) -> bool:
        return bool(self.raw.get("startup_quiet", False))


def _exists(path: str) -> bool:
    try:
        return bool(path) and Path(path).exists()
    except OSError:
        return False


def detect_classification(subject: str, tags: List[str]) -> str:
    if not subject.strip():
        return ""
    upper = subject.upper()
    for tag in tags:
        if tag.upper() in upper:
            return tag
    if "[NR/E]" in upper:
        return "[NR/E]"
    start = upper.find("[SEC")
    if start < 0:
        return ""
    end = upper.find("]", start)
    if end < 0:
        return ""
    return subject[start : end + 1]


def remove_classification(subject: str, tags: List[str]) -> str:
    result = subject
    for tag in tags:
        result = result.replace(tag, "")
        # case-insensitive remove
        idx = result.upper().find(tag.upper())
        while idx >= 0:
            result = result[:idx] + result[idx + len(tag) :]
            idx = result.upper().find(tag.upper())
    existing = detect_classification(result, tags)
    if existing:
        result = result.replace(existing, "")
    while "  " in result:
        result = result.replace("  ", " ")
    return result.strip()


def apply_classification(subject: str, tags: List[str], index: int) -> str:
    if index < 1 or index > len(tags):
        return subject
    clean = remove_classification(subject, tags)
    tag = tags[index - 1]
    return f"{tag} {clean}".strip() if clean else tag


def is_valid_classification(tag: str, tags: List[str]) -> bool:
    return any(tag.upper() == t.upper() for t in tags)


def classification_level(tag: str, tags: List[str]) -> int:
    for i, t in enumerate(tags, start=1):
        if tag.upper() == t.upper():
            return i
    return 0


def test_system(_ctx: DiagContext) -> TestResult:
    lines = [
        f"OS: {platform.system()} {platform.release()}",
        f"Computer: {os.environ.get('COMPUTERNAME', '')}",
        f"User: {os.environ.get('USERNAME', '')}",
        f"Domain: {os.environ.get('USERDOMAIN', '')}",
        f"Python: {sys.version.split()[0]} ({sys.executable})",
        f"Local time: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}",
        f"TEMP: {os.environ.get('TEMP', '')}",
    ]
    return TestResult("info", "System details collected", "\n".join(lines))


def test_outlook(_ctx: DiagContext) -> TestResult:
    try:
        import win32com.client  # type: ignore
    except ImportError:
        return TestResult(
            "warn",
            "pywin32 not installed — Outlook checks skipped",
            "Install pywin32 in the Python used for AES dialogs to enable Outlook tests.",
        )
    try:
        app = win32com.client.Dispatch("Outlook.Application")
        ns = app.GetNamespace("MAPI")
        lines = [
            f"Outlook version: {app.Version}",
            f"Current user: {ns.CurrentUser.Name}",
        ]
        accounts = ns.Accounts
        count = int(accounts.Count)
        lines.append(f"Accounts: {count}")
        for i in range(1, count + 1):
            acc = accounts.Item(i)
            smtp = ""
            try:
                smtp = acc.SmtpAddress
            except Exception:
                pass
            lines.append(f"  {i}. {acc.DisplayName} ({smtp})")

        folder_names = {
            6: "Inbox",
            5: "Sent Items",
            16: "Drafts",
            4: "Outbox",
        }
        fails = []
        for ol_id, name in folder_names.items():
            try:
                fld = ns.GetDefaultFolder(ol_id)
                n = int(fld.Items.Count)
                lines.append(f"{name}: Yes ({n} items)")
            except Exception as exc:
                fails.append(name)
                lines.append(f"{name}: No — {exc}")

        lines.append(f"Explorers: {app.Explorers.Count}")
        lines.append(f"Inspectors: {app.Inspectors.Count}")
        if fails:
            return TestResult("fail", f"Folder access failed: {', '.join(fails)}", "\n".join(lines))
        return TestResult("pass", f"Outlook OK — {count} account(s)", "\n".join(lines))
    except Exception as exc:
        return TestResult("fail", f"Outlook COM failed: {exc}", str(exc))


def _pass_fail(ok: bool, summary_ok: str, summary_fail: str, detail: str = "") -> TestResult:
    if ok:
        return TestResult("pass", summary_ok, detail or summary_ok)
    return TestResult("fail", summary_fail, detail or summary_fail)


# --- AES scanner tests (same 4 checks as MSCANDiagnostics.RunAesScannerTests) ---

def test_python_exe(_ctx: DiagContext) -> TestResult:
    path = r"C:\Python313\python.exe"
    if _exists(path):
        return _pass_fail(True, "Python executable found", "", path)
    if _exists(sys.executable):
        return TestResult(
            "pass",
            "Python via current interpreter",
            f"Default path missing: {path}\nUsing: {sys.executable}",
        )
    return _pass_fail(False, "", f"Python executable not found at {path}", path)


def test_geolocate_file(_ctx: DiagContext) -> TestResult:
    vba = r"C:\GeoFooter\VBA\geolocate_headers.py"
    root = r"C:\GeoFooter\geolocate_headers.py"
    if _exists(vba) or _exists(root):
        found = vba if _exists(vba) else root
        return _pass_fail(True, "geolocate_headers.py found", "", found)
    return _pass_fail(
        False,
        "",
        "geolocate_headers.py not found",
        f"Checked:\n  {vba}\n  {root}",
    )


def test_guri_file(_ctx: DiagContext) -> TestResult:
    path = r"C:\GeoFooter\guri.py"
    return _pass_fail(_exists(path), "guri.py found", "guri.py not found", path)


def test_aes_scanning_on(ctx: DiagContext) -> TestResult:
    on = ctx.service_enabled
    watchers = ctx.watcher_count
    detail = (
        f"Service enabled: {on}\n"
        f"Watcher count (at dialog open): {watchers}\n"
        f"Queue size: {ctx.queue_size}\n"
        f"In-flight scans: {ctx.inflight_scans}\n"
        f"Startup quiet: {ctx.startup_quiet}\n\n"
        "Snapshot from Outlook when this dialog launched.\n"
        "Toggle AES ON/OFF, then reopen Diagnostics to refresh."
    )
    if not on:
        return TestResult("fail", "AES automatic scanning is OFF", detail)
    if watchers <= 0:
        return TestResult(
            "warn",
            "AES is ON but no inbox watchers are attached",
            detail + "\n\nCheck Settings → accounts with SCAN enabled, then turn AES OFF/ON.",
        )
    return TestResult("pass", f"AES automatic scanning is ON ({watchers} watcher(s))", detail)


def test_auto_scan_pipeline(ctx: DiagContext) -> TestResult:
    summary = ctx.auto_scan_summary.strip() or "(no auto-scan summary from Outlook)"
    log = ctx.recent_auto_scan_log.strip() or "(no recent auto-scan log lines)"
    detail = summary + "\n\n" + log

    lower = (summary + "\n" + log).lower()
    if not ctx.service_enabled:
        return TestResult("fail", "Auto-scan pipeline idle — service OFF", detail)
    if ctx.startup_quiet:
        return TestResult(
            "warn",
            "Startup quiet period is still active — new mail is queued but not processed yet",
            detail,
        )
    if "last itemadd: (none" in lower and "last newmailex: (none" in lower:
        return TestResult(
            "warn",
            "No ItemAdd/NewMailEx this session — catch-up heartbeat should recover misses",
            detail,
        )
    if "newmailex: skip (account scan off)" in lower or "account scan off" in lower:
        return TestResult(
            "warn",
            "Some mail was skipped because the account is set to SKIP",
            detail,
        )
    if "catchupmissedinboxmail:" in lower.replace(" ", "") or "catch-up result:" in lower:
        return TestResult("pass", "Auto-scan pipeline snapshot looks live", detail)
    if "queued" in lower or "processqueue" in lower:
        return TestResult("pass", "Auto-scan events are reaching the queue", detail)
    return TestResult("info", "Auto-scan pipeline snapshot", detail)


def test_accounts_scan_settings(ctx: DiagContext) -> TestResult:
    text = ctx.accounts_summary.strip() or "(no account summary)"
    scan = text.upper().count("[SCAN]")
    skip = text.upper().count("[SKIP]")
    detail = text
    if scan == 0 and skip > 0:
        return TestResult("fail", f"All {skip} account(s) are SKIP — nothing will auto-scan", detail)
    if scan == 0:
        return TestResult("warn", "No SCAN accounts reported", detail)
    return TestResult("pass", f"{scan} account(s) set to SCAN, {skip} SKIP", detail)


def test_recent_auto_scan_log(ctx: DiagContext) -> TestResult:
    log = ctx.recent_auto_scan_log.strip()
    if not log or "no recent" in log.lower() or "missing" in log.lower():
        return TestResult("warn", "No recent auto-scan activity in the VBA log", log or "(empty)")
    lines = [ln for ln in log.splitlines() if ln.strip() and not ln.startswith("---")]
    return TestResult(
        "info",
        f"{len(lines)} recent auto-scan log line(s)",
        log,
    )


# --- Project venv / requirements.txt ---

_VENV_ROOT = Path(r"C:\GeoFooter\.venv")
_REQUIREMENTS_PATH = Path(r"C:\GeoFooter\requirements.txt")


def _venv_python() -> Optional[Path]:
    for name in ("python.exe", "pythonw.exe"):
        p = _VENV_ROOT / "Scripts" / name
        if p.is_file():
            return p
    return None


def _parse_requirement_names(req_path: Path) -> List[str]:
    names: List[str] = []
    if not req_path.is_file():
        return names
    for raw in req_path.read_text(encoding="utf-8").splitlines():
        line = raw.split("#", 1)[0].strip()
        if not line or line.startswith("-"):
            continue
        name = line
        for sep in (">=", "==", "<=", "~=", "!=", ">", "<"):
            if sep in line:
                name = line.split(sep, 1)[0].strip()
                break
        if name:
            names.append(name)
    return names


def test_venv_up(_ctx: DiagContext) -> TestResult:
    lines = [
        f"Venv root: {_VENV_ROOT}",
        f"requirements.txt: {_REQUIREMENTS_PATH}",
    ]
    if not _VENV_ROOT.is_dir():
        return TestResult("fail", "GeoFooter .venv missing", "\n".join(lines + ["Directory not found."]))

    py = _venv_python()
    if py is None:
        return TestResult(
            "fail",
            ".venv has no python.exe",
            "\n".join(lines + [f"Expected: {_VENV_ROOT / 'Scripts' / 'python.exe'}"]),
        )
    lines.append(f"Interpreter: {py}")

    try:
        proc = subprocess.run(
            [str(py), "-c", "import sys; print(sys.version.split()[0]); print(sys.executable)"],
            capture_output=True,
            text=True,
            timeout=20,
            check=False,
        )
    except Exception as exc:
        return TestResult("fail", f"Cannot start venv Python: {exc}", "\n".join(lines))

    if proc.returncode != 0:
        err = (proc.stderr or proc.stdout or "").strip()
        return TestResult("fail", "Venv Python failed to start", "\n".join(lines + [err]))

    out_lines = [ln.strip() for ln in (proc.stdout or "").splitlines() if ln.strip()]
    version = out_lines[0] if out_lines else "?"
    exe = out_lines[1] if len(out_lines) > 1 else str(py)
    lines.append(f"Version: {version}")
    lines.append(f"Reports as: {exe}")
    in_venv = ".venv" in exe.replace("/", "\\").lower()
    if not in_venv:
        return TestResult("warn", f"Python ran but path looks odd ({version})", "\n".join(lines))
    return TestResult("pass", f"Venv up — Python {version}", "\n".join(lines))


def test_venv_requirements(_ctx: DiagContext) -> TestResult:
    py = _venv_python()
    if py is None:
        return TestResult(
            "fail",
            "Cannot check requirements — .venv missing",
            f"Expected {_VENV_ROOT / 'Scripts' / 'python.exe'}",
        )
    if not _REQUIREMENTS_PATH.is_file():
        return TestResult("fail", "requirements.txt not found", str(_REQUIREMENTS_PATH))

    names = _parse_requirement_names(_REQUIREMENTS_PATH)
    if not names:
        return TestResult("fail", "requirements.txt has no packages", str(_REQUIREMENTS_PATH))

    # Probe packages inside the venv (not the interpreter running this dialog).
    probe = r"""
import json, sys
from importlib import metadata, import_module
IMPORT_MAP = {
    "Pillow": "PIL",
    "pywin32": "win32com",
    "mysql-connector-python": "mysql.connector",
    "python-whois": "whois",
    "dnspython": "dns",
    "psycopg2-binary": "psycopg2",
}
names = json.loads(sys.argv[1])
rows = []
for name in names:
    imp = IMPORT_MAP.get(name, name.replace("-", "_"))
    try:
        ver = metadata.version(name)
    except metadata.PackageNotFoundError:
        ver = None
    try:
        import_module(imp)
        import_ok = True
        err = ""
    except Exception as e:
        import_ok = False
        err = str(e)
    rows.append({"name": name, "version": ver, "import": imp, "import_ok": import_ok, "error": err})
print(json.dumps(rows))
"""
    try:
        proc = subprocess.run(
            [str(py), "-c", probe, json.dumps(names)],
            capture_output=True,
            text=True,
            timeout=60,
            check=False,
        )
    except Exception as exc:
        return TestResult("fail", f"Requirements probe failed: {exc}", str(py))

    if proc.returncode != 0:
        return TestResult(
            "fail",
            "Requirements probe error in venv",
            (proc.stderr or proc.stdout or "").strip() or f"exit {proc.returncode}",
        )

    try:
        rows = json.loads((proc.stdout or "").strip().splitlines()[-1])
    except Exception as exc:
        return TestResult("fail", f"Bad probe output: {exc}", proc.stdout or "")

    lines: List[str] = [
        f"Venv: {py}",
        f"File: {_REQUIREMENTS_PATH}",
        f"Packages: {len(rows)}",
        "",
    ]
    missing = 0
    import_fail = 0
    for row in rows:
        name = row.get("name", "?")
        ver = row.get("version")
        imp = row.get("import", name)
        if ver is None:
            missing += 1
            lines.append(f"MISSING  {name}")
            continue
        if not row.get("import_ok"):
            import_fail += 1
            lines.append(f"IMPORT FAIL  {name}=={ver}  ({imp}: {row.get('error', '')})")
            continue
        lines.append(f"OK  {name}=={ver}  (import {imp})")

    if missing or import_fail:
        summary = f"{missing} missing, {import_fail} import fail (of {len(rows)})"
        return TestResult("fail", summary, "\n".join(lines))
    return TestResult("pass", f"All {len(rows)} requirements loaded in venv", "\n".join(lines))


# --- Classification tests (same 11 as MSCANDiagnostics.RunClassificationTests) ---

def _need_tags(ctx: DiagContext, n: int = 2) -> Optional[TestResult]:
    if len(ctx.tags) < n:
        return TestResult("fail", f"Need at least {n} tags in context", "Re-open Diagnostics from Outlook.")
    return None


def test_detect_tag1(ctx: DiagContext) -> TestResult:
    bad = _need_tags(ctx, 1)
    if bad:
        return bad
    tag1 = ctx.tags[0]
    got = detect_classification(f"{tag1} Test", ctx.tags)
    return _pass_fail(got == tag1, f"Detect {tag1}", f"Detect {tag1} failed", f"got {got!r}")


def test_detect_none(ctx: DiagContext) -> TestResult:
    got = detect_classification("Plain subject", ctx.tags)
    return _pass_fail(got == "", "Detect no tag", "Detect no tag failed", f"got {got!r}")


def test_has_classification_true(ctx: DiagContext) -> TestResult:
    bad = _need_tags(ctx, 2)
    if bad:
        return bad
    tag2 = ctx.tags[1]
    ok = bool(detect_classification(f"{tag2} Meeting", ctx.tags))
    return _pass_fail(ok, "HasClassification true", "HasClassification true failed")


def test_has_classification_false(ctx: DiagContext) -> TestResult:
    ok = not bool(detect_classification("Plain meeting", ctx.tags))
    return _pass_fail(ok, "HasClassification false", "HasClassification false failed")


def test_apply_by_index(ctx: DiagContext) -> TestResult:
    bad = _need_tags(ctx, 1)
    if bad:
        return bad
    tag1 = ctx.tags[0]
    applied = apply_classification("Test subject", ctx.tags, 1)
    ok = tag1 in applied and "Test subject" in applied
    return _pass_fail(ok, "Apply classification by index", "Apply by index failed", applied)


def test_remove_classification(ctx: DiagContext) -> TestResult:
    bad = _need_tags(ctx, 1)
    if bad:
        return bad
    tag1 = ctx.tags[0]
    removed = remove_classification(f"{tag1} Test subject", ctx.tags)
    ok = tag1 not in removed and "Test subject" in removed
    return _pass_fail(ok, "Remove classification", "Remove classification failed", removed)


def test_replace_classification(ctx: DiagContext) -> TestResult:
    bad = _need_tags(ctx, 2)
    if bad:
        return bad
    tag1, tag2 = ctx.tags[0], ctx.tags[1]
    replaced = apply_classification(f"{tag1} Old subject", ctx.tags, 2)
    ok = tag2 in replaced and tag1 not in replaced
    return _pass_fail(ok, "Replace classification", "Replace classification failed", replaced)


def test_is_valid_true(ctx: DiagContext) -> TestResult:
    bad = _need_tags(ctx, 1)
    if bad:
        return bad
    tag1 = ctx.tags[0]
    ok = is_valid_classification(tag1, ctx.tags)
    return _pass_fail(ok, "IsValidClassification true", "IsValidClassification true failed", tag1)


def test_is_valid_false(ctx: DiagContext) -> TestResult:
    ok = not is_valid_classification("[INVALID]", ctx.tags)
    return _pass_fail(ok, "IsValidClassification false", "IsValidClassification false failed", "[INVALID]")


def test_classification_level(ctx: DiagContext) -> TestResult:
    bad = _need_tags(ctx, 1)
    if bad:
        return bad
    tag1 = ctx.tags[0]
    level = classification_level(tag1, ctx.tags)
    return _pass_fail(
        level == 1,
        "GetClassificationLevel",
        f"GetClassificationLevel (expected 1, got {level})",
        str(level),
    )


def test_apply_by_tag_string(ctx: DiagContext) -> TestResult:
    bad = _need_tags(ctx, 2)
    if bad:
        return bad
    tag2 = ctx.tags[1]
    clean = remove_classification("Test subject", ctx.tags)
    by_tag = f"{tag2} {clean}".strip()
    ok = tag2 in by_tag and "Test subject" in by_tag
    return _pass_fail(ok, "Apply classification by tag string", "Apply by tag string failed", by_tag)


# --- Form / event tests (same as RunFormTests + RunEventTests) ---

def test_classification_tags(ctx: DiagContext) -> TestResult:
    n = len(ctx.tags)
    detail = "\n".join(f"  {i}. {t}" for i, t in enumerate(ctx.tags, start=1))
    detail += "\n\nUses MSCANClassificationDialog (no UserForm import)."
    if n == 0:
        return _pass_fail(False, "", "No classification tags configured", detail)
    return _pass_fail(True, f"Classification tags available ({n})", "", detail)


def test_mailitem_create(_ctx: DiagContext) -> TestResult:
    try:
        import win32com.client  # type: ignore
    except ImportError:
        return TestResult(
            "warn",
            "pywin32 missing — MailItem create skipped",
            "Install pywin32 in the Python used for AES dialogs.",
        )
    try:
        app = win32com.client.Dispatch("Outlook.Application")
        mail = app.CreateItem(0)  # olMailItem
        mail.Close(1)  # olDiscard
        return TestResult("pass", "MailItem creation", "Created a draft MailItem and discarded it.")
    except Exception as exc:
        return TestResult("fail", f"Cannot create MailItem — {exc}", str(exc))


def test_mailitem_properties(_ctx: DiagContext) -> TestResult:
    try:
        import win32com.client  # type: ignore
    except ImportError:
        return TestResult(
            "warn",
            "pywin32 missing — MailItem properties skipped",
            "Install pywin32 in the Python used for AES dialogs.",
        )
    try:
        app = win32com.client.Dispatch("Outlook.Application")
        mail = app.CreateItem(0)
        mail.Subject = "MSCAN Diagnostic Test"
        mail.To = "test@example.com"
        mail.Body = "This is a diagnostic test."
        mail.Close(1)
        return TestResult("pass", "MailItem property assignment", "Subject / To / Body assigned, then discarded.")
    except Exception as exc:
        return TestResult("fail", f"MailItem property assignment failed — {exc}", str(exc))


def build_tests() -> List[TestDef]:
    """Diagnostic suite: original 18 VBA checks + GeoFooter venv/requirements."""
    return [
        # AES scanner (4)
        TestDef("py_exe", "Python executable", "AES scanner", "C:\\Python313\\python.exe (or current interpreter)", test_python_exe),
        TestDef("geolocate", "geolocate_headers.py", "AES scanner", "VBA or root copy present", test_geolocate_file),
        TestDef("guri", "guri.py", "AES scanner", "C:\\GeoFooter\\guri.py", test_guri_file),
        TestDef("aes_on", "AES automatic scanning ON", "AES scanner", "Service enabled snapshot from Outlook", test_aes_scanning_on),
        TestDef("auto_pipeline", "Auto-scan pipeline health", "Auto-scan", "Queue, watchers, last ItemAdd/NewMailEx, catch-up", test_auto_scan_pipeline),
        TestDef("acct_scan", "Account SCAN/SKIP settings", "Auto-scan", "Which mailboxes get automatic footers", test_accounts_scan_settings),
        TestDef("auto_log", "Recent auto-scan log", "Auto-scan", "Last NewMailEx / ItemAdd / queue / catch-up lines", test_recent_auto_scan_log),
        # Project venv (2)
        TestDef("venv_up", "GeoFooter venv is up", "Python venv", "C:\\GeoFooter\\.venv runs and reports version", test_venv_up),
        TestDef("venv_reqs", "requirements.txt loaded in venv", "Python venv", "Every package in requirements.txt installs and imports", test_venv_requirements),
        # Classification (11)
        TestDef("cls_detect1", "Detect tag1", "Classification", "Detect known classification tag", test_detect_tag1),
        TestDef("cls_detect_none", "Detect no tag", "Classification", "Plain subject has no tag", test_detect_none),
        TestDef("cls_has_true", "HasClassification true", "Classification", "Tagged subject returns true", test_has_classification_true),
        TestDef("cls_has_false", "HasClassification false", "Classification", "Plain subject returns false", test_has_classification_false),
        TestDef("cls_apply_idx", "Apply classification by index", "Classification", "Apply tag via index 1", test_apply_by_index),
        TestDef("cls_remove", "Remove classification", "Classification", "Strip tag from subject", test_remove_classification),
        TestDef("cls_replace", "Replace classification", "Classification", "Swap tag1 for tag2", test_replace_classification),
        TestDef("cls_valid_true", "IsValidClassification true", "Classification", "Known tag is valid", test_is_valid_true),
        TestDef("cls_valid_false", "IsValidClassification false", "Classification", "[INVALID] is rejected", test_is_valid_false),
        TestDef("cls_level", "GetClassificationLevel", "Classification", "Tag1 maps to level 1", test_classification_level),
        TestDef("cls_apply_tag", "Apply classification by tag string", "Classification", "Apply tag2 by string", test_apply_by_tag_string),
        # Form / dialog (1)
        TestDef("form_tags", "Classification tags available", "Dialog", "Tags configured for classification dialog", test_classification_tags),
        # Event / MailItem (2)
        TestDef("mail_create", "MailItem creation", "Outlook", "CreateItem(olMailItem)", test_mailitem_create),
        TestDef("mail_props", "MailItem property assignment", "Outlook", "Subject / To / Body then discard", test_mailitem_properties),
    ]


def _geofooter_dir() -> Path:
    return Path(os.environ.get("LOCALAPPDATA", "")) / "GeoFooter"


def _logging_path() -> Path:
    return _geofooter_dir() / "aes_logging.json"


def _diag_commands_path() -> Path:
    return _geofooter_dir() / "aes_diag_commands.json"


def _write_logging_cfg(enabled: bool, levels: Dict[str, bool]) -> Path:
    path = _logging_path()
    path.parent.mkdir(parents=True, exist_ok=True)
    payload = {
        "enabled": bool(enabled),
        "levels": {
            "info": bool(levels.get("info", True)),
            "audit": bool(levels.get("audit", True)),
            "warn": bool(levels.get("warn", True)),
            "debug": bool(levels.get("debug", False)),
        },
    }
    path.write_text(json.dumps(payload, indent=2), encoding="utf-8")
    return path


def _enqueue_diag_ops(*ops: str) -> None:
    """Write ops for VBA DrainDiagCommands, then nudge Outlook via ItemLoad."""
    path = _diag_commands_path()
    path.parent.mkdir(parents=True, exist_ok=True)
    commands = [{"op": op} for op in ops if op]
    path.write_text(json.dumps({"commands": commands}, indent=2), encoding="utf-8")
    _nudge_outlook_itemload()


def _nudge_outlook_itemload() -> None:
    """Short VBS: touch an inbox item so Application_ItemLoad runs NudgeAsyncWork."""
    vbs = _geofooter_dir() / "aes_diag_nudge.vbs"
    vbs.parent.mkdir(parents=True, exist_ok=True)
    vbs.write_text(
        "\r\n".join(
            [
                "On Error Resume Next",
                "WScript.Sleep 200",
                'Set ol = GetObject(, "Outlook.Application")',
                "If ol Is Nothing Then WScript.Quit 1",
                'Set items = ol.GetNamespace("MAPI").GetDefaultFolder(6).Items',
                "If items Is Nothing Then WScript.Quit 1",
                "Set itm = items.GetFirst",
                "If Not itm Is Nothing Then dummy = itm.EntryID",
                "WScript.Quit 0",
                "",
            ]
        ),
        encoding="ascii",
        errors="ignore",
    )
    try:
        subprocess.Popen(
            ["wscript.exe", "//Nologo", "//B", str(vbs)],
            close_fds=True,
            creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0),
        )
    except Exception:
        pass


def _tail_text(path: str, max_bytes: int = 120_000) -> str:
    p = Path(path)
    if not p.is_file():
        return f"(file missing: {path})"
    try:
        size = p.stat().st_size
        with p.open("rb") as fh:
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
        return f"(read error: {exc})"


def load_context(path: Optional[str]) -> DiagContext:
    ctx = DiagContext()
    if not path:
        # Standalone defaults matching MSCANCore tags
        ctx.raw = {
            "tags": [
                "[NR/E]",
                "[SEC1: (C) NOT RATED /EXTERNAL]",
                "[SEC2: (C) COMMERCIAL-IN-CONFIDENCE /UNENCRYPTED]",
                "[SEC3: (C) COMMERCIAL-IN-CONFIDENCE-SENSITIVE/SIGNED/EDIT-ENCRYPTED]",
                "[SEC4: (C) COMMERCIAL-IN-CONFIDENCE-SENSITIVE/SIGNED/ENCRYPTED/TRACKED]",
                "[SEC5: (D) OFFICIAL-SENSITIVE/SIGNED/ENCRYPTED/TRACKED]",
                "[SEC6: (D) CLASSIFIED/SIGNED/ENCRYPTED/TRACKED]",
                "[SEC7: (D) SOA/SIGNED/ENCRYPTED/TRACKED]",
            ],
            "service_enabled": False,
            "watcher_count": 0,
            "log_path": str(Path(os.environ.get("LOCALAPPDATA", "")) / "GeoFooter" / "Logs" / "VBA_Log.txt"),
            "accounts_summary": "",
        }
        return ctx
    try:
        raw = json.loads(Path(path).read_text(encoding="utf-8"))
        if isinstance(raw, dict):
            ctx.raw = raw
    except Exception:
        pass
    if not ctx.tags:
        return load_context(None)
    return ctx


def run_dialog(ctx: DiagContext) -> int:
    try:
        from PySide6.QtCore import Qt, QTimer
        from PySide6.QtWidgets import (
            QApplication,
            QCheckBox,
            QDialog,
            QFrame,
            QGroupBox,
            QHBoxLayout,
            QLabel,
            QProgressBar,
            QPushButton,
            QScrollArea,
            QTextEdit,
            QVBoxLayout,
            QWidget,
        )
    except ImportError:
        print("aes_diagnostics_dialog failed: PySide6 is required", file=sys.stderr)
        return 1

    tests = build_tests()
    results: Dict[str, TestResult] = {}
    log_cfg = ctx.logging_cfg
    capture_active = {"on": bool(log_cfg.get("capture_active"))}
    capture_path = {"path": str(log_cfg.get("capture_path") or "")}

    app = QApplication.instance() or QApplication(sys.argv)
    dlg = QDialog()
    dlg.setWindowTitle("AES Diagnostics")
    dlg.setWindowFlags(dlg.windowFlags() | Qt.WindowStaysOnTopHint)
    dlg.resize(780, 760)
    dlg.setStyleSheet(STYLESHEET)

    root = QVBoxLayout(dlg)
    root.setContentsMargins(16, 16, 16, 16)
    root.setSpacing(10)

    title = QLabel("AES Diagnostics")
    title.setObjectName("title")
    root.addWidget(title)
    sub = QLabel(
        "Live Controls enact Outlook timers / logging without freezing the mail UI. "
        "Checks cover scanner paths, venv, classification, and the auto-scan pipeline."
    )
    sub.setObjectName("subtitle")
    sub.setWordWrap(True)
    root.addWidget(sub)

    # ---- Live Controls ----
    controls = QGroupBox("Live Controls")
    controls_layout = QVBoxLayout(controls)
    controls_layout.setSpacing(8)

    status_line = QLabel("Ready")
    status_line.setObjectName("hint")

    timer_row = QHBoxLayout()
    timer_lbl = QLabel("Timers")
    timer_lbl.setObjectName("section")
    controls_layout.addWidget(timer_lbl)
    timer_hint = QLabel(
        "Writes a command for Outlook VBA, then nudges ItemLoad. "
        "Catch-up walks inboxes (expensive) — use sparingly."
    )
    timer_hint.setObjectName("hint")
    timer_hint.setWordWrap(True)
    controls_layout.addWidget(timer_hint)

    def make_timer_btn(label: str, op: str) -> QPushButton:
        btn = QPushButton(label)
        btn.setObjectName("secondary")
        btn.clicked.connect(lambda _=False, o=op: (
            _enqueue_diag_ops(o),
            status_line.setText(f"Enacted: {o} (Outlook nudged)"),
        ))
        return btn

    for text, op in (
        ("Queue tick now", "queue_tick"),
        ("Catch-up now", "catchup"),
        ("Heartbeat now", "heartbeat"),
        ("Post-send drain", "postsend"),
    ):
        timer_row.addWidget(make_timer_btn(text, op))
    timer_row.addStretch(1)
    controls_layout.addLayout(timer_row)

    log_lbl = QLabel("Logging levels")
    log_lbl.setObjectName("section")
    controls_layout.addWidget(log_lbl)
    log_hint = QLabel(
        "Info / Warning / Audit / Debug — apply writes aes_logging.json and reloads in Outlook."
    )
    log_hint.setObjectName("hint")
    log_hint.setWordWrap(True)
    controls_layout.addWidget(log_hint)

    log_enabled = QCheckBox("Logging on")
    log_enabled.setChecked(bool(log_cfg.get("enabled", True)))
    controls_layout.addWidget(log_enabled)

    level_row = QHBoxLayout()
    level_checks: Dict[str, QCheckBox] = {}
    levels = log_cfg.get("levels") or {}
    for key, label in (
        ("info", "Info"),
        ("warn", "Warning"),
        ("audit", "Audit"),
        ("debug", "Debug"),
    ):
        cb = QCheckBox(label)
        cb.setChecked(bool(levels.get(key, key != "debug")))
        level_checks[key] = cb
        level_row.addWidget(cb)
    level_row.addStretch(1)
    controls_layout.addLayout(level_row)

    def sync_levels_enabled(checked: bool) -> None:
        for cb in level_checks.values():
            cb.setEnabled(checked)

    log_enabled.toggled.connect(sync_levels_enabled)
    sync_levels_enabled(log_enabled.isChecked())

    apply_log_btn = QPushButton("Apply logging")
    apply_log_btn.clicked.connect(lambda: (
        _write_logging_cfg(
            log_enabled.isChecked(),
            {k: cb.isChecked() for k, cb in level_checks.items()},
        ),
        _enqueue_diag_ops("reload_logging"),
        status_line.setText("Logging settings applied + Outlook reloaded"),
    ))
    controls_layout.addWidget(apply_log_btn)

    cap_lbl = QLabel("Capture")
    cap_lbl.setObjectName("section")
    controls_layout.addWidget(cap_lbl)
    cap_hint = QLabel(
        "Capture mirrors every log line to a timestamped sidecar while active "
        "(in addition to VBA_Log.txt)."
    )
    cap_hint.setObjectName("hint")
    cap_hint.setWordWrap(True)
    controls_layout.addWidget(cap_hint)

    live_log = QTextEdit()
    live_log.setReadOnly(True)
    live_log.setPlaceholderText("Log tail — Refresh or auto-refresh while capture is on.")
    live_log.setMinimumHeight(120)
    live_log.setMaximumHeight(160)

    cap_row = QHBoxLayout()
    start_cap = QPushButton("Start capture")
    stop_cap = QPushButton("Stop capture")
    stop_cap.setObjectName("secondary")
    open_cap = QPushButton("Open capture")
    open_cap.setObjectName("secondary")
    open_log = QPushButton("Open main log")
    open_log.setObjectName("secondary")
    refresh_tail = QPushButton("Refresh tail")
    refresh_tail.setObjectName("secondary")

    def start_capture() -> None:
        _enqueue_diag_ops("capture_start")
        capture_active["on"] = True
        status_line.setText("Capture start requested — check status after Outlook nudge")
        QTimer.singleShot(800, refresh_capture_status)

    def stop_capture() -> None:
        _enqueue_diag_ops("capture_stop")
        capture_active["on"] = False
        status_line.setText("Capture stop requested")
        QTimer.singleShot(800, refresh_capture_status)

    def open_path(path: str) -> None:
        if not path:
            status_line.setText("No path yet")
            return
        try:
            os.startfile(path)  # type: ignore[attr-defined]
        except Exception as exc:
            status_line.setText(f"Open failed: {exc}")

    def refresh_capture_status() -> None:
        logs = _geofooter_dir() / "Logs"
        try:
            files = sorted(
                logs.glob("VBA_Capture_*.txt"),
                key=lambda p: p.stat().st_mtime,
                reverse=True,
            )
            if files:
                capture_path["path"] = str(files[0])
        except Exception:
            pass
        state = "ON" if capture_active["on"] else "off"
        status_line.setText(
            f"Capture {state}"
            + (f" — {capture_path['path']}" if capture_path["path"] else "")
        )

    def refresh_log_tail() -> None:
        from PySide6.QtGui import QTextCursor

        path = capture_path["path"] if capture_active["on"] and capture_path["path"] else ctx.log_path
        if not path:
            path = str(_geofooter_dir() / "Logs" / "VBA_Log.txt")
        live_log.setPlainText(_tail_text(path))
        live_log.moveCursor(QTextCursor.End)

    start_cap.clicked.connect(start_capture)
    stop_cap.clicked.connect(stop_capture)
    open_cap.clicked.connect(lambda: open_path(capture_path["path"]))
    open_log.clicked.connect(
        lambda: open_path(ctx.log_path or str(_geofooter_dir() / "Logs" / "VBA_Log.txt"))
    )
    refresh_tail.clicked.connect(refresh_log_tail)

    cap_row.addWidget(start_cap)
    cap_row.addWidget(stop_cap)
    cap_row.addWidget(open_cap)
    cap_row.addWidget(open_log)
    cap_row.addWidget(refresh_tail)
    cap_row.addStretch(1)
    controls_layout.addLayout(cap_row)
    controls_layout.addWidget(status_line)
    controls_layout.addWidget(live_log)

    root.addWidget(controls)

    tail_timer = QTimer(dlg)
    tail_timer.setInterval(2500)
    tail_timer.timeout.connect(refresh_log_tail)
    tail_timer.start()
    refresh_log_tail()

    btn_row = QHBoxLayout()
    run_all_btn = QPushButton("Run all")
    clear_btn = QPushButton("Clear results")
    clear_btn.setObjectName("secondary")
    close_btn = QPushButton("Close")
    close_btn.setObjectName("secondary")
    btn_row.addWidget(run_all_btn)
    btn_row.addWidget(clear_btn)
    btn_row.addStretch(1)
    btn_row.addWidget(close_btn)
    root.addLayout(btn_row)

    progress = QProgressBar()
    progress.setRange(0, len(tests))
    progress.setValue(0)
    root.addWidget(progress)

    scroll = QScrollArea()
    scroll.setWidgetResizable(True)
    scroll.setFrameShape(QFrame.NoFrame)
    host = QWidget()
    host_layout = QVBoxLayout(host)
    host_layout.setSpacing(6)
    host_layout.setContentsMargins(0, 0, 4, 0)

    status_labels: Dict[str, QLabel] = {}
    name_buttons: Dict[str, QPushButton] = {}

    detail = QTextEdit()
    detail.setReadOnly(True)
    detail.setPlaceholderText("Select a test and click Run — details appear here.")
    detail.setMinimumHeight(180)

    def status_style(status: str) -> Tuple[str, str]:
        return {
            "pass": ("PASS", "statusPass"),
            "fail": ("FAIL", "statusFail"),
            "warn": ("WARN", "statusWarn"),
            "info": ("INFO", "statusInfo"),
        }.get(status, ("—", "statusIdle"))

    def show_detail(test_id: str) -> None:
        tdef = next((t for t in tests if t.id == test_id), None)
        if not tdef:
            return
        res = results.get(test_id)
        if not res:
            detail.setPlainText(f"{tdef.name}\n\n{tdef.description}\n\nNot run yet.")
            return
        detail.setPlainText(
            f"{tdef.name}\nStatus: {res.status.upper()}\n{res.summary}\n\n{res.detail}"
        )

    def set_status(test_id: str, res: Optional[TestResult]) -> None:
        lbl = status_labels[test_id]
        if res is None:
            lbl.setText("—")
            lbl.setObjectName("statusIdle")
        else:
            text, obj = status_style(res.status)
            lbl.setText(text)
            lbl.setObjectName(obj)
        lbl.style().unpolish(lbl)
        lbl.style().polish(lbl)

    def run_one(test_id: str) -> None:
        tdef = next(t for t in tests if t.id == test_id)
        name_buttons[test_id].setEnabled(False)
        QApplication.processEvents()
        try:
            res = tdef.runner(ctx)
        except Exception as exc:
            res = TestResult("fail", f"Exception: {exc}", str(exc))
        results[test_id] = res
        set_status(test_id, res)
        show_detail(test_id)
        name_buttons[test_id].setEnabled(True)
        done = sum(1 for t in tests if t.id in results)
        progress.setValue(done)

    def run_all() -> None:
        run_all_btn.setEnabled(False)
        for t in tests:
            run_one(t.id)
            QApplication.processEvents()
        run_all_btn.setEnabled(True)
        passed = sum(1 for r in results.values() if r.status == "pass")
        failed = sum(1 for r in results.values() if r.status == "fail")
        warn = sum(1 for r in results.values() if r.status == "warn")
        detail.setPlainText(
            f"Run all complete.\nPASS: {passed}  FAIL: {failed}  WARN: {warn}  INFO: "
            f"{sum(1 for r in results.values() if r.status == 'info')}\n\n"
            "Click a test row’s Run button to re-run or review details."
        )

    current_group = ""
    for tdef in tests:
        if tdef.group != current_group:
            current_group = tdef.group
            g = QLabel(current_group)
            g.setObjectName("section")
            host_layout.addWidget(g)

        row = QFrame()
        row.setObjectName("row")
        rl = QHBoxLayout(row)
        rl.setContentsMargins(10, 6, 10, 6)

        name = QLabel(tdef.name)
        name.setToolTip(tdef.description)
        st = QLabel("—")
        st.setObjectName("statusIdle")
        st.setMinimumWidth(48)
        status_labels[tdef.id] = st

        run_btn = QPushButton("Run")
        run_btn.setFixedWidth(64)
        tid = tdef.id
        run_btn.clicked.connect(lambda _=False, i=tid: run_one(i))
        name_buttons[tid] = run_btn

        # Clicking the name shows details
        name_click = QPushButton(tdef.name)
        name_click.setObjectName("secondary")
        name_click.setStyleSheet(
            "QPushButton#secondary { text-align: left; border: none; background: transparent; color: #1a2e35; }"
            "QPushButton#secondary:hover { color: #0f6b7c; background: transparent; }"
        )
        name_click.clicked.connect(lambda _=False, i=tid: show_detail(i))

        rl.addWidget(name_click, 1)
        rl.addWidget(st)
        rl.addWidget(run_btn)
        host_layout.addWidget(row)

    host_layout.addStretch(1)
    scroll.setWidget(host)
    root.addWidget(scroll, 1)

    det_lbl = QLabel("Result detail")
    det_lbl.setObjectName("section")
    root.addWidget(det_lbl)
    root.addWidget(detail)

    run_all_btn.clicked.connect(run_all)
    clear_btn.clicked.connect(lambda: (
        results.clear(),
        [set_status(t.id, None) for t in tests],
        progress.setValue(0),
        detail.clear(),
    ))
    close_btn.clicked.connect(dlg.accept)

    dlg.exec()
    return 0


def main() -> int:
    parser = argparse.ArgumentParser(description="AES interactive diagnostics")
    parser.add_argument("--context", default="", help="JSON context from Outlook VBA")
    args = parser.parse_args()
    try:
        ctx = load_context(args.context or None)
        return run_dialog(ctx)
    except Exception as exc:
        print(f"aes_diagnostics_dialog failed: {exc}", file=sys.stderr)
        return 1


if __name__ == "__main__":
    sys.exit(main())
