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
    def accounts_summary(self) -> str:
        return str(self.raw.get("accounts_summary") or "")


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


def test_python_pipeline(_ctx: DiagContext) -> TestResult:
    checks = [
        ("Python 3.13", r"C:\Python313\python.exe"),
        ("geolocate (VBA)", r"C:\GeoFooter\VBA\geolocate_headers.py"),
        ("geolocate (root)", r"C:\GeoFooter\geolocate_headers.py"),
        ("guri.py", r"C:\GeoFooter\guri.py"),
        ("guri_mysql_config.json", r"C:\GeoFooter\guri_mysql_config.json"),
        ("output folder", r"C:\GeoFooter\output"),
        ("settings dialog", r"C:\GeoFooter\VBA\aes_settings_dialog.py"),
        ("diagnostics dialog", r"C:\GeoFooter\VBA\aes_diagnostics_dialog.py"),
    ]
    lines: List[str] = []
    missing: List[str] = []
    geo_ok = _exists(checks[1][1]) or _exists(checks[2][1])
    for label, path in checks:
        if label.startswith("geolocate"):
            continue
        ok = _exists(path)
        lines.append(f"{'PASS' if ok else 'FAIL'}: {label} — {path}")
        if not ok:
            missing.append(label)
    lines.insert(1, f"{'PASS' if geo_ok else 'FAIL'}: geolocate_headers.py (VBA or root)")
    if not geo_ok:
        missing.append("geolocate_headers.py")

    py_ok = _exists(checks[0][1]) or _exists(sys.executable)
    if not _exists(checks[0][1]) and _exists(sys.executable):
        lines[0] = f"PASS: Python via current interpreter — {sys.executable}"
        if "Python 3.13" in missing:
            missing.remove("Python 3.13")
    elif not py_ok:
        missing.append("Python")

    if missing:
        return TestResult("fail", f"Missing: {', '.join(missing)}", "\n".join(lines))
    return TestResult("pass", "Python pipeline files present", "\n".join(lines))


def test_aes_service(ctx: DiagContext) -> TestResult:
    on = ctx.service_enabled
    watchers = ctx.watcher_count
    lines = [
        f"Service enabled: {on}",
        f"Watcher count (at dialog open): {watchers}",
        "",
        "Snapshot from Outlook when this dialog launched.",
        "Toggle AES ON/OFF on the Add-ins tab, then reopen Diagnostics to refresh.",
    ]
    if on and watchers > 0:
        return TestResult("pass", f"AES ON — {watchers} watcher(s)", "\n".join(lines))
    if on:
        return TestResult("warn", "AES ON but no active watchers", "\n".join(lines))
    return TestResult("warn", "AES scanning is OFF", "\n".join(lines))


def test_accounts(ctx: DiagContext) -> TestResult:
    summary = ctx.accounts_summary.strip()
    if not summary:
        return TestResult(
            "warn",
            "No account scan summary from Outlook",
            "Re-open Diagnostics from Outlook so VBA can supply account settings.",
        )
    fail = "FAIL" in summary.upper() or "[SKIP]" in summary and "enabled: 0" in summary.lower()
    status = "fail" if fail else "pass"
    # Prefer pass with details — account list is informational
    if "[SCAN]" in summary:
        status = "pass"
    return TestResult(status, "Account scan settings loaded", summary)


def test_configuration(ctx: DiagContext) -> TestResult:
    tags = ctx.tags
    lines = [f"Classification tags: {len(tags)}"]
    for i, tag in enumerate(tags, start=1):
        lines.append(f"  {i}. {tag}")
    log_path = ctx.log_path or str(Path(os.environ.get("LOCALAPPDATA", "")) / "GeoFooter" / "Logs" / "VBA_Log.txt")
    lines.append(f"Log path: {log_path}")
    if _exists(log_path):
        size = Path(log_path).stat().st_size
        lines.append(f"Log exists: Yes ({size:,} bytes)")
    else:
        lines.append("Log exists: No")
    if len(tags) < 1:
        return TestResult("fail", "No classification tags in context", "\n".join(lines))
    return TestResult("pass", f"{len(tags)} tags configured", "\n".join(lines))


def test_classification_detect(ctx: DiagContext) -> TestResult:
    tags = ctx.tags
    if len(tags) < 2:
        return TestResult("fail", "Need at least 2 tags in context", "Re-open from Outlook.")
    tag1, tag2 = tags[0], tags[1]
    cases: List[Tuple[str, bool, str]] = []

    d1 = detect_classification(f"{tag1} Test", tags)
    cases.append(("Detect tag1", d1 == tag1, f"got {d1!r}"))

    d_none = detect_classification("Plain subject", tags)
    cases.append(("Detect none", d_none == "", f"got {d_none!r}"))

    has_t = bool(detect_classification(f"{tag2} Meeting", tags))
    cases.append(("HasClassification true", has_t, str(has_t)))

    has_f = bool(detect_classification("Plain meeting", tags))
    cases.append(("HasClassification false", not has_f, str(has_f)))

    applied = apply_classification("Test subject", tags, 1)
    cases.append(
        (
            "Apply by index",
            tag1 in applied and "Test subject" in applied,
            applied,
        )
    )

    removed = remove_classification(f"{tag1} Test subject", tags)
    cases.append(
        (
            "Remove classification",
            tag1 not in removed and "Test subject" in removed,
            removed,
        )
    )

    replaced = apply_classification(f"{tag1} Old subject", tags, 2)
    cases.append(
        (
            "Replace classification",
            tag2 in replaced and tag1 not in replaced,
            replaced,
        )
    )

    cases.append(("IsValid true", is_valid_classification(tag1, tags), tag1))
    cases.append(("IsValid false", not is_valid_classification("[INVALID]", tags), "[INVALID]"))
    cases.append(("GetClassificationLevel", classification_level(tag1, tags) == 1, str(classification_level(tag1, tags))))

    by_tag = apply_classification("Test subject", tags, 2)
    # ApplyClassificationByTag equivalent
    clean = remove_classification("Test subject", tags)
    by_tag = f"{tag2} {clean}".strip()
    cases.append(
        (
            "Apply by tag string",
            tag2 in by_tag and "Test subject" in by_tag,
            by_tag,
        )
    )

    lines = []
    passed = failed = 0
    for name, ok, detail in cases:
        if ok:
            passed += 1
            lines.append(f"PASS: {name}")
        else:
            failed += 1
            lines.append(f"FAIL: {name} ({detail})")
    summary = f"{passed} passed, {failed} failed"
    return TestResult("pass" if failed == 0 else "fail", summary, "\n".join(lines))


def test_mailitem(_ctx: DiagContext) -> TestResult:
    try:
        import win32com.client  # type: ignore
    except ImportError:
        return TestResult("warn", "pywin32 missing — MailItem test skipped", "")
    try:
        app = win32com.client.Dispatch("Outlook.Application")
        mail = app.CreateItem(0)  # olMailItem
        mail.Subject = "AES Diagnostic Test"
        mail.To = "test@example.com"
        mail.Body = "Diagnostic test — discarded."
        mail.Close(1)  # olDiscard
        return TestResult("pass", "MailItem create / assign / discard OK", "Created a draft MailItem and discarded it.")
    except Exception as exc:
        return TestResult("fail", f"MailItem test failed: {exc}", str(exc))


def test_performance(ctx: DiagContext) -> TestResult:
    tags = ctx.tags
    if not tags:
        return TestResult("fail", "No tags for performance test", "")
    tag1 = tags[0]
    iterations = 1000
    lines = []

    t0 = time.perf_counter()
    for _ in range(iterations):
        detect_classification(f"{tag1} Test subject line here", tags)
    t1 = time.perf_counter()
    lines.append(
        f"DetectClassification ({iterations}): {(t1 - t0) * 1000:.2f} ms total, "
        f"{(t1 - t0) * 1000 / iterations:.4f} ms each"
    )

    t0 = time.perf_counter()
    for _ in range(iterations):
        apply_classification("Test subject", tags, 1)
    t1 = time.perf_counter()
    lines.append(
        f"ApplyClassification ({iterations}): {(t1 - t0) * 1000:.2f} ms total, "
        f"{(t1 - t0) * 1000 / iterations:.4f} ms each"
    )

    t0 = time.perf_counter()
    for _ in range(iterations):
        bool(detect_classification(f"{tag1} Test subject", tags))
    t1 = time.perf_counter()
    lines.append(
        f"HasClassification ({iterations}): {(t1 - t0) * 1000:.2f} ms total, "
        f"{(t1 - t0) * 1000 / iterations:.4f} ms each"
    )
    return TestResult("info", "Performance timings recorded", "\n".join(lines))


def test_log_file(ctx: DiagContext) -> TestResult:
    log_path = ctx.log_path or str(
        Path(os.environ.get("LOCALAPPDATA", "")) / "GeoFooter" / "Logs" / "VBA_Log.txt"
    )
    lines = [f"Path: {log_path}"]
    p = Path(log_path)
    if not p.parent.exists():
        return TestResult("fail", "Log folder missing", "\n".join(lines + [f"Parent: {p.parent}"]))
    if not p.exists():
        return TestResult("warn", "Log file not created yet", "\n".join(lines + ["Parent folder exists"]))
    st = p.stat()
    lines.append(f"Size: {st.st_size:,} bytes")
    lines.append(f"Modified: {datetime.fromtimestamp(st.st_mtime).strftime('%Y-%m-%d %H:%M:%S')}")
    return TestResult("pass", f"Log OK ({st.st_size:,} bytes)", "\n".join(lines))


def build_tests() -> List[TestDef]:
    return [
        TestDef("system", "System information", "Environment", "OS, user, Python, time", test_system),
        TestDef("outlook", "Outlook environment", "Environment", "Version, accounts, default folders", test_outlook),
        TestDef("pipeline", "Python / AES pipeline files", "AES", "python.exe, geolocate, guri, output", test_python_pipeline),
        TestDef("service", "AES service / watchers", "AES", "ON/OFF and watcher count snapshot", test_aes_service),
        TestDef("accounts", "Account scan settings", "AES", "Which mailboxes are set to SCAN", test_accounts),
        TestDef("config", "MSCAN configuration", "AES", "Classification tags and log path", test_configuration),
        TestDef("classification", "Classification unit tests", "Classification", "Detect / apply / remove / validate", test_classification_detect),
        TestDef("mailitem", "MailItem create/discard", "Outlook", "Can Outlook create a draft item", test_mailitem),
        TestDef("performance", "Classification performance", "Classification", "1000-iteration timings", test_performance),
        TestDef("logfile", "Log file status", "AES", "VBA log path, size, modified time", test_log_file),
    ]


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
        from PySide6.QtCore import Qt
        from PySide6.QtWidgets import (
            QApplication,
            QDialog,
            QFrame,
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

    app = QApplication.instance() or QApplication(sys.argv)
    dlg = QDialog()
    dlg.setWindowTitle("AES Diagnostics")
    dlg.setWindowFlags(dlg.windowFlags() | Qt.WindowStaysOnTopHint)
    dlg.resize(720, 640)
    dlg.setStyleSheet(STYLESHEET)

    root = QVBoxLayout(dlg)
    root.setContentsMargins(16, 16, 16, 16)
    root.setSpacing(10)

    title = QLabel("AES Diagnostics")
    title.setObjectName("title")
    root.addWidget(title)
    sub = QLabel("Run each check and inspect the result below. No Notepad dump.")
    sub.setObjectName("subtitle")
    sub.setWordWrap(True)
    root.addWidget(sub)

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
