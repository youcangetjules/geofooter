#!/usr/bin/env python3
"""AES settings dialog (tabbed: Mail, ASN Risk, External APIs, Logging).

CLI:
  pythonw aes\\settings_dialog.py --accounts accounts.json --out result.json
      [--route aes_risk_route.json] [--logging aes_logging.json]
"""

from __future__ import annotations

import argparse
import json
import os
import re
import sys
import threading
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple

# Launched by path from VBA: put the install root (not aes\) on sys.path.
_PKG_DIR = Path(__file__).resolve().parent
sys.path[:] = [p for p in sys.path if Path(p or ".").resolve() != _PKG_DIR]
if str(_PKG_DIR.parent) not in sys.path:
    sys.path.insert(0, str(_PKG_DIR.parent))


DEFAULT_ROUTE_RISK: Dict[str, Any] = {
    "provider": "abuseipdb",
    "abuse_threshold": 25,
    "extra_high_risk_asns": [],
    "enabled": True,
}

DEFAULT_LOGGING: Dict[str, Any] = {
    "enabled": True,
    "max_mib": 0,
    "levels": {
        "info": True,
        "audit": True,
        "warn": True,
        "debug": False,
    },
}

LOG_LEVEL_HELP = {
    "info": "Operational status — watchers, scans started/finished, queue ticks",
    "audit": "Security actions — settings saved, classification, ReadNotify stamps",
    "warn": "Recoverable issues — retries, skipped items, fallbacks",
    "debug": "Verbose diagnostics — SendDiag, hop/API detail (noisy)",
}

PROVIDERS = [
    ("abuseipdb", "AbuseIPDB (default)"),
]

STYLESHEET = """
QDialog { background: #f5f8fa; }
QLabel#title { font-size: 16px; font-weight: 700; color: #0f6b7c; }
QLabel#brandLogo { background: transparent; }
QLabel#subtitle { color: #4a6570; font-size: 12px; }
QLabel#section { font-size: 13px; font-weight: 700; color: #0f6b7c; margin-top: 4px; }
QLabel#hint { color: #5a7280; font-size: 11px; }
QFrame#column {
    background: #ffffff;
    border: 1px solid #d5dee3;
    border-radius: 8px;
}
QFrame#accountCard {
    background: #f8fbfc;
    border: 1px solid #d5dee3;
    border-radius: 6px;
}
QTabWidget::pane {
    background: #ffffff;
    border: 1px solid #d5dee3;
    border-radius: 8px;
    top: -1px;
    padding: 10px;
}
QTabBar::tab {
    background: #eef3f5;
    border: 1px solid #c5d4da;
    border-bottom: none;
    border-top-left-radius: 6px;
    border-top-right-radius: 6px;
    padding: 8px 14px;
    margin-right: 3px;
    color: #234650;
    font-weight: 600;
}
QTabBar::tab:selected {
    background: #ffffff;
    color: #0f6b7c;
}
QTabBar::tab:hover:!selected {
    background: #e2eaee;
}
QCheckBox#accountOpt {
    border: none;
    background: transparent;
    padding: 2px 4px;
    font-size: 11px;
}
QCheckBox {
    spacing: 10px;
    padding: 6px 8px;
    background: #ffffff;
    border: 1px solid #c5d4da;
    border-radius: 6px;
    font-size: 12px;
    color: #1a2e35;
}
QCheckBox:hover { border-color: #0f6b7c; }
QCheckBox::indicator {
    width: 16px;
    height: 16px;
    border: 1px solid #ffffff;
    border-radius: 3px;
    background: #d9e3e7;
}
QCheckBox::indicator:checked {
    background: #1e6bff;
    border: 1px solid #ffffff;
}
QCheckBox#plainCheck {
    border: none;
    background: transparent;
    padding: 2px 0;
}
QCheckBox#levelCheck {
    border: 1px solid #c5d4da;
    background: #f8fbfc;
    padding: 8px 10px;
}
QComboBox, QSpinBox, QPlainTextEdit, QLineEdit {
    background: #ffffff;
    border: 1px solid #c5d4da;
    border-radius: 4px;
    padding: 4px 6px;
    font-size: 12px;
    color: #1a2e35;
}
QPlainTextEdit { min-height: 72px; max-height: 88px; }
QPlainTextEdit#consoleView {
    background: #1a2e35;
    color: #d7e6eb;
    border: 1px solid #0c5663;
    border-radius: 6px;
    font-family: Consolas, "Cascadia Mono", "Courier New", monospace;
    font-size: 11px;
    padding: 6px;
    min-height: 280px;
    max-height: 16777215;
}
QLineEdit#apiKeyEdit { min-height: 28px; }
QLineEdit#consolePathEdit { min-height: 28px; }
QLabel#apiOk { color: #1b7a3d; font-size: 12px; font-weight: 600; }
QLabel#apiBad { color: #b42318; font-size: 12px; font-weight: 600; }
QLabel#apiWait { color: #5a7280; font-size: 12px; }
QPushButton { min-width: 96px; padding: 8px 16px; border-radius: 4px; font-weight: 600; }
QPushButton#secondaryBtn { background: #eef3f5; border: 1px solid #9bb4be; color: #234650; }
QPushButton#secondaryBtn:hover { background: #e2eaee; }
QPushButton#cancelBtn { background: #eef3f5; border: 1px solid #9bb4be; color: #234650; }
QPushButton#cancelBtn:hover { background: #e2eaee; }
QPushButton#saveBtn { background: #0f6b7c; border: 1px solid #0c5663; color: #ffffff; }
QPushButton#saveBtn:hover { background: #128399; }
QComboBox { padding-right: 24px; }
QComboBox::drop-down {
    subcontrol-origin: padding;
    subcontrol-position: center right;
    width: 22px;
    border: none;
    background: transparent;
}
QComboBox::down-arrow { image: url(__CHEVRON__); width: 12px; height: 12px; }
QComboBox::down-arrow:on { top: 1px; }
""".replace(
    "__CHEVRON__",
    (Path(__file__).resolve().parents[1] / "assets" / "icons" / "chevron_down.svg").as_posix(),
)


def _local_geofooter() -> Path:
    local = os.environ.get("LOCALAPPDATA", "")
    if local:
        return Path(local) / "GeoFooter"
    return Path(r"C:\GeoFooter")


def _vba_log_path() -> Path:
    candidates = [
        _local_geofooter() / "Logs" / "VBA_Log.txt",
        Path(r"C:\GeoFooter\Logs\VBA_Log.txt"),
        Path(os.environ.get("TEMP", "") or ".") / "GeoFooter" / "Logs" / "VBA_Log.txt",
    ]
    for path in candidates:
        if path.is_file():
            return path
    return candidates[0]


def _read_log_tail(path: Path, max_lines: int = 500) -> str:
    """Return the last max_lines of a log file (UTF-8 / system fallback)."""
    if not path.is_file():
        return f"(Log file not found yet)\n{path}"
    try:
        with path.open("rb") as fh:
            fh.seek(0, os.SEEK_END)
            size = fh.tell()
            block = 64 * 1024
            data = b""
            while size > 0 and data.count(b"\n") <= max_lines:
                step = min(block, size)
                size -= step
                fh.seek(size)
                data = fh.read(step) + data
                if size == 0:
                    break
        text = data.decode("utf-8", errors="replace")
        lines = text.splitlines()
        if len(lines) > max_lines:
            lines = lines[-max_lines:]
        return "\n".join(lines)
    except Exception as exc:  # noqa: BLE001
        return f"(Could not read log: {exc})\n{path}"


def _resolve_logo_path() -> Path | None:
    """Prefer the Smart Ass / Aliniant brand mark."""
    candidates: list[Path] = []
    try:
        from geofooter.paths import brand_path

        for name in ("smart-ass.svg", "smart-ass-email.svg", "AES.png"):
            candidates.append(brand_path(name))
    except Exception:
        pass
    parent = Path(__file__).resolve().parent
    candidates.extend(
        [
            parent.parent / "assets" / "brand" / "smart-ass.svg",
            parent.parent / "assets" / "brand" / "AES.png",
            parent / "smart-ass.svg",
            parent / "AES.png",
        ]
    )
    for path in candidates:
        if path.is_file():
            return path
    return None


def _load_logo_pixmap(height: int = 120):
    """Load brand logo as a QPixmap (SVG uses its native tight viewBox)."""
    try:
        from PySide6.QtCore import QRectF, Qt
        from PySide6.QtGui import QPainter, QPixmap
    except ImportError:
        return None

    path = _resolve_logo_path()
    if path is None:
        return None

    if path.suffix.lower() == ".svg":
        try:
            from PySide6.QtSvg import QSvgRenderer
        except ImportError:
            return None
        renderer = QSvgRenderer(str(path))
        if not renderer.isValid():
            return None
        vb = renderer.viewBoxF()
        if vb.width() <= 0 or vb.height() <= 0:
            size = renderer.defaultSize()
            aspect = (size.width() / size.height()) if size.height() else 5.0
        else:
            aspect = vb.width() / vb.height()
        width = max(1, int(round(height * aspect)))
        pix = QPixmap(width, height)
        pix.fill(Qt.GlobalColor.transparent)
        painter = QPainter(pix)
        painter.setRenderHint(QPainter.RenderHint.Antialiasing, True)
        renderer.render(painter, QRectF(0, 0, width, height))
        painter.end()
        return pix

    pix = QPixmap(str(path))
    if pix.isNull():
        return None
    return pix.scaledToHeight(height, Qt.TransformationMode.SmoothTransformation)


def _make_logo_label(parent=None, height: int = 120):
    from PySide6.QtCore import Qt
    from PySide6.QtWidgets import QLabel

    brand = QLabel(parent)
    brand.setObjectName("brandLogo")
    brand.setAlignment(Qt.AlignmentFlag.AlignCenter)
    pix = _load_logo_pixmap(height)
    if pix is not None and not pix.isNull():
        brand.setPixmap(pix)
        brand.setFixedSize(pix.width(), pix.height())
    else:
        brand.setText("Aliniant Security Suite")
        brand.setObjectName("title")
    return brand


def probe_abuseipdb_endpoint(
    api_key: str,
    check_url: str = "",
) -> Tuple[bool, str]:
    """Return (ok, message). ok means the AbuseIPDB API endpoint is reachable
    and the key is accepted (or rate-limited, which still proves reachability).
    """
    key = (api_key or "").strip()
    if not key:
        return False, "No API key"

    try:
        from aes.secret_store import aes_http_get, abuseipdb_check_url as _default_check_url
    except ImportError:
        return False, "aes_secrets / requests unavailable"

    url = (check_url or "").strip() or _default_check_url()

    try:
        resp, insecure = aes_http_get(
            url,
            headers={"Key": key, "Accept": "application/json"},
            params={"ipAddress": "8.8.8.8", "maxAgeInDays": 1},
            timeout=8,
        )
        tls_note = " (TLS via proxy fallback)" if insecure else ""
        if resp.status_code == 200:
            return True, f"AbuseIPDB endpoint reachable{tls_note}"
        if resp.status_code == 429:
            return True, f"AbuseIPDB endpoint reachable (rate limited){tls_note}"
        if resp.status_code in (401, 403):
            return False, f"AbuseIPDB key rejected{tls_note}"
        return False, f"AbuseIPDB HTTP {resp.status_code}{tls_note}"
    except Exception as exc:  # noqa: BLE001
        name = type(exc).__name__
        try:
            from requests import exceptions as req_exc
        except ImportError:
            req_exc = None  # type: ignore
        if req_exc is not None:
            if isinstance(exc, req_exc.SSLError):
                return False, "AbuseIPDB TLS/certificate error"
            if isinstance(exc, req_exc.ConnectionError):
                return False, "AbuseIPDB unreachable (connection failed)"
            if isinstance(exc, req_exc.Timeout):
                return False, "AbuseIPDB unreachable (timeout)"
        return False, f"AbuseIPDB unreachable ({name})"


def run_abuseipdb_dialog(parent=None) -> bool:
    """Modal AbuseIPDB API dialog (URL + key + reachability). Returns True if saved."""
    try:
        from PySide6.QtCore import QObject, Qt, QTimer, Signal
        from PySide6.QtWidgets import (
            QCheckBox,
            QDialog,
            QFormLayout,
            QHBoxLayout,
            QLabel,
            QLineEdit,
            QPushButton,
            QVBoxLayout,
        )
    except ImportError:
        print("PySide6 is required for AbuseIPDB dialog", file=sys.stderr)
        return False

    try:
        from aes.secret_store import (
            DEFAULT_ABUSEIPDB_BASE_URL,
            abuseipdb_check_url,
            abuseipdb_status,
            get_abuseipdb_api_key,
            get_abuseipdb_base_url,
            set_abuseipdb_api_key,
            set_abuseipdb_base_url,
        )
    except ImportError as exc:
        print(f"aes_secrets unavailable: {exc}", file=sys.stderr)
        return False

    dlg = QDialog(parent)
    dlg.setWindowTitle("AbuseIPDB API")
    dlg.setModal(True)
    dlg.setWindowFlag(Qt.WindowType.WindowStaysOnTopHint, True)
    dlg.setStyleSheet(STYLESHEET)
    dlg.setMinimumWidth(520)
    dlg.resize(560, 320)

    root = QVBoxLayout(dlg)
    root.setSpacing(10)
    root.setContentsMargins(16, 14, 16, 14)

    root.addWidget(_make_logo_label(dlg, height=128), 0, Qt.AlignmentFlag.AlignHCenter)
    title = QLabel("AbuseIPDB API")
    title.setObjectName("title")
    title.setAlignment(Qt.AlignmentFlag.AlignHCenter)
    root.addWidget(title)
    sub = QLabel(
        "Configure the AbuseIPDB endpoint URL and API key. "
        "The key is stored with Windows DPAPI for your user only."
    )
    sub.setObjectName("subtitle")
    sub.setWordWrap(True)
    root.addWidget(sub)

    form = QFormLayout()
    form.setSpacing(8)

    url_edit = QLineEdit()
    url_edit.setObjectName("apiKeyEdit")
    url_edit.setText(get_abuseipdb_base_url())
    url_edit.setPlaceholderText(DEFAULT_ABUSEIPDB_BASE_URL)
    url_edit.setClearButtonEnabled(True)
    url_edit.setToolTip("Base API URL, e.g. https://api.abuseipdb.com/api/v2")
    form.addRow("API URL", url_edit)

    status = abuseipdb_status()
    key_edit = QLineEdit()
    key_edit.setObjectName("apiKeyEdit")
    key_edit.setEchoMode(QLineEdit.EchoMode.Password)
    key_edit.setPlaceholderText(
        f"Current key {status.get('hint')} — paste a new key to replace"
        if status.get("configured")
        else "Paste your AbuseIPDB API key"
    )
    key_edit.setClearButtonEnabled(True)
    form.addRow("API key", key_edit)
    root.addLayout(form)

    clear_key = QCheckBox("Clear stored API key on Save")
    clear_key.setObjectName("plainCheck")
    clear_key.setEnabled(bool(status.get("configured")) and status.get("source") != "env")
    root.addWidget(clear_key)

    reach_row = QHBoxLayout()
    reach_lbl = QLabel("Checking AbuseIPDB…")
    reach_lbl.setObjectName("apiWait")
    reach_lbl.setWordWrap(True)
    test_btn = QPushButton("Test")
    test_btn.setObjectName("secondaryBtn")
    reach_row.addWidget(reach_lbl, stretch=1)
    reach_row.addWidget(test_btn)
    root.addLayout(reach_row)

    hint = QLabel(
        f"Default: {DEFAULT_ABUSEIPDB_BASE_URL}\n"
        "Check endpoint used for tests: …/check"
    )
    hint.setObjectName("hint")
    hint.setWordWrap(True)
    root.addWidget(hint)

    buttons = QHBoxLayout()
    buttons.addStretch(1)
    save_btn = QPushButton("Save")
    save_btn.setObjectName("saveBtn")
    save_btn.setDefault(True)
    close_btn = QPushButton("Close")
    close_btn.setObjectName("cancelBtn")
    buttons.addWidget(save_btn)
    buttons.addWidget(close_btn)
    root.addLayout(buttons)

    saved = {"ok": False}
    _probe_gen = {"n": 0}

    def _refresh_key_ui() -> None:
        st = abuseipdb_status()
        key_edit.clear()
        key_edit.setPlaceholderText(
            f"Current key {st.get('hint')} — paste a new key to replace"
            if st.get("configured")
            else "Paste your AbuseIPDB API key"
        )
        clear_key.setChecked(False)
        clear_key.setEnabled(bool(st.get("configured")) and st.get("source") != "env")

    def _effective_key() -> str:
        typed = key_edit.text().strip()
        if typed:
            return typed
        return (get_abuseipdb_api_key() or "").strip()

    def _effective_check_url() -> str:
        base = url_edit.text().strip() or get_abuseipdb_base_url()
        base = base.rstrip("/")
        if base.lower().endswith("/check"):
            return base
        return base + "/check"

    class _ProbeBridge(QObject):
        finished = Signal(bool, str, int)

    probe_bridge = _ProbeBridge(dlg)

    def _apply_probe(ok: bool, message: str, gen: int) -> None:
        if gen != _probe_gen["n"]:
            return
        reach_lbl.setObjectName("apiOk" if ok else "apiBad")
        reach_lbl.setText(("✓  " if ok else "✗  ") + message)
        reach_lbl.style().unpolish(reach_lbl)
        reach_lbl.style().polish(reach_lbl)
        test_btn.setEnabled(True)

    probe_bridge.finished.connect(_apply_probe)

    def start_probe() -> None:
        _probe_gen["n"] += 1
        gen = _probe_gen["n"]
        reach_lbl.setObjectName("apiWait")
        reach_lbl.setText("Checking AbuseIPDB…")
        reach_lbl.style().unpolish(reach_lbl)
        reach_lbl.style().polish(reach_lbl)
        test_btn.setEnabled(False)
        key = _effective_key()
        check_url = _effective_check_url()

        def work() -> None:
            try:
                ok, message = probe_abuseipdb_endpoint(key, check_url)
            except Exception as exc:  # noqa: BLE001
                ok, message = False, f"Probe failed ({type(exc).__name__})"
            probe_bridge.finished.emit(ok, message, gen)

        threading.Thread(target=work, daemon=True).start()

    def on_save() -> None:
        try:
            set_abuseipdb_base_url(url_edit.text())
            if clear_key.isChecked():
                set_abuseipdb_api_key("")
            else:
                typed = key_edit.text().strip()
                if typed:
                    set_abuseipdb_api_key(typed)
            saved["ok"] = True
            _refresh_key_ui()
            url_edit.setText(get_abuseipdb_base_url())
            reach_lbl.setObjectName("apiOk")
            reach_lbl.setText("✓  Settings saved")
            reach_lbl.style().unpolish(reach_lbl)
            reach_lbl.style().polish(reach_lbl)
            start_probe()
        except Exception as exc:  # noqa: BLE001
            reach_lbl.setObjectName("apiBad")
            reach_lbl.setText(f"✗  Save failed: {exc}")
            reach_lbl.style().unpolish(reach_lbl)
            reach_lbl.style().polish(reach_lbl)

    def on_close() -> None:
        if saved["ok"]:
            dlg.accept()
        else:
            dlg.reject()

    test_btn.clicked.connect(start_probe)
    url_edit.editingFinished.connect(start_probe)
    key_edit.editingFinished.connect(start_probe)
    save_btn.clicked.connect(on_save)
    close_btn.clicked.connect(on_close)
    QTimer.singleShot(150, start_probe)

    dlg.exec()
    return bool(saved["ok"])


def _default_route_path() -> Path:
    return _local_geofooter() / "aes_risk_route.json"


def _default_logging_path() -> Path:
    return _local_geofooter() / "aes_logging.json"


def _default_beacon_path() -> Path:
    return _local_geofooter() / "aes_beacon_blocking.json"


DEFAULT_BEACON_BLOCKING: Dict[str, Any] = {
    "enabled": False,
    "mode": "defang",          # "strip" removes the tag; "defang" swaps src for an inert placeholder
    "block_all": True,         # False = only senders on the per-sender block list
    "whitelist_domains": [],
}


def _normalize_beacon_blocking(raw: Any) -> Dict[str, Any]:
    cfg = dict(DEFAULT_BEACON_BLOCKING)
    cfg["whitelist_domains"] = list(DEFAULT_BEACON_BLOCKING["whitelist_domains"])
    if not isinstance(raw, dict):
        return cfg
    if "enabled" in raw:
        cfg["enabled"] = bool(raw.get("enabled"))
    mode = str(raw.get("mode") or cfg["mode"]).strip().lower()
    cfg["mode"] = mode if mode in {"strip", "defang"} else "defang"
    if "block_all" in raw:
        cfg["block_all"] = bool(raw.get("block_all"))
    wl = raw.get("whitelist_domains") or []
    if isinstance(wl, str):
        wl = re.split(r"[\s,;]+", wl)
    seen: List[str] = []
    for item in wl:
        dom = str(item or "").strip().lower().lstrip("@")
        if dom and dom not in seen:
            seen.append(dom)
    cfg["whitelist_domains"] = seen
    return cfg


def _default_sender_status_path() -> Path:
    return _local_geofooter() / "aes_sender_status.json"


DEFAULT_SENDER_STATUS: Dict[str, Any] = {
    # Emails from a sender before they count as Known rather than Unknown.
    "known_threshold": 5,
}


def _normalize_sender_status(raw: Any) -> Dict[str, Any]:
    cfg = dict(DEFAULT_SENDER_STATUS)
    if not isinstance(raw, dict):
        return cfg
    try:
        cfg["known_threshold"] = max(
            0, min(1000, int(raw.get("known_threshold", cfg["known_threshold"])))
        )
    except (TypeError, ValueError):
        pass
    return cfg


def _normalize_asn(value: Any) -> str:
    text = str(value or "").strip().upper()
    if not text:
        return ""
    match = re.search(r"AS\s*(\d+)", text)
    if match:
        return f"AS{match.group(1)}"
    if text.isdigit():
        return f"AS{text}"
    return ""


def _normalize_route_risk(raw: Any) -> Dict[str, Any]:
    cfg = dict(DEFAULT_ROUTE_RISK)
    cfg["extra_high_risk_asns"] = list(DEFAULT_ROUTE_RISK["extra_high_risk_asns"])
    if not isinstance(raw, dict):
        return cfg
    provider = str(raw.get("provider") or cfg["provider"]).strip().lower()
    known = {p[0] for p in PROVIDERS}
    cfg["provider"] = provider if provider in known else "abuseipdb"
    if "enabled" in raw:
        cfg["enabled"] = bool(raw.get("enabled"))
    try:
        cfg["abuse_threshold"] = max(0, min(100, int(raw.get("abuse_threshold", 25))))
    except (TypeError, ValueError):
        cfg["abuse_threshold"] = 25
    extra = raw.get("extra_high_risk_asns") or []
    if isinstance(extra, str):
        extra = re.split(r"[\s,;]+", extra)
    normalized: List[str] = []
    for item in extra:
        asn = _normalize_asn(item)
        if asn and asn not in normalized:
            normalized.append(asn)
    cfg["extra_high_risk_asns"] = normalized
    return cfg


def _normalize_logging(raw: Any) -> Dict[str, Any]:
    cfg = {
        "enabled": bool(DEFAULT_LOGGING["enabled"]),
        "max_mib": int(DEFAULT_LOGGING["max_mib"]),
        "levels": dict(DEFAULT_LOGGING["levels"]),
    }
    if not isinstance(raw, dict):
        return cfg
    if "enabled" in raw:
        cfg["enabled"] = bool(raw.get("enabled"))
    try:
        cfg["max_mib"] = max(0, int(raw.get("max_mib", 0) or 0))
    except (TypeError, ValueError):
        cfg["max_mib"] = 0
    levels = raw.get("levels")
    if isinstance(levels, dict):
        for key in ("info", "audit", "warn", "debug"):
            if key in levels:
                cfg["levels"][key] = bool(levels.get(key))
    else:
        # Flat keys fallback: info/audit/warn/debug at top level
        for key in ("info", "audit", "warn", "debug"):
            if key in raw:
                cfg["levels"][key] = bool(raw.get(key))
    return cfg


def _load_json_file(path: Path, normalizer) -> Dict[str, Any]:
    try:
        if path.is_file():
            return normalizer(json.loads(path.read_text(encoding="utf-8")))
    except Exception:
        pass
    return normalizer(None)


def _write_json(path: Path, payload: Dict[str, Any]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(payload, indent=2), encoding="utf-8")


def run_dialog(
    accounts: List[Dict[str, Any]],
    out_path: Path,
    route_risk: Dict[str, Any],
    route_path: Path,
    logging_cfg: Dict[str, Any],
    logging_path: Path,
) -> int:
    try:
        from PySide6.QtCore import Qt, QTimer
        from PySide6.QtWidgets import (
            QApplication,
            QCheckBox,
            QComboBox,
            QDialog,
            QFormLayout,
            QFrame,
            QFileDialog,
            QHBoxLayout,
            QLabel,
            QLineEdit,
            QPlainTextEdit,
            QPushButton,
            QScrollArea,
            QSpinBox,
            QTabWidget,
            QVBoxLayout,
            QWidget,
        )
    except ImportError:
        print("PySide6 is required for aes_settings_dialog.py", file=sys.stderr)
        return 2

    app = QApplication.instance() or QApplication(sys.argv)

    dialog = QDialog()
    dialog.setWindowTitle("AES Settings")
    dialog.setModal(True)
    dialog.setWindowFlag(Qt.WindowType.WindowStaysOnTopHint, True)
    dialog.setStyleSheet(STYLESHEET)
    dialog.setMinimumSize(960, 780)
    dialog.resize(1000, 840)

    root = QVBoxLayout(dialog)
    root.setSpacing(8)
    root.setContentsMargins(16, 10, 16, 14)

    root.addWidget(_make_logo_label(dialog, height=120), 0, Qt.AlignmentFlag.AlignHCenter)

    tabs = QTabWidget()
    root.addWidget(tabs, stretch=1)

    # ---- Tab 1: Mail settings ----
    mail_tab = QWidget()
    mail_layout = QVBoxLayout(mail_tab)
    mail_layout.setSpacing(8)
    mail_layout.setContentsMargins(8, 10, 8, 8)

    accounts_heading = QLabel("Scan accounts")
    accounts_heading.setObjectName("section")
    mail_layout.addWidget(accounts_heading)
    accounts_sub = QLabel(
        "Choose which mailboxes AES scans automatically. "
        "Under each account, enable Responses and/or In Cc as needed."
    )
    accounts_sub.setObjectName("subtitle")
    accounts_sub.setWordWrap(True)
    mail_layout.addWidget(accounts_sub)

    scroll = QScrollArea()
    scroll.setWidgetResizable(True)
    scroll.setFrameShape(QFrame.Shape.NoFrame)
    scroll_body = QWidget()
    scroll_layout = QVBoxLayout(scroll_body)
    scroll_layout.setSpacing(8)
    scroll_layout.setContentsMargins(0, 0, 4, 0)

    account_rows: List[Dict[str, Any]] = []
    for acc in accounts:
        smtp = str(acc.get("smtp") or "")
        display = str(acc.get("display") or smtp or "Account")
        label = display if not smtp or smtp.lower() == display.lower() else f"{display}  ·  {smtp}"

        card = QFrame()
        card.setObjectName("accountCard")
        card_layout = QVBoxLayout(card)
        card_layout.setSpacing(4)
        card_layout.setContentsMargins(10, 8, 10, 8)

        enabled_cb = QCheckBox(label)
        enabled_cb.setChecked(bool(acc.get("enabled", True)))
        enabled_cb.setProperty("store_id", str(acc.get("store_id") or ""))
        card_layout.addWidget(enabled_cb)

        opts = QHBoxLayout()
        opts.setContentsMargins(22, 0, 0, 0)
        responses_cb = QCheckBox("Responses")
        responses_cb.setObjectName("accountOpt")
        responses_cb.setChecked(bool(acc.get("responses", True)))
        responses_cb.setToolTip("Scan reply / response messages for this account.")
        in_cc_cb = QCheckBox("In Cc")
        in_cc_cb.setObjectName("accountOpt")
        in_cc_cb.setChecked(bool(acc.get("in_cc", True)))
        in_cc_cb.setToolTip("Scan messages where this account is on Cc.")
        opts.addWidget(responses_cb)
        opts.addWidget(in_cc_cb)
        opts.addStretch(1)
        card_layout.addLayout(opts)

        def _sync_opts(checked: bool, r_cb=responses_cb, c_cb=in_cc_cb) -> None:
            r_cb.setEnabled(checked)
            c_cb.setEnabled(checked)

        enabled_cb.toggled.connect(_sync_opts)
        _sync_opts(enabled_cb.isChecked())

        scroll_layout.addWidget(card)
        account_rows.append(
            {
                "enabled": enabled_cb,
                "responses": responses_cb,
                "in_cc": in_cc_cb,
            }
        )

    scroll_layout.addStretch(1)
    scroll.setWidget(scroll_body)
    mail_layout.addWidget(scroll, stretch=1)

    bulk = QHBoxLayout()
    all_on = QPushButton("All ON")
    all_on.setObjectName("secondaryBtn")
    all_off = QPushButton("All OFF")
    all_off.setObjectName("secondaryBtn")
    bulk.addWidget(all_on)
    bulk.addWidget(all_off)
    bulk.addStretch(1)
    mail_layout.addLayout(bulk)
    tabs.addTab(mail_tab, "Mail settings")

    # ---- Tab 2: ASN Risk ----
    asn_tab = QWidget()
    asn_layout = QVBoxLayout(asn_tab)
    asn_layout.setSpacing(8)
    asn_layout.setContentsMargins(8, 10, 8, 8)

    route_heading = QLabel("Route / ASN risk")
    route_heading.setObjectName("section")
    asn_layout.addWidget(route_heading)
    route_sub = QLabel(
        "Flag mail bounced through high-risk hops (+15 once). "
        "AbuseIPDB hop scores are used by default; add ASNs to always treat as high risk."
    )
    route_sub.setObjectName("subtitle")
    route_sub.setWordWrap(True)
    asn_layout.addWidget(route_sub)

    route_enabled = QCheckBox("Enable high-risk ASN / route scoring")
    route_enabled.setObjectName("plainCheck")
    route_enabled.setChecked(bool(route_risk.get("enabled", True)))
    asn_layout.addWidget(route_enabled)

    form = QFormLayout()
    form.setSpacing(8)
    provider_combo = QComboBox()
    for key, label in PROVIDERS:
        provider_combo.addItem(label, key)
    idx = provider_combo.findData(str(route_risk.get("provider") or "abuseipdb"))
    provider_combo.setCurrentIndex(max(0, idx))
    form.addRow("ASN / reputation source", provider_combo)

    threshold = QSpinBox()
    threshold.setRange(0, 100)
    threshold.setValue(int(route_risk.get("abuse_threshold", 25)))
    threshold.setToolTip("Hop AbuseIPDB confidence score at or above this value counts as high risk.")
    form.addRow("AbuseIPDB threshold", threshold)
    asn_layout.addLayout(form)

    asn_hint = QLabel("Extra high-risk ASNs (one per line, e.g. AS12345)")
    asn_hint.setObjectName("hint")
    asn_layout.addWidget(asn_hint)
    asn_edit = QPlainTextEdit()
    asn_edit.setPlainText("\n".join(route_risk.get("extra_high_risk_asns") or []))
    asn_edit.setMaximumHeight(220)
    asn_layout.addWidget(asn_edit)
    asn_layout.addStretch(1)
    tabs.addTab(asn_tab, "ASN Risk")

    # ---- Tab: Beacon Blocking ----
    beacon_path = _default_beacon_path()
    beacon_cfg = _load_json_file(beacon_path, _normalize_beacon_blocking)

    beacon_tab = QWidget()
    beacon_layout = QVBoxLayout(beacon_tab)
    beacon_layout.setSpacing(8)
    beacon_layout.setContentsMargins(8, 10, 8, 8)

    beacon_heading = QLabel("Tracking beacon blocking")
    beacon_heading.setObjectName("section")
    beacon_layout.addWidget(beacon_heading)
    beacon_sub = QLabel(
        "Tracking beacons are hidden 1\u00d71 images that phone home the moment you open a "
        "message, telling the sender when, where, and on what device you read it. "
        "When blocking is on, AES neutralises detected beacons in incoming mail before "
        "they can ping home."
    )
    beacon_sub.setObjectName("subtitle")
    beacon_sub.setWordWrap(True)
    beacon_layout.addWidget(beacon_sub)

    beacon_enabled = QCheckBox("Block incoming tracking beacons (prevent ping-home)")
    beacon_enabled.setObjectName("plainCheck")
    beacon_enabled.setChecked(bool(beacon_cfg.get("enabled", False)))
    beacon_layout.addWidget(beacon_enabled)

    beacon_form = QFormLayout()
    beacon_form.setSpacing(8)
    beacon_mode = QComboBox()
    beacon_mode.addItem("Defang — replace beacon with an inert placeholder (keeps layout)", "defang")
    beacon_mode.addItem("Strip — remove the beacon image tag completely", "strip")
    mode_idx = beacon_mode.findData(str(beacon_cfg.get("mode") or "defang"))
    beacon_mode.setCurrentIndex(max(0, mode_idx))
    beacon_form.addRow("Blocking method", beacon_mode)
    beacon_layout.addLayout(beacon_form)

    beacon_scope_all = QCheckBox("Apply to all scanned mail (untick = only senders on the per-sender block list)")
    beacon_scope_all.setObjectName("plainCheck")
    beacon_scope_all.setChecked(bool(beacon_cfg.get("block_all", True)))
    beacon_layout.addWidget(beacon_scope_all)

    beacon_wl_hint = QLabel("Never block beacons from these sender domains (one per line):")
    beacon_wl_hint.setObjectName("hint")
    beacon_layout.addWidget(beacon_wl_hint)
    beacon_wl_edit = QPlainTextEdit()
    beacon_wl_edit.setPlainText("\n".join(beacon_cfg.get("whitelist_domains") or []))
    beacon_wl_edit.setMaximumHeight(160)
    beacon_layout.addWidget(beacon_wl_edit)

    beacon_note = QLabel(
        "Per-sender rules come from the \u201cBlock beacons from sender\u201d button in AES footers "
        f"and are stored in {_local_geofooter() / 'aes_sender_rules.json'}. "
        "The quick-scan footer reports \u201cBeacons: X detected / Y blocked\u201d."
    )
    beacon_note.setObjectName("hint")
    beacon_note.setWordWrap(True)
    beacon_layout.addWidget(beacon_note)
    beacon_layout.addStretch(1)

    def sync_beacon_enabled(checked: bool) -> None:
        beacon_mode.setEnabled(checked)
        beacon_scope_all.setEnabled(checked)
        beacon_wl_edit.setEnabled(checked)

    beacon_enabled.toggled.connect(sync_beacon_enabled)
    sync_beacon_enabled(beacon_enabled.isChecked())
    tabs.addTab(beacon_tab, "Beacon Blocking")

    # ---- Tab: Sender Status ----
    sender_status_path = _default_sender_status_path()
    sender_status_cfg = _load_json_file(sender_status_path, _normalize_sender_status)

    sender_tab = QWidget()
    sender_layout = QVBoxLayout(sender_tab)
    sender_layout.setSpacing(8)
    sender_layout.setContentsMargins(8, 10, 8, 8)

    sender_heading = QLabel("Sender status")
    sender_heading.setObjectName("section")
    sender_layout.addWidget(sender_heading)
    sender_sub = QLabel(
        "The status strip shows Unknown, Known/Untrust or Known/Trust. "
        "A sender becomes Known once you have received more than this many "
        "emails from them; Trust is only ever your own decision, made with the "
        "\u201cTrust sender\u201d button in an AES footer. A Known sender you have "
        "not trusted counts as Untrust."
    )
    sender_sub.setObjectName("subtitle")
    sender_sub.setWordWrap(True)
    sender_layout.addWidget(sender_sub)

    sender_form = QFormLayout()
    sender_form.setSpacing(8)
    known_threshold = QSpinBox()
    known_threshold.setRange(0, 1000)
    known_threshold.setValue(int(sender_status_cfg.get("known_threshold", 5)))
    known_threshold.setToolTip(
        "Emails received from a sender before they count as Known. "
        "More than this value, not equal to it."
    )
    sender_form.addRow("Emails before Known", known_threshold)
    sender_layout.addLayout(sender_form)

    sender_note = QLabel(
        "Counting starts from when this build was installed, so senders you "
        "have known for years still begin at zero and will show as Unknown "
        f"until the count builds up. History is stored in "
        f"{_local_geofooter() / 'aes_sender_history.json'}. "
        "Trusting a sender also reduces their risk score by 33%."
    )
    sender_note.setObjectName("hint")
    sender_note.setWordWrap(True)
    sender_layout.addWidget(sender_note)
    try:
        from aes.settings_panels import build_trusted_senders_panel

        sender_layout.addWidget(build_trusted_senders_panel(sender_tab), 1)
    except Exception as exc:  # noqa: BLE001
        err = QLabel(f"Trusted-sender list unavailable: {exc}")
        err.setObjectName("hint")
        sender_layout.addWidget(err)
        sender_layout.addStretch(1)
    tabs.addTab(sender_tab, "Sender Status")

    # ---- Tab 3: External APIs ----
    try:
        from aes.secret_store import abuseipdb_status, get_abuseipdb_base_url
    except ImportError:
        abuseipdb_status = None  # type: ignore
        get_abuseipdb_base_url = None  # type: ignore

    api_tab = QWidget()
    api_layout = QVBoxLayout(api_tab)
    api_layout.setSpacing(8)
    api_layout.setContentsMargins(8, 10, 8, 8)

    api_heading = QLabel("External APIs")
    api_heading.setObjectName("section")
    api_layout.addWidget(api_heading)
    api_sub = QLabel(
        "Configure third-party API endpoints and keys used by AES. "
        "Keys are stored with Windows DPAPI for your user only."
    )
    api_sub.setObjectName("subtitle")
    api_sub.setWordWrap(True)
    api_layout.addWidget(api_sub)

    abuse_heading = QLabel("AbuseIPDB")
    abuse_heading.setObjectName("section")
    api_layout.addWidget(abuse_heading)

    api_status_lbl = QLabel("Status: …")
    api_status_lbl.setObjectName("hint")
    api_status_lbl.setWordWrap(True)
    api_layout.addWidget(api_status_lbl)

    def refresh_abuseipdb_summary() -> None:
        status = (
            abuseipdb_status()
            if abuseipdb_status
            else {"configured": False, "hint": "", "source": "none"}
        )
        base = get_abuseipdb_base_url() if get_abuseipdb_base_url else "https://api.abuseipdb.com/api/v2"
        if status.get("configured"):
            api_status_lbl.setText(
                f"URL: {base}\nKey: configured ({status.get('hint')}, via {status.get('source')})"
            )
        else:
            api_status_lbl.setText(f"URL: {base}\nKey: not configured")

    refresh_abuseipdb_summary()

    abuse_btn = QPushButton("AbuseIPDB API…")
    abuse_btn.setObjectName("secondaryBtn")
    abuse_btn.setToolTip("Open the AbuseIPDB API URL and key dialog.")
    api_layout.addWidget(abuse_btn, 0, Qt.AlignmentFlag.AlignLeft)

    def open_abuseipdb_dialog() -> None:
        run_abuseipdb_dialog(dialog)
        refresh_abuseipdb_summary()

    abuse_btn.clicked.connect(open_abuseipdb_dialog)
    try:
        from aes.settings_panels import build_threat_intel_panel

        api_layout.addWidget(build_threat_intel_panel(api_tab))
    except Exception as exc:  # noqa: BLE001
        err = QLabel(f"Threat-intel settings unavailable: {exc}")
        err.setObjectName("hint")
        api_layout.addWidget(err)
    api_layout.addStretch(1)
    api_scroll = QScrollArea()
    api_scroll.setWidgetResizable(True)
    api_scroll.setFrameShape(QFrame.Shape.NoFrame)
    api_scroll.setWidget(api_tab)
    tabs.addTab(api_scroll, "External APIs")

    # ---- Tab 4: Logging ----
    log_tab = QWidget()
    log_layout = QVBoxLayout(log_tab)
    log_layout.setSpacing(8)
    log_layout.setContentsMargins(8, 10, 8, 8)

    log_heading = QLabel("Logging")
    log_heading.setObjectName("section")
    log_layout.addWidget(log_heading)
    log_sub = QLabel(
        "Control AES file logging. Turn logging off to silence all levels, "
        "or enable specific categories below."
    )
    log_sub.setObjectName("subtitle")
    log_sub.setWordWrap(True)
    log_layout.addWidget(log_sub)

    log_enabled = QCheckBox("Logging on")
    log_enabled.setObjectName("plainCheck")
    log_enabled.setChecked(bool(logging_cfg.get("enabled", True)))
    log_enabled.setToolTip("Master switch for AES log file output.")
    log_layout.addWidget(log_enabled)

    size_row = QHBoxLayout()
    size_row.setSpacing(8)
    size_lbl = QLabel("Max log size")
    size_spin = QSpinBox()
    size_spin.setRange(0, 10240)
    size_spin.setSingleStep(1)
    size_spin.setSuffix(" MiB")
    size_spin.setValue(int(logging_cfg.get("max_mib") or 0))
    size_spin.setToolTip(
        "Maximum size of VBA_Log.txt in mebibytes. "
        "0 is unlimited. When the file grows past this, the oldest lines are removed."
    )
    size_hint = QLabel("0 = unlimited. Oldest lines are removed once the file passes this size.")
    size_hint.setObjectName("hint")
    size_row.addWidget(size_lbl)
    size_row.addWidget(size_spin)
    size_row.addWidget(size_hint, stretch=1)
    log_layout.addLayout(size_row)

    levels_heading = QLabel("Levels to write")
    levels_heading.setObjectName("section")
    log_layout.addWidget(levels_heading)

    level_checks: Dict[str, QCheckBox] = {}
    levels = logging_cfg.get("levels") or {}
    for key in ("info", "audit", "warn", "debug"):
        cb = QCheckBox(key.capitalize())
        cb.setObjectName("levelCheck")
        cb.setChecked(bool(levels.get(key, DEFAULT_LOGGING["levels"][key])))
        cb.setToolTip(LOG_LEVEL_HELP[key])
        log_layout.addWidget(cb)
        help_lbl = QLabel(LOG_LEVEL_HELP[key])
        help_lbl.setObjectName("hint")
        help_lbl.setWordWrap(True)
        log_layout.addWidget(help_lbl)
        level_checks[key] = cb

    log_hint = QLabel(
        "Recommended: Info + Audit + Warn on; Debug only while troubleshooting.\n"
        f"Log file: {_local_geofooter() / 'Logs' / 'VBA_Log.txt'}"
    )
    log_hint.setObjectName("hint")
    log_hint.setWordWrap(True)
    log_layout.addWidget(log_hint)
    log_layout.addStretch(1)

    def sync_level_enabled(checked: bool) -> None:
        for cb in level_checks.values():
            cb.setEnabled(checked)

    log_enabled.toggled.connect(sync_level_enabled)
    sync_level_enabled(log_enabled.isChecked())
    tabs.addTab(log_tab, "Logging")

    # ---- Tab 5: Console ----
    console_tab = QWidget()
    console_layout = QVBoxLayout(console_tab)
    console_layout.setSpacing(6)
    console_layout.setContentsMargins(8, 8, 8, 8)

    path_row = QHBoxLayout()
    path_row.setSpacing(8)
    path_lbl = QLabel("Log file")
    path_lbl.setObjectName("hint")
    log_path = _vba_log_path()
    console_path_edit = QLineEdit(str(log_path))
    console_path_edit.setObjectName("consolePathEdit")
    console_path_edit.setReadOnly(True)
    console_path_edit.setToolTip("Current AES log file. Use Browse… to choose another.")
    browse_btn = QPushButton("Browse…")
    browse_btn.setObjectName("secondaryBtn")
    browse_btn.setToolTip("Choose a log file to view in the console.")
    path_row.addWidget(path_lbl)
    path_row.addWidget(console_path_edit, stretch=1)
    path_row.addWidget(browse_btn)
    console_layout.addLayout(path_row)

    console_view = QPlainTextEdit()
    console_view.setObjectName("consoleView")
    console_view.setReadOnly(True)
    console_view.setLineWrapMode(QPlainTextEdit.LineWrapMode.NoWrap)
    console_view.setMaximumBlockCount(2000)
    console_view.setMinimumHeight(320)
    console_layout.addWidget(console_view, stretch=1)

    console_bar = QHBoxLayout()
    follow_cb = QCheckBox("Follow")
    follow_cb.setObjectName("plainCheck")
    follow_cb.setChecked(True)
    follow_cb.setToolTip("Keep scrolled to the latest messages.")
    refresh_btn = QPushButton("Refresh")
    refresh_btn.setObjectName("secondaryBtn")
    clear_btn = QPushButton("Clear view")
    clear_btn.setObjectName("secondaryBtn")
    clear_btn.setToolTip("Clear the console view only — does not delete the log file.")
    console_bar.addWidget(follow_cb)
    console_bar.addStretch(1)
    console_bar.addWidget(refresh_btn)
    console_bar.addWidget(clear_btn)
    console_layout.addLayout(console_bar)

    console_state = {"text": "", "path": log_path}

    def refresh_console(force: bool = False) -> None:
        path = Path(console_path_edit.text().strip() or str(console_state["path"]))
        if not path.is_file():
            path = _vba_log_path()
            console_path_edit.setText(str(path))
        console_state["path"] = path
        text = _read_log_tail(path, max_lines=800)
        if not force and text == console_state["text"]:
            return
        console_state["text"] = text
        bar = console_view.verticalScrollBar()
        at_bottom = bar.value() >= bar.maximum() - 4
        console_view.setPlainText(text)
        if follow_cb.isChecked() or at_bottom:
            bar.setValue(bar.maximum())

    def browse_log_file() -> None:
        start_dir = str(console_state["path"].parent if console_state["path"] else _vba_log_path().parent)
        chosen, _ = QFileDialog.getOpenFileName(
            dialog,
            "Select AES log file",
            start_dir,
            "Log files (*.txt *.log);;All files (*.*)",
        )
        if chosen:
            console_path_edit.setText(chosen)
            console_state["path"] = Path(chosen)
            refresh_console(force=True)

    def clear_console_view() -> None:
        console_state["text"] = ""
        console_view.clear()

    browse_btn.clicked.connect(browse_log_file)
    refresh_btn.clicked.connect(lambda: refresh_console(force=True))
    clear_btn.clicked.connect(clear_console_view)

    console_timer = QTimer(dialog)
    console_timer.setInterval(1500)

    def on_console_tick() -> None:
        if tabs.currentWidget() is console_tab:
            refresh_console(force=False)

    console_timer.timeout.connect(on_console_tick)

    def on_tab_changed(index: int) -> None:
        if tabs.widget(index) is console_tab:
            refresh_console(force=True)
            if not console_timer.isActive():
                console_timer.start()
        elif console_timer.isActive():
            console_timer.stop()

    tabs.currentChanged.connect(on_tab_changed)
    tabs.addTab(console_tab, "Console")

    buttons = QHBoxLayout()
    buttons.addStretch(1)
    save_btn = QPushButton("Save")
    save_btn.setObjectName("saveBtn")
    save_btn.setDefault(True)
    close_btn = QPushButton("Close")
    close_btn.setObjectName("cancelBtn")
    buttons.addWidget(save_btn)
    buttons.addWidget(close_btn)
    root.addLayout(buttons)

    result: Dict[str, Any] = {
        "cancelled": True,
        "accounts": [],
        "route_risk": route_risk,
        "logging": logging_cfg,
    }

    def collect_route() -> Dict[str, Any]:
        asns: List[str] = []
        for line in asn_edit.toPlainText().splitlines():
            asn = _normalize_asn(line)
            if asn and asn not in asns:
                asns.append(asn)
        return {
            "provider": str(provider_combo.currentData() or "abuseipdb"),
            "abuse_threshold": int(threshold.value()),
            "extra_high_risk_asns": asns,
            "enabled": bool(route_enabled.isChecked()),
        }

    def collect_logging() -> Dict[str, Any]:
        return {
            "enabled": bool(log_enabled.isChecked()),
            "max_mib": int(size_spin.value()),
            "levels": {key: bool(cb.isChecked()) for key, cb in level_checks.items()},
        }

    def collect_beacon() -> Dict[str, Any]:
        return _normalize_beacon_blocking(
            {
                "enabled": bool(beacon_enabled.isChecked()),
                "mode": str(beacon_mode.currentData() or "defang"),
                "block_all": bool(beacon_scope_all.isChecked()),
                "whitelist_domains": beacon_wl_edit.toPlainText(),
            }
        )

    def on_all(enabled: bool) -> None:
        for row in account_rows:
            row["enabled"].setChecked(enabled)

    def on_cancel() -> None:
        result["cancelled"] = True
        result["accounts"] = []
        result["route_risk"] = route_risk
        result["logging"] = logging_cfg
        dialog.reject()

    def on_save() -> None:
        out_accounts = []
        for row in account_rows:
            enabled_cb = row["enabled"]
            out_accounts.append(
                {
                    "store_id": str(enabled_cb.property("store_id") or ""),
                    "enabled": bool(enabled_cb.isChecked()),
                    "responses": bool(row["responses"].isChecked()),
                    "in_cc": bool(row["in_cc"].isChecked()),
                }
            )
        route_out = collect_route()
        logging_out = collect_logging()
        beacon_out = collect_beacon()
        sender_status_out = {"known_threshold": int(known_threshold.value())}
        result["cancelled"] = False
        result["accounts"] = out_accounts
        result["route_risk"] = route_out
        result["logging"] = logging_out
        result["beacon_blocking"] = beacon_out
        result["sender_status"] = sender_status_out
        try:
            _write_json(route_path, route_out)
        except Exception as exc:  # noqa: BLE001
            print(f"Failed to write route risk config: {exc}", file=sys.stderr)
        try:
            _write_json(logging_path, logging_out)
        except Exception as exc:  # noqa: BLE001
            print(f"Failed to write logging config: {exc}", file=sys.stderr)
        try:
            _write_json(beacon_path, beacon_out)
        except Exception as exc:  # noqa: BLE001
            print(f"Failed to write beacon blocking config: {exc}", file=sys.stderr)
        try:
            _write_json(sender_status_path, sender_status_out)
        except Exception as exc:  # noqa: BLE001
            print(f"Failed to write sender status config: {exc}", file=sys.stderr)

        dialog.accept()

    all_on.clicked.connect(lambda: on_all(True))
    all_off.clicked.connect(lambda: on_all(False))
    close_btn.clicked.connect(on_cancel)
    save_btn.clicked.connect(on_save)

    dialog.exec()
    _write_json(out_path, result)
    return 0


def main() -> int:
    parser = argparse.ArgumentParser(description="AES settings dialog")
    parser.add_argument("--accounts", required=True, help="Input JSON path with account list")
    parser.add_argument("--out", required=True, help="Output JSON path")
    parser.add_argument("--route", default="", help="Path to aes_risk_route.json")
    parser.add_argument("--logging", default="", help="Path to aes_logging.json")
    args = parser.parse_args()

    accounts_path = Path(args.accounts)
    out_path = Path(args.out)
    route_path = Path(args.route) if args.route else _default_route_path()
    logging_path = Path(args.logging) if args.logging else _default_logging_path()

    try:
        data = json.loads(accounts_path.read_text(encoding="utf-8"))
        accounts = list(data.get("accounts") or [])
        if not accounts:
            print("No accounts in input JSON", file=sys.stderr)
            _write_json(out_path, {"cancelled": True, "accounts": [], "error": "no accounts"})
            return 1

        route_risk = _normalize_route_risk(data.get("route_risk"))
        if route_path.is_file():
            route_risk = _load_json_file(route_path, _normalize_route_risk)

        logging_cfg = _normalize_logging(data.get("logging"))
        if logging_path.is_file():
            logging_cfg = _load_json_file(logging_path, _normalize_logging)

        return run_dialog(
            accounts, out_path, route_risk, route_path, logging_cfg, logging_path
        )
    except Exception as exc:  # noqa: BLE001
        print(f"aes_settings_dialog failed: {exc}", file=sys.stderr)
        try:
            _write_json(
                out_path,
                {"cancelled": True, "accounts": [], "error": str(exc)},
            )
        except Exception:
            pass
        return 1


if __name__ == "__main__":
    sys.exit(main())
