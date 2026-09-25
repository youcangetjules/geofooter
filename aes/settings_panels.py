#!/usr/bin/env python3
"""Extra AES Settings panels: threat-intel keys, trusted senders, witticisms."""
from __future__ import annotations

import json
import os
import threading
from datetime import datetime, timezone
from pathlib import Path
from typing import Any, Dict, List

from PySide6.QtCore import QObject, Qt, Signal
from PySide6.QtWidgets import (
    QAbstractItemView,
    QCheckBox,
    QComboBox,
    QGridLayout,
    QHBoxLayout,
    QHeaderView,
    QLabel,
    QLineEdit,
    QListWidget,
    QMessageBox,
    QPushButton,
    QTableWidget,
    QTableWidgetItem,
    QVBoxLayout,
    QWidget,
)


def _geofooter_dir() -> Path:
    local = os.environ.get("LOCALAPPDATA", "")
    base = Path(local) / "GeoFooter" if local else Path(r"C:\GeoFooter")
    base.mkdir(parents=True, exist_ok=True)
    return base


# --------------------------------------------------------------------------- #
# Threat-intel keys
# --------------------------------------------------------------------------- #

class _ProbeBridge(QObject):
    done = Signal(str, bool, str)


def _api_config_path() -> Path:
    return _geofooter_dir() / "aes_api.json"


def _read_api_config() -> Dict[str, Any]:
    try:
        data = json.loads(_api_config_path().read_text(encoding="utf-8"))
        return data if isinstance(data, dict) else {}
    except Exception:
        return {}


def _write_api_config(updates: Dict[str, Any]) -> None:
    data = _read_api_config()
    data.update(updates)
    path = _api_config_path()
    tmp = path.with_suffix(".json.tmp")
    tmp.write_text(json.dumps(data, indent=2), encoding="utf-8")
    os.replace(tmp, path)


def build_threat_intel_panel(parent: QWidget) -> QWidget:
    import aes.threat_intel as ti
    from aes.secret_store import secret_hint, set_secret

    panel = QWidget(parent)
    outer = QVBoxLayout(panel)
    outer.setContentsMargins(0, 6, 0, 0)
    outer.setSpacing(6)

    heading = QLabel("Threat intelligence")
    heading.setObjectName("section")
    outer.addWidget(heading)
    sub = QLabel(
        "Used for hop IPs, sender and link domains, links and attachment hashes. "
        "Shodan InternetDB and the domain blocklists need no key. Nothing is "
        "uploaded: files are checked by hash only, and links are looked up — "
        "urlscan.io is the one exception, submitting suspicious links during Deep Scans."
    )
    sub.setObjectName("hint")
    sub.setWordWrap(True)
    outer.addWidget(sub)

    grid = QGridLayout()
    grid.setHorizontalSpacing(8)
    grid.setVerticalSpacing(4)
    outer.addLayout(grid)

    bridge = _ProbeBridge(panel)
    status_labels: Dict[str, QLabel] = {}
    edits: Dict[str, QLineEdit] = {}

    def set_status(provider: str, ok: bool, msg: str) -> None:
        lbl = status_labels.get(provider)
        if lbl is not None:
            lbl.setText(("\u2714 " if ok else "\u2716 ") + msg)
            lbl.setStyleSheet(f"color: {'#1b7f3b' if ok else '#b3261e'};")

    bridge.done.connect(set_status)

    def probe_async(provider: str, key: str | None = None) -> None:
        status_labels[provider].setText("Checking\u2026")
        status_labels[provider].setStyleSheet("color: #666;")

        def work() -> None:
            ok, msg = ti.probe_provider(provider, key)
            bridge.done.emit(provider, ok, msg)

        threading.Thread(target=work, daemon=True).start()

    def configured_text(provider: str) -> str:
        key_name = ti.PROVIDERS[provider].get("key")
        if not key_name:
            return "no key needed"
        if os.environ.get(f"AES_{key_name.upper()}_API_KEY"):
            return "configured (environment variable)"
        hint = secret_hint(key_name) if ti.provider_key(provider) else ""
        return f"configured ({hint})" if hint else "not configured"

    row = 0
    for provider, meta in ti.PROVIDERS.items():
        name_lbl = QLabel(meta["label"])
        if meta.get("signup"):
            name_lbl.setText(f"<a href='{meta['signup']}'>{meta['label']}</a>")
            name_lbl.setOpenExternalLinks(True)
            name_lbl.setToolTip("Open the page where this API key is issued")
        grid.addWidget(name_lbl, row, 0)

        status = QLabel(configured_text(provider))
        status.setObjectName("hint")
        status_labels[provider] = status

        if meta.get("key"):
            edit = QLineEdit()
            edit.setEchoMode(QLineEdit.EchoMode.Password)
            edit.setPlaceholderText("Paste key to replace" if ti.provider_key(provider) else "Paste API key")
            edit.setMinimumWidth(230)
            edits[provider] = edit
            grid.addWidget(edit, row, 1)

            save_btn = QPushButton("Save && test")
            save_btn.setObjectName("secondaryBtn")
            clear_btn = QPushButton("Clear")
            clear_btn.setObjectName("secondaryBtn")

            def on_save(_=False, p=provider, e=edit, key_name=meta["key"]) -> None:
                typed = e.text().strip()
                if typed:
                    set_secret(key_name, typed)
                    e.clear()
                    e.setPlaceholderText("Paste key to replace")
                if ti.provider_key(p):
                    probe_async(p)
                else:
                    status_labels[p].setText("not configured")

            def on_clear(_=False, p=provider, e=edit, key_name=meta["key"]) -> None:
                if QMessageBox.question(panel, "Clear key", f"Remove the stored {ti.PROVIDERS[p]['label']} key?") \
                        != QMessageBox.StandardButton.Yes:
                    return
                set_secret(key_name, "")
                e.clear()
                e.setPlaceholderText("Paste API key")
                status_labels[p].setText(configured_text(p))
                status_labels[p].setStyleSheet("")

            save_btn.clicked.connect(on_save)
            edit.returnPressed.connect(on_save)
            clear_btn.clicked.connect(on_clear)
            grid.addWidget(save_btn, row, 2)
            grid.addWidget(clear_btn, row, 3)
        else:
            test_btn = QPushButton("Test")
            test_btn.setObjectName("secondaryBtn")
            test_btn.clicked.connect(lambda _=False, p=provider: probe_async(p))
            grid.addWidget(QLabel("\u2014"), row, 1)
            grid.addWidget(test_btn, row, 2)
        grid.addWidget(status, row, 4)
        row += 1
    grid.setColumnStretch(4, 1)

    vis_lbl = QLabel("urlscan.io visibility (Deep Scan)")
    vis_combo = QComboBox()
    vis_combo.setMinimumWidth(230)
    for value, label in (("private", "Private (only you)"),
                         ("unlisted", "Unlisted (you + vetted researchers)"),
                         ("public", "Public (anyone — avoid for real mail)")):
        vis_combo.addItem(label, value)
    idx = vis_combo.findData(str(_read_api_config().get("urlscan_visibility") or "private"))
    vis_combo.setCurrentIndex(max(0, idx))
    vis_combo.currentIndexChanged.connect(
        lambda _i: _write_api_config({"urlscan_visibility": vis_combo.currentData()})
    )
    grid.addWidget(vis_lbl, row, 0)
    grid.addWidget(vis_combo, row, 1)

    test_all = QPushButton("Test all configured")
    test_all.setObjectName("secondaryBtn")
    test_all.clicked.connect(
        lambda: [probe_async(p) for p in ti.PROVIDERS if ti.provider_enabled(p)]
    )
    outer.addWidget(test_all, 0, Qt.AlignmentFlag.AlignLeft)
    return panel


# --------------------------------------------------------------------------- #
# Trusted / untrusted senders
# --------------------------------------------------------------------------- #

RULE_KEYS = (
    "block_attachments",
    "block_beacons",
    "allow_beacons",
    "full_no_trust",
    "trusted",
    "untrusted",
)


def _rules_path() -> Path:
    return _geofooter_dir() / "aes_sender_rules.json"


def load_rules() -> Dict[str, Any]:
    data: Dict[str, Any] = {}
    try:
        raw = json.loads(_rules_path().read_text(encoding="utf-8"))
        if isinstance(raw, dict):
            data = raw
    except Exception:
        pass
    for key in RULE_KEYS:
        if not isinstance(data.get(key), list):
            data[key] = []
    return data


def save_rules(rules: Dict[str, Any]) -> None:
    rules["updated"] = datetime.now(timezone.utc).strftime("%Y-%m-%d %H:%M:%S UTC")
    path = _rules_path()
    tmp = path.with_suffix(".json.tmp")
    tmp.write_text(json.dumps(rules, indent=2), encoding="utf-8")
    os.replace(tmp, path)


def set_sender_status(who: str, status: str) -> None:
    """status: 'trusted' | 'untrusted' | 'none'. Mirrors aes_action_handler's consistency rules."""
    who = who.strip().lower()
    if not who:
        return
    identities = [who]
    if "@" in who:
        domain = who.rsplit("@", 1)[-1]
        if domain and domain not in identities:
            identities.append(domain)
    rules = load_rules()
    idset = set(identities)
    rules["trusted"] = [e for e in rules["trusted"] if e not in idset]
    rules["untrusted"] = [e for e in rules["untrusted"] if e not in idset]
    if status == "trusted":
        rules["trusted"].append(who)
        rules["block_attachments"] = [e for e in rules["block_attachments"] if e not in idset]
        rules["block_beacons"] = [e for e in rules["block_beacons"] if e not in idset]
        rules["allow_beacons"] = [e for e in rules.get("allow_beacons") or [] if e not in idset]
        rules["full_no_trust"] = [e for e in rules.get("full_no_trust") or [] if e not in idset]
    elif status == "untrusted":
        rules["untrusted"].append(who)
    save_rules(rules)


def _sender_history() -> Dict[str, Any]:
    try:
        data = json.loads((_geofooter_dir() / "aes_sender_history.json").read_text(encoding="utf-8"))
        return (data or {}).get("senders") or {}
    except Exception:
        return {}


def build_trusted_senders_panel(parent: QWidget) -> QWidget:
    panel = QWidget(parent)
    layout = QVBoxLayout(panel)
    layout.setContentsMargins(0, 6, 0, 0)
    layout.setSpacing(6)

    heading = QLabel("Trusted and untrusted senders")
    heading.setObjectName("section")
    layout.addWidget(heading)
    sub = QLabel(
        "Senders you marked with the Trust / Untrust footer buttons. Changes are "
        "saved immediately and apply from the next scan of mail from that sender."
    )
    sub.setObjectName("hint")
    sub.setWordWrap(True)
    layout.addWidget(sub)

    top = QHBoxLayout()
    show_combo = QComboBox()
    show_combo.addItem("Trusted", "trusted")
    show_combo.addItem("Untrusted", "untrusted")
    show_combo.addItem("Both", "both")
    search = QLineEdit()
    search.setPlaceholderText("Filter by address or domain\u2026")
    count_lbl = QLabel("")
    count_lbl.setObjectName("hint")
    top.addWidget(QLabel("Show"))
    top.addWidget(show_combo)
    top.addWidget(search, 1)
    top.addWidget(count_lbl)
    layout.addLayout(top)

    table = QTableWidget(0, 5)
    table.setHorizontalHeaderLabels(["Sender", "Status", "Emails", "First seen", "Last seen"])
    table.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectRows)
    table.setSelectionMode(QAbstractItemView.SelectionMode.ExtendedSelection)
    table.setEditTriggers(QAbstractItemView.EditTrigger.NoEditTriggers)
    table.setSortingEnabled(True)
    table.verticalHeader().setVisible(False)
    hdr = table.horizontalHeader()
    hdr.setSectionResizeMode(0, QHeaderView.ResizeMode.Stretch)
    for col in range(1, 5):
        hdr.setSectionResizeMode(col, QHeaderView.ResizeMode.ResizeToContents)
    table.setMinimumHeight(220)
    layout.addWidget(table, 1)

    btns = QHBoxLayout()
    untrust_btn = QPushButton("Move to Untrusted")
    trust_btn = QPushButton("Move to Trusted")
    remove_btn = QPushButton("Remove from lists")
    for b in (untrust_btn, trust_btn, remove_btn):
        b.setObjectName("secondaryBtn")
        btns.addWidget(b)
    btns.addStretch(1)
    add_edit = QLineEdit()
    add_edit.setPlaceholderText("name@domain.com or domain.com")
    add_btn = QPushButton("Add as Trusted")
    add_btn.setObjectName("secondaryBtn")
    btns.addWidget(add_edit)
    btns.addWidget(add_btn)
    layout.addLayout(btns)

    class _NumItem(QTableWidgetItem):
        def __lt__(self, other: QTableWidgetItem) -> bool:  # type: ignore[override]
            try:
                return int(self.text() or 0) < int(other.text() or 0)
            except ValueError:
                return super().__lt__(other)

    def refresh() -> None:
        rules = load_rules()
        history = _sender_history()
        mode = show_combo.currentData()
        needle = search.text().strip().lower()
        rows: List[tuple] = []
        for status in ("trusted", "untrusted"):
            if mode not in (status, "both"):
                continue
            for who in rules[status]:
                if needle and needle not in who:
                    continue
                h = history.get(who) or {}
                if not h and "@" not in who:
                    matches = [v for k, v in history.items() if k.endswith("@" + who) or k.endswith("." + who)]
                    if matches:
                        h = {
                            "count": sum(int(m.get("count") or 0) for m in matches),
                            "first_seen": min(str(m.get("first_seen") or "") for m in matches),
                            "last_seen": max(str(m.get("last_seen") or "") for m in matches),
                        }
                rows.append((who, status, h))
        table.setSortingEnabled(False)
        table.setRowCount(len(rows))
        for r, (who, status, h) in enumerate(rows):
            table.setItem(r, 0, QTableWidgetItem(who))
            st = QTableWidgetItem("Trusted" if status == "trusted" else "Untrusted")
            st.setForeground(Qt.GlobalColor.darkGreen if status == "trusted" else Qt.GlobalColor.darkRed)
            table.setItem(r, 1, st)
            table.setItem(r, 2, _NumItem(str(h.get("count") or 0)))
            table.setItem(r, 3, QTableWidgetItem(str(h.get("first_seen") or "\u2014")))
            table.setItem(r, 4, QTableWidgetItem(str(h.get("last_seen") or "\u2014")))
        table.setSortingEnabled(True)
        count_lbl.setText(
            f"{len(rules['trusted'])} trusted, {len(rules['untrusted'])} untrusted"
        )

    def selected() -> List[str]:
        rows = sorted({i.row() for i in table.selectedIndexes()})
        return [table.item(r, 0).text() for r in rows if table.item(r, 0)]

    def apply(status: str) -> None:
        who_list = selected()
        if not who_list:
            QMessageBox.information(panel, "Senders", "Select one or more senders first.")
            return
        for who in who_list:
            set_sender_status(who, status)
        refresh()

    untrust_btn.clicked.connect(lambda: apply("untrusted"))
    trust_btn.clicked.connect(lambda: apply("trusted"))
    remove_btn.clicked.connect(lambda: apply("none"))

    def on_add() -> None:
        who = add_edit.text().strip().lower()
        if not who or "." not in who or " " in who:
            QMessageBox.information(panel, "Senders", "Enter an email address or a domain.")
            return
        set_sender_status(who, "trusted")
        add_edit.clear()
        refresh()

    add_btn.clicked.connect(on_add)
    add_edit.returnPressed.connect(on_add)
    show_combo.currentIndexChanged.connect(lambda _i: refresh())
    search.textChanged.connect(lambda _t: refresh())
    refresh()
    return panel


def build_witticisms_panel(parent: QWidget) -> QWidget:
    """Edit footer one-liners and ask Poe for a new one."""
    import aes.witticisms as wit
    from aes.secret_store import get_poe_api_key, poe_status, set_poe_api_key

    stored = wit.load_witticisms()
    banks: Dict[str, List[str]] = {
        level: list(stored["lines"].get(level) or []) for level in wit.LEVELS
    }
    current = {"level": "LOW"}

    panel = QWidget(parent)
    outer = QVBoxLayout(panel)
    outer.setContentsMargins(8, 10, 8, 8)
    outer.setSpacing(8)

    heading = QLabel("Witticisms")
    heading.setObjectName("section")
    outer.addWidget(heading)
    sub = QLabel(
        "These are the cheeky lines in the top-right of the scan footer. "
        "AES picks one at random for the mail's risk band. "
        "Click Save to keep additions and removals. An empty band uses the built-in lines."
    )
    sub.setObjectName("subtitle")
    sub.setWordWrap(True)
    outer.addWidget(sub)

    band_row = QHBoxLayout()
    band_lbl = QLabel("Risk band")
    band_combo = QComboBox()
    for level in wit.LEVELS:
        band_combo.addItem(wit.LEVEL_LABELS[level], level)
    band_row.addWidget(band_lbl)
    band_row.addWidget(band_combo, stretch=1)
    outer.addLayout(band_row)

    lines = QListWidget()
    lines.setSelectionMode(QAbstractItemView.SelectionMode.ExtendedSelection)
    lines.setToolTip("Click a line to edit it. AES may show any of these for the selected risk band.")
    outer.addWidget(lines, stretch=1)

    edit_row = QHBoxLayout()
    line_edit = QLineEdit()
    line_edit.setPlaceholderText("Type a new line")
    add_btn = QPushButton("Add")
    add_btn.setObjectName("secondaryBtn")
    remove_btn = QPushButton("Remove")
    remove_btn.setObjectName("secondaryBtn")
    edit_row.addWidget(line_edit, stretch=1)
    edit_row.addWidget(add_btn)
    edit_row.addWidget(remove_btn)
    outer.addLayout(edit_row)

    poe_heading = QLabel("Suggest a new one")
    poe_heading.setObjectName("section")
    outer.addWidget(poe_heading)
    poe_sub = QLabel(
        "Uses Poe (poe.com/api/keys). The key is stored with Windows DPAPI "
        "for your user when you click Save."
    )
    poe_sub.setObjectName("hint")
    poe_sub.setWordWrap(True)
    outer.addWidget(poe_sub)

    key_row = QHBoxLayout()
    key_edit = QLineEdit()
    key_edit.setEchoMode(QLineEdit.EchoMode.Password)
    key_edit.setPlaceholderText("Poe API key")
    model_lbl = QLabel("Model")
    model_combo = QComboBox()
    model_combo.setToolTip("Poe model used when suggesting a new line")
    model_combo.setMinimumWidth(220)
    saved_model = str(stored.get("model") or wit.DEFAULT_MODEL).strip() or wit.DEFAULT_MODEL
    model_names = list(wit.POE_MODELS)
    if saved_model not in model_names:
        model_names.insert(0, saved_model)
    for name in model_names:
        model_combo.addItem(name, name)
    model_combo.setCurrentIndex(max(0, model_combo.findData(saved_model)))
    key_row.addWidget(key_edit, stretch=1)
    key_row.addWidget(model_lbl)
    key_row.addWidget(model_combo)
    outer.addLayout(key_row)

    status = poe_status()
    if status.get("configured"):
        key_edit.setPlaceholderText(
            f"Current key {status.get('hint')} — paste a new key to replace"
        )
    clear_key = QCheckBox("Clear stored Poe key on Save")
    clear_key.setObjectName("plainCheck")
    clear_key.setEnabled(bool(status.get("configured")) and status.get("source") != "env")
    outer.addWidget(clear_key)

    suggest_row = QHBoxLayout()
    suggest_btn = QPushButton("Suggest a new one")
    suggest_btn.setObjectName("secondaryBtn")
    test_btn = QPushButton("Test Connection")
    test_btn.setObjectName("secondaryBtn")
    test_btn.setToolTip("Check the Poe key and that the selected model is available")
    suggest_lbl = QLabel("")
    suggest_lbl.setObjectName("hint")
    suggest_lbl.setWordWrap(True)
    suggest_row.addWidget(suggest_btn)
    suggest_row.addWidget(test_btn)
    suggest_row.addWidget(suggest_lbl, stretch=1)
    outer.addLayout(suggest_row)

    def read_list() -> List[str]:
        return [lines.item(i).text() for i in range(lines.count())]

    def fill(level: str) -> None:
        lines.clear()
        for text in banks.get(level) or []:
            lines.addItem(text)

    def flush() -> None:
        banks[current["level"]] = read_list()

    fill("LOW")

    editing = {"row": -1}
    filling_edit = {"on": False}

    def set_edit_mode(row: int) -> None:
        editing["row"] = row if row >= 0 else -1
        editing_now = editing["row"] >= 0
        add_btn.setText("Edit" if editing_now else "Add")
        add_btn.setToolTip(
            "Replace the selected line" if editing_now else "Add this line to the band"
        )

    def on_band(index: int) -> None:
        new_level = str(band_combo.itemData(index) or "LOW")
        if new_level == current["level"]:
            return
        flush()
        current["level"] = new_level
        set_edit_mode(-1)
        line_edit.clear()
        fill(new_level)
        suggest_lbl.setText("")
        suggest_lbl.setStyleSheet("")

    def on_line_clicked(item) -> None:
        filling_edit["on"] = True
        set_edit_mode(lines.row(item))
        line_edit.setText(item.text())
        line_edit.setFocus()
        filling_edit["on"] = False
        suggest_lbl.setText("")
        suggest_lbl.setStyleSheet("")

    def on_line_text(text: str) -> None:
        if filling_edit["on"]:
            return
        if not str(text or "").strip():
            set_edit_mode(-1)

    def on_add() -> None:
        text = wit._clean_line(line_edit.text())
        if not text:
            return
        row = editing["row"]
        if 0 <= row < lines.count():
            for i in range(lines.count()):
                if i != row and lines.item(i).text().lower() == text.lower():
                    suggest_lbl.setText("That line is already in this band.")
                    suggest_lbl.setStyleSheet("color: #b3261e;")
                    return
            lines.item(row).setText(text)
            suggest_lbl.setText("")
            suggest_lbl.setStyleSheet("")
            return
        if any(lines.item(i).text().lower() == text.lower() for i in range(lines.count())):
            suggest_lbl.setText("That line is already in this band.")
            suggest_lbl.setStyleSheet("color: #b3261e;")
            return
        lines.addItem(text)
        line_edit.clear()
        set_edit_mode(-1)
        suggest_lbl.setText("")
        suggest_lbl.setStyleSheet("")

    def on_remove() -> None:
        set_edit_mode(-1)
        line_edit.clear()
        for item in list(lines.selectedItems()):
            lines.takeItem(lines.row(item))

    class _PoeBridge(QObject):
        suggested = Signal(bool, str)
        tested = Signal(bool, str)

    bridge = _PoeBridge(panel)

    def selected_model() -> str:
        return str(model_combo.currentData() or wit.DEFAULT_MODEL)

    def on_suggested(ok: bool, text: str) -> None:
        suggest_btn.setEnabled(True)
        if ok:
            set_edit_mode(-1)
            line_edit.setText(text)
            line_edit.setFocus()
            suggest_lbl.setText("Suggestion ready. Click Add to keep it.")
            suggest_lbl.setStyleSheet("")
        else:
            suggest_lbl.setText(text)
            suggest_lbl.setStyleSheet("color: #b3261e;")

    def on_tested(ok: bool, text: str) -> None:
        test_btn.setEnabled(True)
        suggest_lbl.setText(("\u2714 " if ok else "\u2716 ") + text)
        suggest_lbl.setStyleSheet(f"color: {'#1b7f3b' if ok else '#b3261e'};")

    bridge.suggested.connect(on_suggested)
    bridge.tested.connect(on_tested)

    def current_key() -> str:
        return key_edit.text().strip() or get_poe_api_key()

    def on_suggest() -> None:
        key = current_key()
        if not key:
            suggest_lbl.setText("Paste a Poe API key first. Create one at poe.com/api/keys.")
            suggest_lbl.setStyleSheet("color: #b3261e;")
            return
        suggest_btn.setEnabled(False)
        suggest_lbl.setText("Asking Poe…")
        suggest_lbl.setStyleSheet("color: #666;")
        level = current["level"]
        existing = read_list()
        model = selected_model()

        def work() -> None:
            try:
                line = wit.suggest_witticism(level, existing, key, model)
                bridge.suggested.emit(True, line)
            except Exception as exc:  # noqa: BLE001
                bridge.suggested.emit(False, str(exc) or "Poe request failed.")

        threading.Thread(target=work, daemon=True).start()

    def on_test() -> None:
        key = current_key()
        if not key:
            suggest_lbl.setText("\u2716 Paste a Poe API key first. Create one at poe.com/api/keys.")
            suggest_lbl.setStyleSheet("color: #b3261e;")
            return
        test_btn.setEnabled(False)
        suggest_lbl.setText("Checking Poe…")
        suggest_lbl.setStyleSheet("color: #666;")
        model = selected_model()

        def work() -> None:
            try:
                message = wit.test_poe_connection(key, model)
                bridge.tested.emit(True, message)
            except Exception as exc:  # noqa: BLE001
                bridge.tested.emit(False, str(exc) or "Poe request failed.")

        threading.Thread(target=work, daemon=True).start()

    def collect() -> Dict[str, Any]:
        flush()
        return {
            "model": selected_model(),
            "lines": {level: list(banks[level]) for level in wit.LEVELS},
        }

    def save_key() -> None:
        if clear_key.isChecked():
            set_poe_api_key("")
            return
        typed = key_edit.text().strip()
        if typed:
            set_poe_api_key(typed)

    band_combo.currentIndexChanged.connect(on_band)
    lines.itemClicked.connect(on_line_clicked)
    line_edit.textChanged.connect(on_line_text)
    add_btn.clicked.connect(on_add)
    line_edit.returnPressed.connect(on_add)
    remove_btn.clicked.connect(on_remove)
    suggest_btn.clicked.connect(on_suggest)
    test_btn.clicked.connect(on_test)
    panel.collect_witticisms = collect  # type: ignore[attr-defined]
    panel.save_poe_key = save_key  # type: ignore[attr-defined]
    return panel
