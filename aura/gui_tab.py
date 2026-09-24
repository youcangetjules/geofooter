"""GURI Data Brokers tab — identity, catalog, requests, AES pending queue."""

from __future__ import annotations

import webbrowser
from pathlib import Path
from typing import Any, Callable, Dict, List, Optional

from PySide6.QtCore import Qt
from PySide6.QtGui import QGuiApplication
from PySide6.QtWidgets import (
    QComboBox,
    QFormLayout,
    QGroupBox,
    QHBoxLayout,
    QInputDialog,
    QLabel,
    QLineEdit,
    QMessageBox,
    QPlainTextEdit,
    QPushButton,
    QSplitter,
    QTabWidget,
    QTreeWidget,
    QTreeWidgetItem,
    QVBoxLayout,
    QWidget,
    QFileDialog,
)

from aura.catalog import find_broker, load_catalog
from aura.pending import accept_pending, dismiss_pending, list_pending
from aura.profile import IdentityProfile, load_profile, save_profile
from aura.store import STATUSES, RemovalStore
from aura.templates import render_removal


class DataBrokersPanel(QWidget):
    """Incogni-style data-broker removal workspace."""

    def __init__(
        self,
        parent: Optional[QWidget] = None,
        *,
        status_fn: Optional[Callable[[str], None]] = None,
        accent: str = "#0f6b7c",
        muted: str = "#5a7280",
    ) -> None:
        super().__init__(parent)
        self._status = status_fn or (lambda _m: None)
        self._accent = accent
        self._muted = muted
        self._store = RemovalStore()
        self._build()
        self.reload_all()

    def _msg(self, text: str) -> None:
        self._status(text)

    def _build(self) -> None:
        root = QVBoxLayout(self)
        root.setContentsMargins(12, 12, 12, 12)
        root.setSpacing(8)

        head = QLabel("Aura — Aliniant Universal Removal Application")
        head.setStyleSheet(f"font-size: 14px; font-weight: 700; color: {self._accent};")
        root.addWidget(head)

        sub = QLabel(
            "Track opt-out / erasure requests against people-search and data brokers "
            "(Incogni-style). Templates are filled from your local identity profile — "
            "you copy or export and send manually. Nothing is submitted automatically."
        )
        sub.setWordWrap(True)
        sub.setStyleSheet(f"color: {self._muted};")
        root.addWidget(sub)

        warn = QLabel(
            "Identity profile is stored in plaintext under %LOCALAPPDATA%\\GeoFooter\\. "
            "Use only data you are prepared to put in removal letters."
        )
        warn.setWordWrap(True)
        warn.setStyleSheet(f"color: #b54708; font-size: 11px;")
        root.addWidget(warn)

        tabs = QTabWidget()
        root.addWidget(tabs, stretch=1)

        tabs.addTab(self._build_identity_tab(), "Identity")
        tabs.addTab(self._build_brokers_tab(), "Brokers")
        tabs.addTab(self._build_requests_tab(), "Requests")
        tabs.addTab(self._build_aes_pending_tab(), "AES inbox")

    # ---- Identity ----

    def _build_identity_tab(self) -> QWidget:
        w = QWidget()
        layout = QVBoxLayout(w)
        form = QFormLayout()
        self.id_name = QLineEdit()
        self.id_aliases = QPlainTextEdit()
        self.id_aliases.setPlaceholderText("One alias per line")
        self.id_aliases.setMaximumHeight(60)
        self.id_emails = QPlainTextEdit()
        self.id_emails.setPlaceholderText("One email per line")
        self.id_emails.setMaximumHeight(60)
        self.id_phones = QPlainTextEdit()
        self.id_phones.setPlaceholderText("One phone per line")
        self.id_phones.setMaximumHeight(50)
        self.id_addresses = QPlainTextEdit()
        self.id_addresses.setPlaceholderText("One address per line")
        self.id_addresses.setMaximumHeight(80)
        self.id_dob = QLineEdit()
        self.id_dob.setPlaceholderText("YYYY-MM-DD (optional)")
        self.id_notes = QPlainTextEdit()
        self.id_notes.setMaximumHeight(50)
        form.addRow("Full name", self.id_name)
        form.addRow("Aliases", self.id_aliases)
        form.addRow("Emails", self.id_emails)
        form.addRow("Phones", self.id_phones)
        form.addRow("Addresses", self.id_addresses)
        form.addRow("Date of birth", self.id_dob)
        form.addRow("Notes", self.id_notes)
        layout.addLayout(form)
        row = QHBoxLayout()
        save_btn = QPushButton("Save profile")
        save_btn.clicked.connect(self._save_identity)
        reload_btn = QPushButton("Reload")
        reload_btn.clicked.connect(self._load_identity)
        row.addWidget(save_btn)
        row.addWidget(reload_btn)
        row.addStretch(1)
        layout.addLayout(row)
        layout.addStretch(1)
        return w

    def _lines(self, edit: QPlainTextEdit) -> List[str]:
        return [ln.strip() for ln in edit.toPlainText().splitlines() if ln.strip()]

    def _set_lines(self, edit: QPlainTextEdit, values: List[str]) -> None:
        edit.setPlainText("\n".join(values))

    def _load_identity(self) -> None:
        p = load_profile()
        self.id_name.setText(p.full_name)
        self._set_lines(self.id_aliases, p.aliases)
        self._set_lines(self.id_emails, p.emails)
        self._set_lines(self.id_phones, p.phones)
        self._set_lines(self.id_addresses, p.addresses)
        self.id_dob.setText(p.date_of_birth)
        self.id_notes.setPlainText(p.notes)

    def _save_identity(self) -> None:
        p = IdentityProfile(
            full_name=self.id_name.text().strip(),
            aliases=self._lines(self.id_aliases),
            emails=self._lines(self.id_emails),
            phones=self._lines(self.id_phones),
            addresses=self._lines(self.id_addresses),
            date_of_birth=self.id_dob.text().strip(),
            notes=self.id_notes.toPlainText().strip(),
        )
        path = save_profile(p)
        self._msg(f"Identity profile saved → {path}")

    # ---- Brokers catalog ----

    def _build_brokers_tab(self) -> QWidget:
        w = QWidget()
        layout = QVBoxLayout(w)
        filt_row = QHBoxLayout()
        self.broker_filter = QLineEdit()
        self.broker_filter.setPlaceholderText("Filter brokers…")
        self.broker_filter.setClearButtonEnabled(True)
        self.broker_filter.textChanged.connect(self._reload_brokers)
        filt_row.addWidget(self.broker_filter, stretch=1)
        refresh = QPushButton("Refresh")
        refresh.clicked.connect(self._reload_brokers)
        filt_row.addWidget(refresh)
        layout.addLayout(filt_row)

        self.broker_tree = QTreeWidget()
        self.broker_tree.setHeaderLabels(
            ("Name", "Domains", "Method", "Jurisdiction", "Opt-out")
        )
        self.broker_tree.setRootIsDecorated(False)
        self.broker_tree.setSortingEnabled(True)
        self.broker_tree.setColumnWidth(0, 180)
        self.broker_tree.setColumnWidth(1, 220)
        layout.addWidget(self.broker_tree, stretch=1)

        actions = QHBoxLayout()
        create_btn = QPushButton("Create request")
        create_btn.clicked.connect(self._broker_create_request)
        open_btn = QPushButton("Open opt-out URL")
        open_btn.clicked.connect(self._broker_open_url)
        copy_email = QPushButton("Copy opt-out email")
        copy_email.clicked.connect(self._broker_copy_email)
        preview_btn = QPushButton("Preview template…")
        preview_btn.clicked.connect(self._broker_preview_template)
        actions.addWidget(create_btn)
        actions.addWidget(open_btn)
        actions.addWidget(copy_email)
        actions.addWidget(preview_btn)
        actions.addStretch(1)
        layout.addLayout(actions)
        return w

    def _selected_broker_id(self) -> str:
        item = self.broker_tree.currentItem()
        return str(item.data(0, Qt.ItemDataRole.UserRole) or "") if item else ""

    def _reload_brokers(self) -> None:
        needle = self.broker_filter.text().strip().lower()
        self.broker_tree.clear()
        for b in load_catalog():
            blob = " ".join(
                [
                    b["name"],
                    " ".join(b.get("domains") or []),
                    b.get("method") or "",
                    b.get("jurisdiction") or "",
                ]
            ).lower()
            if needle and needle not in blob:
                continue
            opt = b.get("opt_out_email") or b.get("opt_out_url") or "—"
            item = QTreeWidgetItem(
                [
                    b["name"],
                    ", ".join(b.get("domains") or []),
                    b.get("method") or "",
                    b.get("jurisdiction") or "",
                    opt[:60],
                ]
            )
            item.setData(0, Qt.ItemDataRole.UserRole, b["id"])
            self.broker_tree.addTopLevelItem(item)

    def _broker_create_request(self) -> None:
        bid = self._selected_broker_id()
        b = find_broker(bid) if bid else None
        if not b:
            QMessageBox.information(self, "Brokers", "Select a broker first.")
            return
        req = self._store.create(
            broker_id=b["id"],
            broker_name=b["name"],
            status="draft",
            source="manual",
            domain=(b.get("domains") or [""])[0],
        )
        self._reload_requests()
        self._msg(f"Created draft request for {b['name']} ({req.id})")

    def _broker_open_url(self) -> None:
        bid = self._selected_broker_id()
        b = find_broker(bid) if bid else None
        url = (b or {}).get("opt_out_url") or ""
        if not url:
            QMessageBox.information(self, "Brokers", "No opt-out URL for this broker.")
            return
        webbrowser.open(url)
        self._msg(f"Opened {url}")

    def _broker_copy_email(self) -> None:
        bid = self._selected_broker_id()
        b = find_broker(bid) if bid else None
        email = (b or {}).get("opt_out_email") or ""
        if not email:
            QMessageBox.information(self, "Brokers", "No opt-out email for this broker.")
            return
        QGuiApplication.clipboard().setText(email)
        self._msg(f"Copied {email}")

    def _broker_preview_template(self) -> None:
        bid = self._selected_broker_id()
        b = find_broker(bid) if bid else None
        if not b:
            QMessageBox.information(self, "Brokers", "Select a broker first.")
            return
        self._show_template(b)

    def _show_template(self, broker: Dict[str, Any]) -> None:
        label, subject, body = render_removal(broker, load_profile())
        dlg = QMessageBox(self)
        dlg.setWindowTitle(f"Template — {broker.get('name')}")
        dlg.setText(f"{label}\n\nSubject: {subject}")
        dlg.setDetailedText(body)
        copy = dlg.addButton("Copy body", QMessageBox.ButtonRole.ActionRole)
        export = dlg.addButton("Export…", QMessageBox.ButtonRole.ActionRole)
        dlg.addButton(QMessageBox.StandardButton.Close)
        dlg.exec()
        clicked = dlg.clickedButton()
        if clicked == copy:
            QGuiApplication.clipboard().setText(f"Subject: {subject}\n\n{body}")
            self._msg("Template copied to clipboard")
        elif clicked == export:
            path, _ = QFileDialog.getSaveFileName(
                self,
                "Export removal letter",
                f"removal_{broker.get('id') or 'broker'}.txt",
                "Text (*.txt)",
            )
            if path:
                Path(path).write_text(
                    f"Subject: {subject}\n\n{body}", encoding="utf-8"
                )
                self._msg(f"Exported {path}")

    # ---- Requests ----

    def _build_requests_tab(self) -> QWidget:
        w = QWidget()
        layout = QVBoxLayout(w)
        filt = QHBoxLayout()
        self.req_status_filter = QComboBox()
        self.req_status_filter.addItem("All statuses", "")
        for s in STATUSES:
            self.req_status_filter.addItem(s, s)
        self.req_status_filter.currentIndexChanged.connect(self._reload_requests)
        filt.addWidget(QLabel("Status"))
        filt.addWidget(self.req_status_filter)
        refresh = QPushButton("Refresh")
        refresh.clicked.connect(self._reload_requests)
        filt.addWidget(refresh)
        filt.addStretch(1)
        layout.addLayout(filt)

        self.req_tree = QTreeWidget()
        self.req_tree.setHeaderLabels(
            ("Broker", "Status", "Source", "Updated", "Follow-up", "Sender")
        )
        self.req_tree.setRootIsDecorated(False)
        self.req_tree.setSortingEnabled(True)
        self.req_tree.setColumnWidth(0, 160)
        layout.addWidget(self.req_tree, stretch=1)

        actions = QHBoxLayout()
        for label, slot in (
            ("Template…", self._req_template),
            ("Mark status…", self._req_mark_status),
            ("Set follow-up…", self._req_follow_up),
            ("Edit notes…", self._req_notes),
            ("Delete…", self._req_delete),
        ):
            btn = QPushButton(label)
            btn.clicked.connect(slot)
            actions.addWidget(btn)
        actions.addStretch(1)
        layout.addLayout(actions)
        return w

    def _selected_request_id(self) -> str:
        item = self.req_tree.currentItem()
        return str(item.data(0, Qt.ItemDataRole.UserRole) or "") if item else ""

    def _reload_requests(self) -> None:
        status = self.req_status_filter.currentData()
        self.req_tree.clear()
        for req in self._store.list_requests(status=status or None):
            item = QTreeWidgetItem(
                [
                    req.broker_name,
                    req.status,
                    req.source,
                    req.updated_at[:16].replace("T", " "),
                    req.follow_up_due or "—",
                    req.sender or req.domain or "—",
                ]
            )
            item.setData(0, Qt.ItemDataRole.UserRole, req.id)
            self.req_tree.addTopLevelItem(item)

    def _req_template(self) -> None:
        rid = self._selected_request_id()
        req = self._store.get(rid) if rid else None
        if not req:
            QMessageBox.information(self, "Requests", "Select a request.")
            return
        b = find_broker(req.broker_id) or {
            "id": req.broker_id,
            "name": req.broker_name,
            "template_id": "generic",
        }
        self._show_template(b)

    def _req_mark_status(self) -> None:
        rid = self._selected_request_id()
        if not rid:
            QMessageBox.information(self, "Requests", "Select a request.")
            return
        status, ok = QInputDialog.getItem(
            self, "Mark status", "New status:", list(STATUSES), 0, False
        )
        if not ok:
            return
        self._store.update(rid, status=status, last_action=f"status→{status}")
        self._reload_requests()
        self._msg(f"Request {rid} → {status}")

    def _req_follow_up(self) -> None:
        rid = self._selected_request_id()
        if not rid:
            return
        due, ok = QInputDialog.getText(
            self, "Follow-up", "Follow-up date (YYYY-MM-DD):"
        )
        if not ok:
            return
        self._store.update(
            rid, follow_up_due=due.strip(), last_action="follow_up set", status="follow_up"
        )
        self._reload_requests()

    def _req_notes(self) -> None:
        rid = self._selected_request_id()
        req = self._store.get(rid) if rid else None
        if not req:
            return
        notes, ok = QInputDialog.getMultiLineText(
            self, "Notes", "Request notes:", req.notes
        )
        if not ok:
            return
        self._store.update(rid, notes=notes, last_action="notes edited")
        self._reload_requests()

    def _req_delete(self) -> None:
        rid = self._selected_request_id()
        if not rid:
            return
        if (
            QMessageBox.question(self, "Delete", f"Delete request {rid}?")
            != QMessageBox.StandardButton.Yes
        ):
            return
        self._store.delete(rid)
        self._reload_requests()
        self._msg(f"Deleted {rid}")

    # ---- AES pending ----

    def _build_aes_pending_tab(self) -> QWidget:
        w = QWidget()
        layout = QVBoxLayout(w)
        hint = QLabel(
            "Items queued from an AES footer (Queue data removal) on broker mail. "
            "Accept creates a draft request; Dismiss drops the item."
        )
        hint.setWordWrap(True)
        hint.setStyleSheet(f"color: {self._muted};")
        layout.addWidget(hint)
        self.pending_tree = QTreeWidget()
        self.pending_tree.setHeaderLabels(
            ("Broker", "Sender", "Subject", "Queued")
        )
        self.pending_tree.setRootIsDecorated(False)
        layout.addWidget(self.pending_tree, stretch=1)
        row = QHBoxLayout()
        accept = QPushButton("Accept → draft request")
        accept.clicked.connect(self._pending_accept)
        dismiss = QPushButton("Dismiss")
        dismiss.clicked.connect(self._pending_dismiss)
        refresh = QPushButton("Refresh")
        refresh.clicked.connect(self._reload_pending)
        row.addWidget(accept)
        row.addWidget(dismiss)
        row.addWidget(refresh)
        row.addStretch(1)
        layout.addLayout(row)
        return w

    def _selected_pending_id(self) -> str:
        item = self.pending_tree.currentItem()
        return str(item.data(0, Qt.ItemDataRole.UserRole) or "") if item else ""

    def _reload_pending(self) -> None:
        self.pending_tree.clear()
        for item in list_pending():
            row = QTreeWidgetItem(
                [
                    str(item.get("broker_name") or ""),
                    str(item.get("sender") or item.get("domain") or ""),
                    str(item.get("subject") or "")[:80],
                    str(item.get("scanned_at") or "")[:16].replace("T", " "),
                ]
            )
            row.setData(0, Qt.ItemDataRole.UserRole, item.get("id"))
            self.pending_tree.addTopLevelItem(row)

    def _pending_accept(self) -> None:
        pid = self._selected_pending_id()
        if not pid:
            QMessageBox.information(self, "AES inbox", "Select a pending item.")
            return
        req = accept_pending(pid, self._store)
        if not req:
            QMessageBox.warning(self, "AES inbox", "Item not found.")
            return
        self._reload_pending()
        self._reload_requests()
        self._msg(f"Accepted → draft {req.id} ({req.broker_name})")

    def _pending_dismiss(self) -> None:
        pid = self._selected_pending_id()
        if not pid:
            return
        dismiss_pending(pid)
        self._reload_pending()
        self._msg("Dismissed pending AES item")

    def reload_all(self) -> None:
        self._load_identity()
        self._reload_brokers()
        self._reload_requests()
        self._reload_pending()
