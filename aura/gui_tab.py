"""GURI Data Brokers tab — identity, catalog, requests, AES pending queue."""

from __future__ import annotations

import threading
import webbrowser
from pathlib import Path
from typing import Any, Callable, Dict, List, Optional

from PySide6.QtCore import QObject, Qt, QTimer, Signal
from PySide6.QtGui import QGuiApplication
from PySide6.QtWidgets import (
    QComboBox,
    QFormLayout,
    QHBoxLayout,
    QInputDialog,
    QLabel,
    QLineEdit,
    QMessageBox,
    QPlainTextEdit,
    QPushButton,
    QTabWidget,
    QTreeWidget,
    QTreeWidgetItem,
    QVBoxLayout,
    QWidget,
    QFileDialog,
)

from aura import automation
from aura.catalog import find_broker, load_catalog
from aura.pending import accept_pending, dismiss_pending, list_pending
from aura.profile import IdentityProfile, load_profile, save_profile
from aura.settings import load_settings, save_settings
from aura.store import STATUSES, RemovalStore
from aura.templates import render_removal

AUTO_RUN_MS = 5 * 60 * 1000


class _ComJob(QObject):
    """Runs one Outlook COM call on a worker thread; result comes back on the GUI thread."""

    finished = Signal(object, str)  # (result, error text)

    def __init__(self, fn: Callable[[], Any], parent: Optional[QObject] = None) -> None:
        super().__init__(parent)
        self._fn = fn

    def start(self) -> None:
        threading.Thread(target=self._run, daemon=True).start()

    def _run(self) -> None:
        try:
            with automation._com_apartment():
                result = self._fn()
            self.finished.emit(result, "")
        except Exception as exc:
            self.finished.emit(None, str(exc) or exc.__class__.__name__)


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
        self._jobs: List[_ComJob] = []
        self._auto_busy = False
        self._auto_started = False
        self._build()
        self.reload_all()
        self._auto_timer = QTimer(self)
        self._auto_timer.setInterval(AUTO_RUN_MS)
        self._auto_timer.timeout.connect(self._auto_run)

    def showEvent(self, event) -> None:  # noqa: N802 (Qt override)
        super().showEvent(event)
        if not self._auto_started:
            self._auto_started = True
            self._auto_timer.start()
            QTimer.singleShot(500, self._auto_run)
            QTimer.singleShot(0, self._load_send_accounts)

    def _run_com(self, fn: Callable[[], Any], done: Callable[[Any, str], None]) -> None:
        job = _ComJob(fn, self)
        self._jobs.append(job)

        def _finish(result: Any, err: str) -> None:
            self._jobs.remove(job)
            job.deleteLater()
            done(result, err)

        job.finished.connect(_finish)
        job.start()

    def _auto_run(self) -> None:
        if self._auto_busy:
            return
        self._auto_busy = True
        self._run_com(automation.prepare_and_sync, self._auto_done)

    def _auto_done(self, summary: Any, err: str) -> None:
        self._auto_busy = False
        self._msg(f"Aura: {err}" if err else str(summary))
        self._reload_requests()
        self._reload_review()
        self._reload_pending()

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
            "Aura prepares removal emails as Outlook drafts. You review and press Send. "
            "Letters are filled from your identity profile; brokers that only accept "
            "a web form are tracked in Requests with their opt-out URL."
        )
        sub.setWordWrap(True)
        sub.setStyleSheet(f"color: {self._muted};")
        root.addWidget(sub)

        warn = QLabel(
            "Identity profile is stored in plaintext under %LOCALAPPDATA%\\GeoFooter\\. "
            "Use only data you are prepared to put in removal letters."
        )
        warn.setWordWrap(True)
        warn.setStyleSheet("color: #b54708; font-size: 11px;")
        root.addWidget(warn)

        tabs = QTabWidget()
        root.addWidget(tabs, stretch=1)

        tabs.addTab(self._build_review_tab(), "Review")
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
        self.id_send_from = QComboBox()
        self.id_send_from.setEditable(True)
        self.id_send_from.setToolTip("Outlook account the removal drafts are sent from")
        form.addRow("Send from", self.id_send_from)
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
        settings = load_settings()
        settings.send_account = self.id_send_from.currentText().strip().lower()
        save_settings(settings)
        self._msg(f"Identity profile saved → {path}")

    def _load_send_accounts(self) -> None:
        def _done(accounts: Any, err: str) -> None:
            current = load_settings().send_account
            self.id_send_from.clear()
            for addr in accounts or []:
                self.id_send_from.addItem(addr)
            if not current:
                emails = {e.lower() for e in load_profile().emails}
                current = next((a for a in accounts or [] if a.lower() in emails), "")
            if current:
                self.id_send_from.setCurrentText(current)
            if err:
                self._msg(f"Aura: could not list Outlook accounts ({err})")

        self._run_com(lambda: automation.OutlookDrafts().accounts(), _done)

    # ---- Review (drafts awaiting the user's Send) ----

    def _build_review_tab(self) -> QWidget:
        w = QWidget()
        layout = QVBoxLayout(w)
        hint = QLabel(
            "Drafts Aura has prepared in Outlook (category \"Aura\"). Open a draft, "
            "check it, and press Send in Outlook; Aura marks it sent automatically. "
            "Aura re-checks every 5 minutes while GURI is open."
        )
        hint.setWordWrap(True)
        hint.setStyleSheet(f"color: {self._muted};")
        layout.addWidget(hint)
        self.review_tree = QTreeWidget()
        self.review_tree.setHeaderLabels(("Broker", "To", "Subject", "Prepared"))
        self.review_tree.setRootIsDecorated(False)
        self.review_tree.setSortingEnabled(True)
        self.review_tree.setColumnWidth(0, 160)
        self.review_tree.setColumnWidth(1, 200)
        self.review_tree.setColumnWidth(2, 320)
        self.review_tree.itemDoubleClicked.connect(lambda *_: self._review_open())
        layout.addWidget(self.review_tree, stretch=1)
        row = QHBoxLayout()
        for label, slot in (
            ("Open draft", self._review_open),
            ("Regenerate", self._review_regenerate),
            ("Skip", self._review_skip),
            ("Prepare all now", self._auto_run),
        ):
            btn = QPushButton(label)
            btn.clicked.connect(slot)
            row.addWidget(btn)
        row.addStretch(1)
        layout.addLayout(row)
        return w

    def _selected_review_id(self) -> str:
        item = self.review_tree.currentItem()
        return str(item.data(0, Qt.ItemDataRole.UserRole) or "") if item else ""

    def _reload_review(self) -> None:
        self.review_tree.clear()
        for req in self._store.list_requests(status="ready"):
            broker = find_broker(req.broker_id) or {}
            item = QTreeWidgetItem(
                [
                    req.broker_name,
                    str(broker.get("opt_out_email") or ""),
                    req.subject,
                    req.updated_at[:16].replace("T", " "),
                ]
            )
            item.setData(0, Qt.ItemDataRole.UserRole, req.id)
            self.review_tree.addTopLevelItem(item)

    def _review_request(self):
        rid = self._selected_review_id()
        req = self._store.get(rid) if rid else None
        if not req:
            QMessageBox.information(self, "Review", "Select a draft.")
        return req

    def _review_open(self) -> None:
        req = self._review_request()
        if not req:
            return

        def _done(_r: Any, err: str) -> None:
            if err:
                QMessageBox.warning(self, "Review", f"Could not open the draft:\n{err}")

        self._run_com(lambda: automation.open_draft(req), _done)

    def _review_regenerate(self) -> None:
        req = self._review_request()
        if not req:
            return

        def _done(_r: Any, err: str) -> None:
            if err:
                QMessageBox.warning(self, "Review", f"Could not regenerate:\n{err}")
            else:
                self._msg(f"Aura: regenerated draft for {req.broker_name}")
            self._reload_review()
            self._reload_requests()

        self._run_com(lambda: automation.regenerate(req.id, store=self._store), _done)

    def _review_skip(self) -> None:
        req = self._review_request()
        if not req:
            return

        def _done(_r: Any, err: str) -> None:
            if err:
                QMessageBox.warning(self, "Review", f"Could not skip:\n{err}")
            else:
                self._msg(f"Aura: skipped {req.broker_name} (draft deleted)")
            self._reload_review()
            self._reload_requests()

        self._run_com(lambda: automation.skip(req.id, store=self._store), _done)

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
            "Broker mail AES detected (dated on or after the Aura start date) or queued "
            "from a footer. Aura turns these into Outlook drafts on its next run; "
            "Accept / Dismiss handle them by hand."
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
        self._reload_review()
        self._reload_pending()
