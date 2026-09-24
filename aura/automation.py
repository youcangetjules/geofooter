#!/usr/bin/env python3
"""Aura automation: prepare removal emails as Outlook drafts for user review.

Nothing here ever sends mail. ``prepare()`` renders each letter and saves it
to Outlook Drafts (category "Aura"); the user opens the draft and presses Send.
``sync_sent()`` later notices the send and moves the request to ``sent``.

CLI (run by file path):
    python aura\\automation.py --prepare
    python aura\\automation.py --sync
"""

from __future__ import annotations

import sys
from pathlib import Path

if __name__ == "__main__":
    _here = Path(__file__).resolve().parent
    _root = _here.parent
    if str(_here) in sys.path:
        sys.path.remove(str(_here))
    if str(_root) not in sys.path:
        sys.path.insert(0, str(_root))

import argparse
import logging
from contextlib import contextmanager
from dataclasses import dataclass, field
from datetime import datetime, timedelta, timezone
from typing import Any, Dict, Iterator, List, Optional, Protocol

from aura.catalog import find_broker, load_catalog
from aura.pending import accept_pending, dismiss_pending, list_pending
from aura.profile import IdentityProfile, load_profile
from aura.settings import AuraSettings, load_settings
from aura.store import ACTIVE_STATUSES, RemovalRequest, RemovalStore
from aura.templates import render_removal

log = logging.getLogger("aura.automation")

DRAFT_CATEGORY = "Aura"
OL_MAIL_ITEM = 0
OL_FOLDER_SENT = 5
OL_FOLDER_DRAFTS = 16
# DISPID of MailItem.SendUsingAccount; pywin32 can't assign object props by name.
_DISPID_SEND_USING_ACCOUNT = 64209
# last_action values meaning the user declined this draft; the timer must not recreate it.
ACTION_SKIPPED = "skipped by user"
ACTION_DRAFT_DELETED = "draft deleted"
USER_PARKED = (ACTION_SKIPPED, ACTION_DRAFT_DELETED)


class DraftBackend(Protocol):
    def create(self, to: str, subject: str, body: str, send_account: str) -> str: ...
    def in_drafts(self, entry_id: str) -> bool: ...
    def find_sent(self, to: str, subject: str, since: datetime) -> Optional[datetime]: ...
    def display(self, entry_id: str) -> None: ...
    def delete(self, entry_id: str) -> None: ...
    def accounts(self) -> List[str]: ...


@dataclass
class PrepareReport:
    created: List[str] = field(default_factory=list)
    skipped: List[str] = field(default_factory=list)
    tracked: List[str] = field(default_factory=list)
    errors: List[str] = field(default_factory=list)
    blocked: str = ""

    def summary(self) -> str:
        if self.blocked:
            return f"Aura: {self.blocked}"
        parts = [f"{len(self.created)} draft(s) ready for review"]
        if self.tracked:
            parts.append(f"{len(self.tracked)} web-form broker(s) tracked")
        if self.errors:
            parts.append(f"{len(self.errors)} error(s)")
        return "Aura: " + ", ".join(parts)


@dataclass
class SyncReport:
    sent: List[str] = field(default_factory=list)
    reverted: List[str] = field(default_factory=list)
    errors: List[str] = field(default_factory=list)

    def summary(self) -> str:
        parts = [f"{len(self.sent)} sent"]
        if self.reverted:
            parts.append(f"{len(self.reverted)} draft(s) deleted")
        if self.errors:
            parts.append(f"{len(self.errors)} error(s)")
        return "Aura sync: " + ", ".join(parts)


# ---------------------------------------------------------------------------
# Outlook COM backend
# ---------------------------------------------------------------------------


@contextmanager
def _com_apartment() -> Iterator[None]:
    import pythoncom  # type: ignore

    pythoncom.CoInitialize()
    try:
        yield
    finally:
        pythoncom.CoUninitialize()


def _naive_local(value: Any) -> Optional[datetime]:
    """pywin32 COM dates come back tz-labelled but are really local wall time."""
    if value is None:
        return None
    try:
        return datetime(*value.timetuple()[:6])
    except Exception:
        return None


class OutlookDrafts:
    """Outlook COM draft operations. Call from a thread inside ``_com_apartment``."""

    def __init__(self) -> None:
        import win32com.client  # type: ignore

        try:
            self._app = win32com.client.GetObject(Class="Outlook.Application")
        except Exception:
            self._app = win32com.client.Dispatch("Outlook.Application")
        self._ns = self._app.GetNamespace("MAPI")

    def _account(self, smtp: str) -> Any:
        smtp = (smtp or "").strip().lower()
        if not smtp:
            return None
        for acc in self._ns.Accounts:
            try:
                if str(acc.SmtpAddress or "").strip().lower() == smtp:
                    return acc
            except Exception:
                continue
        return None

    def accounts(self) -> List[str]:
        out: List[str] = []
        for acc in self._ns.Accounts:
            try:
                addr = str(acc.SmtpAddress or "").strip()
            except Exception:
                addr = ""
            if addr:
                out.append(addr)
        return out

    def create(self, to: str, subject: str, body: str, send_account: str) -> str:
        mail = self._app.CreateItem(OL_MAIL_ITEM)
        mail.To = to
        mail.Subject = subject
        mail.Body = body
        mail.Categories = DRAFT_CATEGORY
        acc = self._account(send_account)
        if acc is not None:
            try:
                mail.SendUsingAccount = acc
            except Exception:
                mail._oleobj_.Invoke(_DISPID_SEND_USING_ACCOUNT, 0, 8, 0, acc)
        mail.Save()
        return str(mail.EntryID)

    def _item(self, entry_id: str) -> Any:
        if not entry_id:
            return None
        try:
            return self._ns.GetItemFromID(entry_id)
        except Exception:
            return None

    def in_drafts(self, entry_id: str) -> bool:
        item = self._item(entry_id)
        if item is None:
            return False
        try:
            if bool(item.Sent):
                return False
            parent = item.Parent
            drafts = parent.Store.GetDefaultFolder(OL_FOLDER_DRAFTS)
            return str(parent.EntryID) == str(drafts.EntryID)
        except Exception:
            return False

    def find_sent(self, to: str, subject: str, since: datetime) -> Optional[datetime]:
        to_l = (to or "").strip().lower()
        esc = subject.replace("'", "''")
        query = f"@SQL=\"urn:schemas:httpmail:subject\" = '{esc}'"
        since_local = since.astimezone().replace(tzinfo=None) - timedelta(minutes=5)
        for store in self._ns.Stores:
            try:
                folder = store.GetDefaultFolder(OL_FOLDER_SENT)
                items = folder.Items.Restrict(query)
            except Exception:
                continue
            for item in items:
                try:
                    sent_on = _naive_local(item.SentOn)
                    recipients = " ".join(
                        str(r.Address or "") + " " + str(r.Name or "")
                        for r in item.Recipients
                    ).lower()
                except Exception:
                    continue
                if sent_on and sent_on >= since_local and to_l in recipients:
                    return sent_on
        return None

    def display(self, entry_id: str) -> None:
        item = self._item(entry_id)
        if item is None:
            raise LookupError("Draft not found in Outlook (was it deleted or sent?)")
        item.Display(False)

    def delete(self, entry_id: str) -> None:
        item = self._item(entry_id)
        if item is not None:
            item.Delete()


# ---------------------------------------------------------------------------
# Core logic (backend-injected so it runs without Outlook in tests)
# ---------------------------------------------------------------------------


def _parse_iso(value: str) -> datetime:
    try:
        return datetime.fromisoformat((value or "").replace("Z", "+00:00"))
    except Exception:
        return datetime.now(timezone.utc)


def email_brokers() -> List[Dict[str, Any]]:
    return [b for b in load_catalog() if b.get("opt_out_email")]


def profile_gap(profile: IdentityProfile) -> str:
    if not profile.full_name:
        return "fill in Full name on the Identity tab before preparing drafts."
    if not profile.emails:
        return "add at least one email on the Identity tab before preparing drafts."
    return ""


def _collect_targets() -> Dict[str, Dict[str, Any]]:
    """Catalog email brokers + AES pending items, keyed by broker_id.

    AES pending items for brokers without an opt-out email become tracking
    requests (the user handles those by web form).
    """
    targets: Dict[str, Dict[str, Any]] = {
        b["id"]: {"broker": b, "source": "catalog", "pending": None} for b in email_brokers()
    }
    for item in list_pending():
        broker = find_broker(str(item.get("broker_id") or ""))
        if broker and broker.get("opt_out_email"):
            targets[broker["id"]] = {"broker": broker, "source": "aes", "pending": item}
        else:
            accept_pending(str(item.get("id") or ""))
    return targets


def _draft_one(
    store: RemovalStore,
    backend: DraftBackend,
    broker: Dict[str, Any],
    source: str,
    pending: Optional[Dict[str, Any]],
    profile: IdentityProfile,
    settings: AuraSettings,
    existing: Optional[RemovalRequest],
) -> RemovalRequest:
    _label, subject, body = render_removal(broker, profile)
    entry_id = backend.create(broker["opt_out_email"], subject, body, settings.send_account)
    note = f"Outlook draft prepared {datetime.now().strftime('%Y-%m-%d %H:%M')} - review and Send."
    if existing is not None:
        req = store.update(
            existing.id,
            status="ready",
            subject=subject,
            source=source,
            draft_entry_id=entry_id,
            last_action="draft prepared",
            notes=(existing.notes + "\n" + note).strip(),
        )
        assert req is not None
    else:
        req = store.create(
            broker_id=broker["id"],
            broker_name=broker["name"],
            status="ready",
            source=source,
            sender=str((pending or {}).get("sender") or ""),
            domain=str((pending or {}).get("domain") or ""),
            subject=subject,
            guri=str((pending or {}).get("guri") or ""),
            notes=note,
            draft_entry_id=entry_id,
        )
    if pending is not None:
        dismiss_pending(str(pending.get("id") or ""))
    return req


def prepare(
    store: Optional[RemovalStore] = None,
    profile: Optional[IdentityProfile] = None,
    settings: Optional[AuraSettings] = None,
    backend: Optional[DraftBackend] = None,
) -> PrepareReport:
    store = store or RemovalStore()
    profile = profile or load_profile()
    settings = settings or load_settings()
    report = PrepareReport()

    gap = profile_gap(profile)
    if gap:
        report.blocked = gap
        return report

    if backend is None:
        backend = OutlookDrafts()

    for broker_id, target in _collect_targets().items():
        broker = target["broker"]
        existing = store.latest_for_broker(broker_id)
        if existing is not None and (
            existing.status in ACTIVE_STATUSES or existing.last_action in USER_PARKED
        ):
            report.skipped.append(broker["name"])
            if target["pending"] is not None:
                dismiss_pending(str(target["pending"].get("id") or ""))
            continue
        try:
            _draft_one(
                store, backend, broker, target["source"], target["pending"],
                profile, settings, existing,
            )
            report.created.append(broker["name"])
        except Exception as exc:
            log.exception("Draft failed for %s", broker["name"])
            report.errors.append(f"{broker['name']}: {exc}")

    for broker in load_catalog():
        if broker.get("opt_out_email") or store.latest_for_broker(broker["id"]):
            continue
        url = broker.get("opt_out_url") or "(no URL in catalog)"
        store.create(
            broker_id=broker["id"],
            broker_name=broker["name"],
            status="draft",
            source="catalog",
            notes=f"Manual web form (not automated): {url}",
        )
        report.tracked.append(broker["name"])

    return report


def regenerate(
    request_id: str,
    store: Optional[RemovalStore] = None,
    profile: Optional[IdentityProfile] = None,
    settings: Optional[AuraSettings] = None,
    backend: Optional[DraftBackend] = None,
) -> RemovalRequest:
    store = store or RemovalStore()
    profile = profile or load_profile()
    settings = settings or load_settings()
    gap = profile_gap(profile)
    if gap:
        raise ValueError(gap)
    req = store.get(request_id)
    if req is None:
        raise LookupError("Request not found")
    broker = find_broker(req.broker_id)
    if not broker or not broker.get("opt_out_email"):
        raise ValueError(f"{req.broker_name} has no opt-out email; it is handled by web form.")
    backend = backend or OutlookDrafts()
    if req.draft_entry_id:
        backend.delete(req.draft_entry_id)
    return _draft_one(store, backend, broker, req.source, None, profile, settings, req)


def skip(request_id: str, store: Optional[RemovalStore] = None,
         backend: Optional[DraftBackend] = None) -> Optional[RemovalRequest]:
    """Drop the Outlook draft and park the request back in ``draft``."""
    store = store or RemovalStore()
    req = store.get(request_id)
    if req is None:
        return None
    if req.draft_entry_id:
        (backend or OutlookDrafts()).delete(req.draft_entry_id)
    return store.update(
        request_id, status="draft", draft_entry_id="", last_action=ACTION_SKIPPED,
    )


def sync_sent(
    store: Optional[RemovalStore] = None,
    settings: Optional[AuraSettings] = None,
    backend: Optional[DraftBackend] = None,
) -> SyncReport:
    store = store or RemovalStore()
    settings = settings or load_settings()
    report = SyncReport()
    ready = store.list_requests(status="ready")
    if not ready:
        return report
    backend = backend or OutlookDrafts()
    for req in ready:
        try:
            if req.draft_entry_id and backend.in_drafts(req.draft_entry_id):
                continue
            broker = find_broker(req.broker_id) or {}
            to = str(broker.get("opt_out_email") or "")
            sent_on = backend.find_sent(to, req.subject, _parse_iso(req.created_at)) if to else None
            if sent_on is not None:
                store.update(
                    req.id,
                    status="sent",
                    draft_entry_id="",
                    last_action=f"sent {sent_on.strftime('%Y-%m-%d %H:%M')}",
                    follow_up_due=(sent_on + timedelta(days=settings.follow_up_days)).strftime("%Y-%m-%d"),
                )
                report.sent.append(req.broker_name)
            else:
                store.update(
                    req.id,
                    status="draft",
                    draft_entry_id="",
                    last_action=ACTION_DRAFT_DELETED,
                    notes=(req.notes + "\nOutlook draft deleted without sending.").strip(),
                )
                report.reverted.append(req.broker_name)
        except Exception as exc:
            log.exception("Sync failed for %s", req.broker_name)
            report.errors.append(f"{req.broker_name}: {exc}")
    return report


def open_draft(request: RemovalRequest, backend: Optional[DraftBackend] = None) -> None:
    if not request.draft_entry_id:
        raise LookupError("This request has no Outlook draft yet.")
    (backend or OutlookDrafts()).display(request.draft_entry_id)


def prepare_and_sync() -> str:
    """Both passes on one Outlook connection. Caller owns the COM apartment."""
    backend = OutlookDrafts()
    rep = prepare(backend=backend)
    if rep.blocked:
        return rep.summary()
    syn = sync_sent(backend=backend)
    return f"{rep.summary()}; {syn.summary()}"


def main(argv: Optional[List[str]] = None) -> int:
    parser = argparse.ArgumentParser(description="Aura: prepare removal drafts / sync sent state.")
    parser.add_argument("--prepare", action="store_true", help="create Outlook drafts for review")
    parser.add_argument("--sync", action="store_true", help="mark sent drafts as sent")
    args = parser.parse_args(argv)
    logging.basicConfig(level=logging.INFO, format="%(levelname)s %(message)s")
    if not (args.prepare or args.sync):
        parser.print_help()
        return 2
    with _com_apartment():
        backend = OutlookDrafts()
        if args.prepare:
            rep = prepare(backend=backend)
            print(rep.summary())
            for err in rep.errors:
                print("  " + err)
            if rep.blocked:
                return 1
        if args.sync:
            print(sync_sent(backend=backend).summary())
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
