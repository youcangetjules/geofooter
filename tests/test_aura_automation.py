#!/usr/bin/env python3
"""Aura automation tests with a fake draft backend (no Outlook / pywin32)."""

from __future__ import annotations

import sqlite3
from datetime import datetime, timezone
from typing import Dict, List, Optional

import pytest

from aura import automation, pending, settings as aura_settings
from aura.catalog import load_catalog
from aura.profile import IdentityProfile
from aura.settings import AuraSettings
from aura.store import RemovalStore


class FakeDrafts:
    def __init__(self) -> None:
        self.drafts: Dict[str, Dict[str, str]] = {}
        self.sent: List[Dict[str, str]] = []
        self._n = 0

    def create(self, to: str, subject: str, body: str, send_account: str) -> str:
        self._n += 1
        eid = f"EID{self._n}"
        self.drafts[eid] = {"to": to, "subject": subject, "body": body}
        return eid

    def in_drafts(self, entry_id: str) -> bool:
        return entry_id in self.drafts

    def find_sent(self, to: str, subject: str, since: datetime) -> Optional[datetime]:
        for s in self.sent:
            if s["to"] == to and s["subject"] == subject:
                return datetime(2026, 9, 25, 10, 0)
        return None

    def display(self, entry_id: str) -> None:
        pass

    def delete(self, entry_id: str) -> None:
        self.drafts.pop(entry_id, None)

    def accounts(self) -> List[str]:
        return ["me@example.com"]

    def user_sends(self, entry_id: str) -> None:
        self.sent.append(self.drafts.pop(entry_id))


@pytest.fixture
def env(tmp_path, monkeypatch):
    monkeypatch.setattr(pending, "pending_path", lambda: tmp_path / "pending.json")
    monkeypatch.setattr(aura_settings, "settings_path", lambda: tmp_path / "aura_settings.json")
    store = RemovalStore(tmp_path / "requests.sqlite")
    profile = IdentityProfile(full_name="Test Person", emails=["me@example.com"])
    return store, profile, AuraSettings(), FakeDrafts()


def _email_broker_ids() -> List[str]:
    return [b["id"] for b in load_catalog() if b.get("opt_out_email")]


def test_detect_since_blocks_backdated_mail(env) -> None:
    before = datetime(2026, 9, 23, 18, 0, tzinfo=timezone.utc).isoformat()
    after = datetime(2026, 9, 24, 9, 0, tzinfo=timezone.utc).isoformat()
    assert pending.enqueue_from_aes(sender="x@pipl.com", received_at=before) is None
    assert pending.list_pending() == []
    item = pending.enqueue_from_aes(sender="x@pipl.com", received_at=after)
    assert item is not None
    assert [p["id"] for p in pending.list_pending()] == [item["id"]]


def test_prepare_blocks_without_identity(env) -> None:
    store, _profile, cfg, fake = env
    rep = automation.prepare(store, IdentityProfile(), cfg, fake)
    assert rep.blocked
    assert fake.drafts == {}


def test_prepare_drafts_email_brokers_and_tracks_web_forms(env) -> None:
    store, profile, cfg, fake = env
    rep = automation.prepare(store, profile, cfg, fake)
    assert not rep.errors
    assert sorted(rep.created) == sorted(
        b["name"] for b in load_catalog() if b.get("opt_out_email")
    )
    assert len(fake.drafts) == len(_email_broker_ids())
    web_forms = [b for b in load_catalog() if not b.get("opt_out_email")]
    assert len(rep.tracked) == len(web_forms)
    for bid in _email_broker_ids():
        req = store.latest_for_broker(bid)
        assert req is not None and req.status == "ready" and req.draft_entry_id


def test_prepare_is_idempotent_and_respects_skip(env) -> None:
    store, profile, cfg, fake = env
    automation.prepare(store, profile, cfg, fake)
    rep2 = automation.prepare(store, profile, cfg, fake)
    assert rep2.created == [] and rep2.tracked == []

    bid = _email_broker_ids()[0]
    req = store.latest_for_broker(bid)
    automation.skip(req.id, store=store, backend=fake)
    rep3 = automation.prepare(store, profile, cfg, fake)
    assert rep3.created == []
    assert store.latest_for_broker(bid).status == "draft"


def test_sync_marks_sent_and_reverts_deleted(env) -> None:
    store, profile, cfg, fake = env
    automation.prepare(store, profile, cfg, fake)
    ids = _email_broker_ids()
    sent_req = store.latest_for_broker(ids[0])
    deleted_req = store.latest_for_broker(ids[1])
    fake.user_sends(sent_req.draft_entry_id)
    fake.delete(deleted_req.draft_entry_id)

    rep = automation.sync_sent(store, cfg, fake)
    assert rep.sent == [sent_req.broker_name]
    assert rep.reverted == [deleted_req.broker_name]
    done = store.get(sent_req.id)
    assert done.status == "sent" and done.follow_up_due == "2026-10-25"
    assert store.get(deleted_req.id).status == "draft"


def test_store_migrates_draft_entry_id(tmp_path) -> None:
    path = tmp_path / "legacy.sqlite"
    with sqlite3.connect(path) as conn:
        conn.execute(
            "CREATE TABLE requests (id TEXT PRIMARY KEY, broker_id TEXT NOT NULL, "
            "broker_name TEXT NOT NULL, status TEXT NOT NULL, created_at TEXT NOT NULL, "
            "updated_at TEXT NOT NULL, notes TEXT DEFAULT '', last_action TEXT DEFAULT '', "
            "follow_up_due TEXT DEFAULT '', source TEXT DEFAULT 'manual', sender TEXT DEFAULT '', "
            "domain TEXT DEFAULT '', subject TEXT DEFAULT '', guri TEXT DEFAULT '', "
            "evidence TEXT DEFAULT '')"
        )
        conn.execute(
            "INSERT INTO requests (id, broker_id, broker_name, status, created_at, updated_at) "
            "VALUES ('r1', 'pipl', 'Pipl', 'draft', 'x', 'x')"
        )
    store = RemovalStore(path)
    req = store.get("r1")
    assert req is not None and req.draft_entry_id == ""
    store.update("r1", draft_entry_id="EID9")
    assert store.get("r1").draft_entry_id == "EID9"
