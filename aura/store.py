"""SQLite store for data-broker removal requests."""

from __future__ import annotations

import os
import sqlite3
import uuid
from dataclasses import dataclass
from datetime import datetime, timezone
from pathlib import Path
from typing import List, Optional


STATUSES = (
    "draft",
    "ready",
    "sent",
    "awaiting",
    "confirmed",
    "denied",
    "follow_up",
)


def data_dir() -> Path:
    local = os.environ.get("LOCALAPPDATA", "")
    return Path(local) / "GeoFooter" if local else Path(r"C:\GeoFooter")


def db_path() -> Path:
    """SQLite file under GeoFooter/datastore/ (migrates legacy flat path once)."""
    store = data_dir() / "datastore"
    store.mkdir(parents=True, exist_ok=True)
    path = store / "broker_removal.sqlite"
    legacy = data_dir() / "broker_removal.sqlite"
    if not path.is_file() and legacy.is_file():
        try:
            legacy.replace(path)
        except Exception:
            pass
    return path


def _now_iso() -> str:
    return datetime.now(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ")


@dataclass
class RemovalRequest:
    id: str
    broker_id: str
    broker_name: str
    status: str
    created_at: str
    updated_at: str
    notes: str = ""
    last_action: str = ""
    follow_up_due: str = ""
    source: str = "manual"  # manual | aes
    sender: str = ""
    domain: str = ""
    subject: str = ""
    guri: str = ""
    evidence: str = ""
    draft_entry_id: str = ""


# Requests in these states already cover the broker; don't prepare another.
ACTIVE_STATUSES = ("ready", "sent", "awaiting", "confirmed", "denied")


class RemovalStore:
    def __init__(self, path: Optional[Path] = None) -> None:
        self.path = path or db_path()
        self.path.parent.mkdir(parents=True, exist_ok=True)
        self._init_db()

    def _connect(self) -> sqlite3.Connection:
        conn = sqlite3.connect(str(self.path))
        conn.row_factory = sqlite3.Row
        return conn

    def _init_db(self) -> None:
        with self._connect() as conn:
            conn.execute(
                """
                CREATE TABLE IF NOT EXISTS requests (
                    id TEXT PRIMARY KEY,
                    broker_id TEXT NOT NULL,
                    broker_name TEXT NOT NULL,
                    status TEXT NOT NULL,
                    created_at TEXT NOT NULL,
                    updated_at TEXT NOT NULL,
                    notes TEXT DEFAULT '',
                    last_action TEXT DEFAULT '',
                    follow_up_due TEXT DEFAULT '',
                    source TEXT DEFAULT 'manual',
                    sender TEXT DEFAULT '',
                    domain TEXT DEFAULT '',
                    subject TEXT DEFAULT '',
                    guri TEXT DEFAULT '',
                    evidence TEXT DEFAULT ''
                )
                """
            )
            conn.execute(
                "CREATE INDEX IF NOT EXISTS idx_requests_broker ON requests(broker_id)"
            )
            conn.execute(
                "CREATE INDEX IF NOT EXISTS idx_requests_status ON requests(status)"
            )
            cols = {r[1] for r in conn.execute("PRAGMA table_info(requests)")}
            if "draft_entry_id" not in cols:
                conn.execute(
                    "ALTER TABLE requests ADD COLUMN draft_entry_id TEXT DEFAULT ''"
                )
            conn.commit()

    @staticmethod
    def _row(row: sqlite3.Row) -> RemovalRequest:
        return RemovalRequest(
            id=str(row["id"]),
            broker_id=str(row["broker_id"]),
            broker_name=str(row["broker_name"]),
            status=str(row["status"]),
            created_at=str(row["created_at"]),
            updated_at=str(row["updated_at"]),
            notes=str(row["notes"] or ""),
            last_action=str(row["last_action"] or ""),
            follow_up_due=str(row["follow_up_due"] or ""),
            source=str(row["source"] or "manual"),
            sender=str(row["sender"] or ""),
            domain=str(row["domain"] or ""),
            subject=str(row["subject"] or ""),
            guri=str(row["guri"] or ""),
            evidence=str(row["evidence"] or ""),
            draft_entry_id=str(row["draft_entry_id"] or ""),
        )

    def latest_for_broker(self, broker_id: str) -> Optional[RemovalRequest]:
        with self._connect() as conn:
            cur = conn.execute(
                "SELECT * FROM requests WHERE broker_id = ? ORDER BY updated_at DESC LIMIT 1",
                (broker_id,),
            )
            row = cur.fetchone()
            return self._row(row) if row else None

    def list_requests(self, status: Optional[str] = None) -> List[RemovalRequest]:
        with self._connect() as conn:
            if status:
                cur = conn.execute(
                    "SELECT * FROM requests WHERE status = ? ORDER BY updated_at DESC",
                    (status,),
                )
            else:
                cur = conn.execute(
                    "SELECT * FROM requests ORDER BY updated_at DESC"
                )
            return [self._row(r) for r in cur.fetchall()]

    def get(self, request_id: str) -> Optional[RemovalRequest]:
        with self._connect() as conn:
            cur = conn.execute(
                "SELECT * FROM requests WHERE id = ?", (request_id,)
            )
            row = cur.fetchone()
            return self._row(row) if row else None

    def create(
        self,
        *,
        broker_id: str,
        broker_name: str,
        status: str = "draft",
        source: str = "manual",
        sender: str = "",
        domain: str = "",
        subject: str = "",
        guri: str = "",
        notes: str = "",
        evidence: str = "",
        draft_entry_id: str = "",
    ) -> RemovalRequest:
        status = status if status in STATUSES else "draft"
        now = _now_iso()
        req = RemovalRequest(
            id=uuid.uuid4().hex[:16],
            broker_id=broker_id,
            broker_name=broker_name,
            status=status,
            created_at=now,
            updated_at=now,
            notes=notes,
            last_action="created",
            source=source,
            sender=sender,
            domain=domain,
            subject=subject,
            guri=guri,
            evidence=evidence,
            draft_entry_id=draft_entry_id,
        )
        with self._connect() as conn:
            conn.execute(
                """
                INSERT INTO requests (
                    id, broker_id, broker_name, status, created_at, updated_at,
                    notes, last_action, follow_up_due, source, sender, domain,
                    subject, guri, evidence, draft_entry_id
                ) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)
                """,
                (
                    req.id,
                    req.broker_id,
                    req.broker_name,
                    req.status,
                    req.created_at,
                    req.updated_at,
                    req.notes,
                    req.last_action,
                    req.follow_up_due,
                    req.source,
                    req.sender,
                    req.domain,
                    req.subject,
                    req.guri,
                    req.evidence,
                    req.draft_entry_id,
                ),
            )
            conn.commit()
        return req

    def update(
        self,
        request_id: str,
        *,
        status: Optional[str] = None,
        notes: Optional[str] = None,
        follow_up_due: Optional[str] = None,
        last_action: Optional[str] = None,
        evidence: Optional[str] = None,
        subject: Optional[str] = None,
        source: Optional[str] = None,
        draft_entry_id: Optional[str] = None,
    ) -> Optional[RemovalRequest]:
        req = self.get(request_id)
        if not req:
            return None
        if subject is not None:
            req.subject = subject
        if source is not None:
            req.source = source
        if draft_entry_id is not None:
            req.draft_entry_id = draft_entry_id
        if status is not None and status in STATUSES:
            req.status = status
        if notes is not None:
            req.notes = notes
        if follow_up_due is not None:
            req.follow_up_due = follow_up_due
        if last_action is not None:
            req.last_action = last_action
        if evidence is not None:
            req.evidence = evidence
        req.updated_at = _now_iso()
        with self._connect() as conn:
            conn.execute(
                """
                UPDATE requests SET
                    status=?, notes=?, follow_up_due=?, last_action=?,
                    evidence=?, subject=?, source=?, draft_entry_id=?,
                    updated_at=?
                WHERE id=?
                """,
                (
                    req.status,
                    req.notes,
                    req.follow_up_due,
                    req.last_action,
                    req.evidence,
                    req.subject,
                    req.source,
                    req.draft_entry_id,
                    req.updated_at,
                    req.id,
                ),
            )
            conn.commit()
        return req

    def delete(self, request_id: str) -> bool:
        with self._connect() as conn:
            cur = conn.execute("DELETE FROM requests WHERE id = ?", (request_id,))
            conn.commit()
            return cur.rowcount > 0
