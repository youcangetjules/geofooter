#!/usr/bin/env python3
"""Live Outlook inbox scraper for GURI action / importance detection."""

from __future__ import annotations

import json
import logging
import os
import re
from dataclasses import asdict, dataclass, field
from datetime import datetime, timedelta
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple

logger = logging.getLogger(__name__)

OL_FOLDER_INBOX = 6
OL_FOLDER_SENTMAIL = 5
OL_MAIL_ITEM = 43
OL_MEETING_REQUEST = 53


class _OutlookComScope:
    """Ensure COM is initialized on the calling thread (required for Outlook STA)."""

    def __enter__(self):
        self._pythoncom = None
        try:
            import pythoncom  # type: ignore

            pythoncom.CoInitialize()
            self._pythoncom = pythoncom
        except Exception:
            self._pythoncom = None
        return self

    def __exit__(self, exc_type, exc, tb):
        if self._pythoncom is not None:
            try:
                self._pythoncom.CoUninitialize()
            except Exception:
                pass
        return False


@dataclass
class AesAccountScanConfig:
    store_id: str
    smtp: str = ""
    display: str = ""
    enabled: bool = True
    responses: bool = True
    in_cc: bool = True

    def to_dict(self) -> Dict[str, Any]:
        return {
            "store_id": self.store_id,
            "smtp": self.smtp,
            "display": self.display,
            "enabled": bool(self.enabled),
            "responses": bool(self.responses),
            "in_cc": bool(self.in_cc),
        }


@dataclass
class ScrapedMail:
    entry_id: str
    store_id: str
    account_smtp: str
    subject: str
    sender: str
    to: str
    cc: str
    received: str  # ISO-ish local time string
    received_dt: Optional[datetime]
    unread: bool
    importance: int  # Outlook OlImportance 0=low 1=normal 2=high
    flag_status: int
    message_class: str
    conversation: str
    body_preview: str
    recipient_role: str  # to | cc | unknown
    is_meeting_request: bool = False
    conversation_id: str = ""
    folder: str = "Inbox"

    def to_dict(self) -> Dict[str, Any]:
        d = asdict(self)
        d.pop("received_dt", None)
        return d


def _local_geofooter() -> Path:
    local = os.environ.get("LOCALAPPDATA", "")
    if local:
        return Path(local) / "GeoFooter"
    return Path(r"C:\GeoFooter")


def load_aes_account_configs() -> List[AesAccountScanConfig]:
    """Load enabled-account scan flags from AES settings INI (and optional JSON)."""
    configs: Dict[str, AesAccountScanConfig] = {}

    ini_path = _local_geofooter() / "aes_scan_accounts.ini"
    if ini_path.is_file():
        try:
            for raw in ini_path.read_text(encoding="utf-8", errors="replace").splitlines():
                line = raw.strip()
                if not line or line.startswith(";") or line.startswith("#") or "=" not in line:
                    continue
                key, value = line.split("=", 1)
                key = key.strip()
                value = value.strip()
                smtp = ""
                store_id = key
                if "|" in key:
                    store_id, smtp = key.split("|", 1)
                    store_id = store_id.strip()
                    smtp = smtp.strip()
                enabled, responses, in_cc = True, True, True
                parts = [p.strip() for p in value.split(",")]
                if parts:
                    enabled = parts[0] in {"1", "true", "TRUE", "ON", "on"}
                if len(parts) >= 2:
                    responses = parts[1] in {"1", "true", "TRUE", "ON", "on"}
                if len(parts) >= 3:
                    in_cc = parts[2] in {"1", "true", "TRUE", "ON", "on"}
                if store_id:
                    configs[store_id] = AesAccountScanConfig(
                        store_id=store_id,
                        smtp=smtp,
                        enabled=enabled,
                        responses=responses,
                        in_cc=in_cc,
                    )
        except Exception as exc:
            logger.warning("Failed reading AES account INI: %s", exc)

    # Optional richer JSON from last settings dialog write
    json_path = _local_geofooter() / "aes_settings_result.json"
    if json_path.is_file():
        try:
            payload = json.loads(json_path.read_text(encoding="utf-8"))
            for acc in payload.get("accounts") or []:
                store_id = str(acc.get("store_id") or "").strip()
                if not store_id:
                    continue
                existing = configs.get(store_id)
                configs[store_id] = AesAccountScanConfig(
                    store_id=store_id,
                    smtp=str(acc.get("smtp") or (existing.smtp if existing else "")),
                    display=str(
                        acc.get("display")
                        or (existing.display if existing else "")
                        or acc.get("smtp")
                        or ""
                    ),
                    enabled=bool(acc.get("enabled", existing.enabled if existing else True)),
                    responses=bool(acc.get("responses", existing.responses if existing else True)),
                    in_cc=bool(acc.get("in_cc", existing.in_cc if existing else True)),
                )
        except Exception as exc:
            logger.warning("Failed reading AES settings result JSON: %s", exc)

    return list(configs.values())


def aes_accounts_ini_path() -> Path:
    return _local_geofooter() / "aes_scan_accounts.ini"


def aes_accounts_json_path() -> Path:
    return _local_geofooter() / "aes_settings_accounts.json"


def aes_settings_result_path() -> Path:
    return _local_geofooter() / "aes_settings_result.json"


def discover_outlook_accounts() -> List[AesAccountScanConfig]:
    """Enumerate Outlook mail accounts (DeliveryStore + SMTP) for the Accounts UI."""
    with _OutlookComScope():
        return _discover_outlook_accounts_impl()


def _discover_outlook_accounts_impl() -> List[AesAccountScanConfig]:
    app = _outlook_app()
    ns = app.GetNamespace("MAPI")
    saved = {c.store_id: c for c in load_aes_account_configs()}
    out: List[AesAccountScanConfig] = []
    seen: set = set()

    try:
        count = int(ns.Accounts.Count)
    except Exception:
        count = 0

    for i in range(1, count + 1):
        try:
            acc = ns.Accounts.Item(i)
        except Exception:
            continue
        store_id = ""
        display = ""
        smtp = ""
        try:
            display = str(getattr(acc, "DisplayName", "") or "")
        except Exception:
            display = ""
        try:
            smtp = str(getattr(acc, "SmtpAddress", "") or "")
        except Exception:
            smtp = ""
        try:
            delivery = acc.DeliveryStore
            if delivery is not None:
                store_id = str(delivery.StoreID or "")
                if not display:
                    display = str(getattr(delivery, "DisplayName", "") or "")
        except Exception:
            store_id = ""
        if not store_id or store_id in seen:
            continue
        seen.add(store_id)
        prev = saved.get(store_id)
        out.append(
            AesAccountScanConfig(
                store_id=store_id,
                smtp=smtp or (prev.smtp if prev else ""),
                display=display or smtp or store_id[:24],
                enabled=prev.enabled if prev else True,
                responses=prev.responses if prev else True,
                in_cc=prev.in_cc if prev else True,
            )
        )

    # Keep any saved accounts Outlook didn't return (disconnected / temporary)
    for store_id, prev in saved.items():
        if store_id not in seen:
            out.append(prev)

    out.sort(key=lambda c: (c.display or c.smtp or c.store_id).lower())
    return out


def save_aes_account_configs(configs: List[AesAccountScanConfig]) -> Path:
    """Persist account scan flags for AES + GURI (INI + JSON companions)."""
    base = _local_geofooter()
    base.mkdir(parents=True, exist_ok=True)
    ini_path = aes_accounts_ini_path()
    lines = [
        "; AES scan account settings",
        "; store_id|smtp=enabled,responses,in_cc  (1=on, 0=off)",
        f"; Updated: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}",
    ]
    for cfg in configs:
        if not cfg.store_id:
            continue
        flag = "1" if cfg.enabled else "0"
        resp = "1" if cfg.responses else "0"
        cc = "1" if cfg.in_cc else "0"
        smtp = (cfg.smtp or "").replace("\n", " ").strip()
        lines.append(f"{cfg.store_id}|{smtp}={flag},{resp},{cc}")
    ini_path.write_text("\n".join(lines) + "\n", encoding="utf-8")

    accounts_payload = {"accounts": [c.to_dict() for c in configs if c.store_id]}
    try:
        aes_accounts_json_path().write_text(
            json.dumps(accounts_payload, indent=2),
            encoding="utf-8",
        )
    except Exception as exc:
        logger.warning("Failed writing aes_settings_accounts.json: %s", exc)

    # Merge into last settings-dialog result so AES VBA stays in sync
    result_path = aes_settings_result_path()
    result: Dict[str, Any] = {"cancelled": False, "accounts": []}
    try:
        if result_path.is_file():
            raw = json.loads(result_path.read_text(encoding="utf-8"))
            if isinstance(raw, dict):
                result.update(raw)
    except Exception:
        pass
    result["cancelled"] = False
    result["accounts"] = [
        {
            "store_id": c.store_id,
            "smtp": c.smtp,
            "display": c.display,
            "enabled": bool(c.enabled),
            "responses": bool(c.responses),
            "in_cc": bool(c.in_cc),
        }
        for c in configs
        if c.store_id
    ]
    try:
        result_path.write_text(json.dumps(result, indent=2), encoding="utf-8")
    except Exception as exc:
        logger.warning("Failed writing aes_settings_result.json: %s", exc)

    return ini_path


def _outlook_app():
    import win32com.client  # type: ignore

    try:
        return win32com.client.GetObject(Class="Outlook.Application")
    except Exception:
        return win32com.client.Dispatch("Outlook.Application")


def _safe_str(value: Any, limit: int = 0) -> str:
    text = "" if value is None else str(value)
    text = text.replace("\r\n", "\n").replace("\r", "\n").strip()
    if limit and len(text) > limit:
        return text[:limit]
    return text


def _resolve_sender(item) -> str:
    try:
        addr = getattr(item, "SenderEmailAddress", None) or ""
        name = getattr(item, "SenderName", None) or ""
        # Exchange DN → try PropertyAccessor SMTP
        if addr and "/o=" in addr.lower():
            try:
                smtp = item.PropertyAccessor.GetProperty(
                    "http://schemas.microsoft.com/mapi/proptag/0x5D01001E"
                )
                if smtp:
                    return str(smtp)
            except Exception:
                pass
        if addr and "@" in addr:
            return str(addr)
        if name and "@" in name:
            return str(name)
        return str(addr or name or "")
    except Exception:
        return ""


def _account_smtp_for_store(ns, store) -> str:
    try:
        for i in range(1, ns.Accounts.Count + 1):
            acc = ns.Accounts.Item(i)
            try:
                delivery = acc.DeliveryStore
                if delivery is not None and delivery.StoreID == store.StoreID:
                    return str(getattr(acc, "SmtpAddress", "") or acc.DisplayName or "")
            except Exception:
                continue
    except Exception:
        pass
    try:
        return str(store.DisplayName or "")
    except Exception:
        return ""


def _recipient_role_from_headers(to_line: str, cc_line: str, account_smtp: str) -> str:
    """Classify To/Cc from header strings only (no Recipients COM walk)."""
    smtp = (account_smtp or "").strip().lower()
    if not smtp or "@" not in smtp:
        return "unknown"
    to_l = (to_line or "").lower()
    cc_l = (cc_line or "").lower()
    local = smtp.split("@", 1)[0]
    if smtp in to_l or (local and local in to_l):
        return "to"
    if smtp in cc_l or (local and local in cc_l):
        return "cc"
    return "unknown"


def _recipient_role(item, account_smtp: str) -> str:
    """Classify whether this mailbox is on To or Cc (headers only — Outlook-safe)."""
    try:
        return _recipient_role_from_headers(
            _safe_str(getattr(item, "To", "")),
            _safe_str(getattr(item, "CC", "")),
            account_smtp,
        )
    except Exception:
        return "unknown"


def _as_naive_local(dt: datetime) -> datetime:
    """Drop tzinfo for safe compares.

    Outlook/pywin32 often tags ReceivedTime as UTC (+00:00) even when the
    clock value is already local wall time. Converting with astimezone()
    shifts display by the DST offset; stripping tzinfo keeps the stamp.
    """
    if getattr(dt, "tzinfo", None) is None:
        return dt
    return dt.replace(tzinfo=None)


def _parse_received(value: Any) -> Tuple[Optional[datetime], str]:
    if value is None:
        return None, ""
    if isinstance(value, datetime):
        dt = _as_naive_local(value)
        return dt, dt.strftime("%Y-%m-%d %H:%M:%S")
    try:
        # pywin32 may return time.struct-like or datetime
        dt = value
        if hasattr(value, "year") and hasattr(value, "month"):
            dt = datetime(
                int(value.year),
                int(value.month),
                int(value.day),
                int(getattr(value, "hour", 0) or 0),
                int(getattr(value, "minute", 0) or 0),
                int(getattr(value, "second", 0) or 0),
            )
            # Preserve tz if present on the COM value
            tz = getattr(value, "tzinfo", None)
            if tz is not None:
                dt = _as_naive_local(dt.replace(tzinfo=tz))
            return dt, dt.strftime("%Y-%m-%d %H:%M:%S")
    except Exception:
        pass
    text = _safe_str(value)
    return None, text


def _outlook_received_since_filter(since: datetime) -> str:
    """Locale-safe Outlook Restrict filter for ReceivedTime >= since.

    Jet filters like ``[ReceivedTime] >= '07/11/2026'`` are parsed with the
    Windows regional date order. On UK systems ``mm/dd`` is read as ``dd/mm``,
    so a 30+ day US-formatted lookback becomes a *future* date and returns 0
    messages with no error. DASL + ISO datetime avoids that.
    """
    stamp = since.strftime("%Y-%m-%d %H:%M")
    return f'@SQL="urn:schemas:httpmail:datereceived" >= \'{stamp}\''


_DASL_ADDR_SAFE = re.compile(r"[^a-z0-9@._+\-]")
_DASL_KEYWORD_SAFE = re.compile(r"[^a-z0-9+\-]")


def _dasl_quote(value: str) -> str:
    """Escape a value for inclusion in a DASL single-quoted string."""
    return (value or "").replace("\\", "\\\\").replace("'", "''")


def _outlook_important_senders_filter(
    since: datetime,
    important_domains: Optional[List[str]],
    important_senders: Optional[List[str]],
) -> Optional[str]:
    """DASL filter: mail since ``since`` whose sender matches an important
    domain or VIP address. Used as a second targeted Restrict so important
    mail is found across the whole lookback, not just the newest N items."""
    terms: List[str] = []
    for domain in (important_domains or [])[:25]:
        d = _DASL_ADDR_SAFE.sub("", str(domain).strip().lstrip("@").lower())
        if d and "." in d:
            terms.append(f'"urn:schemas:httpmail:fromemail" LIKE \'%{d}%\'')
    for sender in (important_senders or [])[:25]:
        s = _DASL_ADDR_SAFE.sub("", str(sender).strip().lower())
        if s and "@" in s:
            terms.append(f'"urn:schemas:httpmail:fromemail" LIKE \'%{s}%\'')
    if not terms:
        return None
    stamp = since.strftime("%Y-%m-%d %H:%M")
    return (
        '@SQL=("urn:schemas:httpmail:datereceived" >= \'' + stamp + "') AND ("
        + " OR ".join(terms)
        + ")"
    )


def _iter_inbox_since(items, since: datetime, *, max_take: int):
    """Yield Outlook items received on/after since (newest first).

    Prefer Restrict; if it yields nothing, fall back to a manual scan so a
    bad/locale filter cannot blank the Welcome tab.
    """
    restrict = _outlook_received_since_filter(since)
    filtered = None
    try:
        items.Sort("[ReceivedTime]", True)
    except Exception:
        pass
    try:
        filtered = items.Restrict(restrict)
        count = int(filtered.Count)
    except Exception:
        filtered = None
        count = 0

    if filtered is not None and count > 0:
        # Only materialize what we will keep — touching Item(n) is expensive.
        limit = min(count, max(1, int(max_take)))
        for idx in range(1, limit + 1):
            try:
                yield filtered.Item(idx)
            except Exception:
                continue
        return

    # Fallback: walk newest-first and stop once older than the window (capped)
    try:
        items.Sort("[ReceivedTime]", True)
        total = int(items.Count)
    except Exception:
        return
    yielded = 0
    scan_cap = max(1, int(max_take)) * 3
    for idx in range(1, min(total, scan_cap) + 1):
        if yielded >= max_take:
            break
        try:
            item = items.Item(idx)
        except Exception:
            continue
        received_dt, _ = _parse_received(getattr(item, "ReceivedTime", None))
        if received_dt is not None and received_dt < since:
            break
        yielded += 1
        yield item


def _mail_from_item(
    item,
    *,
    store_id: str,
    account_smtp: str,
    cfg: "AesAccountScanConfig",
    since: datetime,
    unread_only: bool,
    include_body: bool,
    folder: str = "Inbox",
    body_limit: int = 1200,
    ignore_since: bool = False,
) -> Optional[ScrapedMail]:
    """Convert one Outlook item to ScrapedMail, or None if filtered out."""
    try:
        # 43 = MailItem; meeting request often still mail-like
        msg_class = _safe_str(getattr(item, "MessageClass", ""))
        is_meeting = msg_class.lower().startswith("ipm.schedule.meeting")
        try:
            item_class = int(getattr(item, "Class", 0) or 0)
        except Exception:
            item_class = 0
        if item_class not in (OL_MAIL_ITEM, OL_MEETING_REQUEST) and not is_meeting:
            if not msg_class.lower().startswith("ipm.note") and not is_meeting:
                return None

        unread = bool(getattr(item, "UnRead", False))
        if unread_only and not unread:
            return None

        # Header strings only — never walk Recipients / GetExchangeUser
        to_line = _safe_str(getattr(item, "To", ""), limit=1000)
        cc_line = _safe_str(getattr(item, "CC", ""), limit=1000)
        role = _recipient_role_from_headers(to_line, cc_line, account_smtp)
        if role == "cc" and not cfg.in_cc:
            return None

        received_dt, received_str = _parse_received(
            getattr(item, "ReceivedTime", None)
        )
        if (
            not ignore_since
            and received_dt is not None
            and received_dt < since
        ):
            return None

        limit = max(200, int(body_limit or 1200))
        # Body is the main Outlook freeze source — skip unless requested
        body = ""
        if include_body:
            try:
                body = _safe_str(
                    item.PropertyAccessor.GetProperty(
                        "urn:schemas:httpmail:textdescription"
                    ),
                    limit=limit,
                )
            except Exception:
                try:
                    body = _safe_str(getattr(item, "Body", ""), limit=limit)
                except Exception:
                    body = ""

        conversation_id = ""
        try:
            conversation_id = _safe_str(
                getattr(item, "ConversationID", ""), limit=200
            )
        except Exception:
            conversation_id = ""

        mail = ScrapedMail(
            entry_id=_safe_str(getattr(item, "EntryID", "")),
            store_id=store_id,
            account_smtp=account_smtp,
            subject=_safe_str(getattr(item, "Subject", ""), limit=500),
            sender=_resolve_sender(item),
            to=to_line,
            cc=cc_line,
            received=received_str,
            received_dt=received_dt,
            unread=unread,
            importance=int(getattr(item, "Importance", 1) or 1),
            flag_status=int(getattr(item, "FlagStatus", 0) or 0),
            message_class=msg_class,
            conversation=_safe_str(
                getattr(item, "ConversationTopic", ""), limit=500
            ),
            body_preview=body,
            recipient_role=role,
            is_meeting_request=is_meeting,
            conversation_id=conversation_id,
            folder=folder or "Inbox",
        )
        if not mail.entry_id:
            return None
        return mail
    except Exception as exc:
        logger.debug("Skip item in %s: %s", account_smtp, exc)
        return None


def scrape_outlook_inboxes(
    *,
    days: int = 7,
    max_per_account: int = 200,
    unread_only: bool = False,
    account_configs: Optional[List[AesAccountScanConfig]] = None,
    include_body: bool = False,
    important_domains: Optional[List[str]] = None,
    important_senders: Optional[List[str]] = None,
) -> List[ScrapedMail]:
    """Scrape recent Inbox mail from AES-enabled accounts (or all if none configured).

    ``include_body`` defaults False — reading MailItem.Body across many stores
    freezes the Outlook UI. Pass True only for an explicit deep manual scrape.

    ``important_domains`` / ``important_senders`` add a second, targeted
    Restrict per store: sender-matched mail across the *whole* lookback window.
    Without it, high-volume inboxes push older important mail past the
    ``max_per_account`` recency cap and it never appears.
    """
    with _OutlookComScope():
        return _scrape_outlook_inboxes_impl(
            days=days,
            max_per_account=max_per_account,
            unread_only=unread_only,
            account_configs=account_configs,
            include_body=include_body,
            important_domains=important_domains,
            important_senders=important_senders,
        )


def _scrape_outlook_inboxes_impl(
    *,
    days: int = 7,
    max_per_account: int = 200,
    unread_only: bool = False,
    account_configs: Optional[List[AesAccountScanConfig]] = None,
    include_body: bool = False,
    important_domains: Optional[List[str]] = None,
    important_senders: Optional[List[str]] = None,
) -> List[ScrapedMail]:
    try:
        app = _outlook_app()
        ns = app.GetNamespace("MAPI")
    except Exception as exc:
        raise RuntimeError(f"Outlook is not available: {exc}") from exc

    configs = account_configs if account_configs is not None else load_aes_account_configs()
    enabled = [c for c in configs if c.enabled]
    if configs and not enabled:
        # Never fall back to "scrape every store" when accounts are configured
        # but all switched off — that scans stores the user opted out of and
        # freezes Outlook. Fail loudly so the UI can tell the user.
        raise RuntimeError(
            "All AES email accounts are disabled — nothing to scrape. "
            "Enable at least one account in AES Settings / GURI Accounts tab."
        )
    config_by_store = {c.store_id: c for c in enabled}

    since = datetime.now() - timedelta(days=max(1, int(days)))
    important_filter = _outlook_important_senders_filter(
        since, important_domains, important_senders
    )

    results: List[ScrapedMail] = []
    seen_entry_ids: set = set()
    stores_scraped = 0

    try:
        store_count = ns.Stores.Count
    except Exception as exc:
        raise RuntimeError(f"Cannot enumerate Outlook stores: {exc}") from exc

    for si in range(1, store_count + 1):
        try:
            store = ns.Stores.Item(si)
            store_id = str(store.StoreID)
        except Exception:
            continue

        cfg = config_by_store.get(store_id)
        if enabled and cfg is None:
            # Also try matching by smtp against store display / account
            account_smtp = _account_smtp_for_store(ns, store)
            cfg = next(
                (
                    c
                    for c in enabled
                    if c.smtp and c.smtp.lower() == account_smtp.lower()
                ),
                None,
            )
            if cfg is None:
                continue
        elif not enabled:
            cfg = AesAccountScanConfig(store_id=store_id, smtp=_account_smtp_for_store(ns, store))

        account_smtp = cfg.smtp or _account_smtp_for_store(ns, store)

        try:
            inbox = store.GetDefaultFolder(OL_FOLDER_INBOX)
            items = inbox.Items
        except Exception as exc:
            logger.warning("Skip store %s: %s", account_smtp or store_id[:20], exc)
            continue

        stores_scraped += 1
        taken = 0
        for item in _iter_inbox_since(items, since, max_take=max_per_account):
            if taken >= max_per_account:
                break
            mail = _mail_from_item(
                item,
                store_id=store_id,
                account_smtp=account_smtp,
                cfg=cfg,
                since=since,
                unread_only=unread_only,
                include_body=include_body,
            )
            if mail is None or mail.entry_id in seen_entry_ids:
                continue
            seen_entry_ids.add(mail.entry_id)
            results.append(mail)
            taken += 1

        # Targeted pass: important senders/domains over the FULL lookback.
        # The recency-capped pass above misses older important mail in busy
        # inboxes; this Restrict pulls it regardless of overall volume.
        if important_filter:
            try:
                flagged = items.Restrict(important_filter)
                try:
                    flagged.Sort("[ReceivedTime]", True)
                except Exception:
                    pass
                flagged_count = int(flagged.Count)
            except Exception as exc:
                logger.debug(
                    "Important-sender Restrict failed for %s: %s",
                    account_smtp or store_id[:20],
                    exc,
                )
                flagged = None
                flagged_count = 0

            important_taken = 0
            important_cap = max(50, int(max_per_account))
            for idx in range(1, flagged_count + 1):
                if important_taken >= important_cap:
                    break
                try:
                    item = flagged.Item(idx)
                except Exception:
                    continue
                mail = _mail_from_item(
                    item,
                    store_id=store_id,
                    account_smtp=account_smtp,
                    cfg=cfg,
                    since=since,
                    unread_only=unread_only,
                    include_body=include_body,
                )
                if mail is None or mail.entry_id in seen_entry_ids:
                    continue
                seen_entry_ids.add(mail.entry_id)
                results.append(mail)
                important_taken += 1
            if important_taken:
                logger.info(
                    "Targeted important-sender pass: +%d mail(s) from %s",
                    important_taken,
                    account_smtp or store_id[:20],
                )

    # Newest first
    results.sort(key=lambda m: m.received_dt or datetime.min, reverse=True)
    logger.info(
        "Outlook scrape: %d messages from %d store(s) (days=%s, body=%s)",
        len(results),
        stores_scraped,
        days,
        include_body,
    )
    return results


def open_mail_in_outlook(entry_id: str, store_id: str = "") -> bool:
    """Display an Outlook item by EntryID (and optional StoreID)."""
    if not entry_id:
        return False
    with _OutlookComScope():
        app = _outlook_app()
        ns = app.GetNamespace("MAPI")
        try:
            if store_id:
                item = ns.GetItemFromID(entry_id, store_id)
            else:
                item = ns.GetItemFromID(entry_id)
            item.Display()
            return True
        except Exception as exc:
            logger.error("Open in Outlook failed: %s", exc)
            return False


def _outlook_keyword_subject_filter(
    since: datetime,
    keywords: Optional[List[str]],
) -> Optional[str]:
    """DASL filter: mail since ``since`` whose subject matches any keyword.

    Subject-only — body LIKE across a store is too expensive for Outlook.
    Used as a second pass so themed threads older than the recency cap still
    surface when the query names a distinctive topic.
    """
    terms: List[str] = []
    seen: set = set()
    for raw in (keywords or [])[:8]:
        word = _DASL_KEYWORD_SAFE.sub("", str(raw).strip().lower())
        if len(word) < 3 or word in seen:
            continue
        seen.add(word)
        terms.append(
            f"\"urn:schemas:httpmail:subject\" LIKE '%{_dasl_quote(word)}%'"
        )
    if not terms:
        return None
    stamp = since.strftime("%Y-%m-%d %H:%M")
    return (
        '@SQL=("urn:schemas:httpmail:datereceived" >= \''
        + stamp
        + "') AND ("
        + " OR ".join(terms)
        + ")"
    )


def _find_store(ns, *, store_id: str = "", smtp: str = ""):
    """Return (store, resolved_store_id) matching store_id or SMTP, else None."""
    want_store = (store_id or "").strip()
    want_smtp = (smtp or "").strip().lower()
    try:
        store_count = int(ns.Stores.Count)
    except Exception:
        return None
    for si in range(1, store_count + 1):
        try:
            store = ns.Stores.Item(si)
            sid = str(store.StoreID or "")
        except Exception:
            continue
        if want_store and sid == want_store:
            return store, sid
        if want_smtp:
            account_smtp = _account_smtp_for_store(ns, store)
            if account_smtp and account_smtp.lower() == want_smtp:
                return store, sid
            try:
                display = str(getattr(store, "DisplayName", "") or "")
            except Exception:
                display = ""
            if display and display.lower() == want_smtp:
                return store, sid
    return None


def scrape_inbox_for_search(
    *,
    store_id: str = "",
    smtp: str = "",
    days: int = 30,
    max_messages: int = 150,
    include_sent: bool = True,
    unread_only: bool = False,
    keywords: Optional[List[str]] = None,
    progress: Optional[Any] = None,
) -> List[ScrapedMail]:
    """Scrape one Outlook mailbox (Inbox, optionally Sent) with bodies.

    Unlike the live Welcome scrape this is a deliberate, mailbox-scoped
    conversation harvest: bodies are included so sentiment and subject
    matter can be scored across a thread, not a single header.
    """
    if not (store_id or "").strip() and not (smtp or "").strip():
        raise ValueError("Choose an inbox (store or SMTP) before scanning.")
    with _OutlookComScope():
        return _scrape_inbox_for_search_impl(
            store_id=store_id,
            smtp=smtp,
            days=days,
            max_messages=max_messages,
            include_sent=include_sent,
            unread_only=unread_only,
            keywords=keywords,
            progress=progress,
        )


def _emit_progress(progress: Optional[Any], message: str) -> None:
    if callable(progress):
        try:
            progress(message)
        except Exception:
            pass


def _scrape_folder_items(
    items,
    *,
    store_id: str,
    account_smtp: str,
    cfg: AesAccountScanConfig,
    since: datetime,
    unread_only: bool,
    folder: str,
    max_take: int,
    seen_entry_ids: set,
    body_limit: int,
    keyword_filter: Optional[str] = None,
) -> List[ScrapedMail]:
    """Harvest a folder: recency window first, then optional keyword Restrict."""
    results: List[ScrapedMail] = []

    def _keep(item, *, ignore_since: bool = False) -> Optional[ScrapedMail]:
        mail = _mail_from_item(
            item,
            store_id=store_id,
            account_smtp=account_smtp,
            cfg=cfg,
            since=since,
            unread_only=unread_only,
            include_body=True,
            folder=folder,
            body_limit=body_limit,
            ignore_since=ignore_since,
        )
        if mail is None or mail.entry_id in seen_entry_ids:
            return None
        seen_entry_ids.add(mail.entry_id)
        return mail

    for item in _iter_inbox_since(items, since, max_take=max_take):
        if len(results) >= max_take:
            break
        mail = _keep(item)
        if mail is not None:
            results.append(mail)

    if keyword_filter and len(results) < max_take:
        try:
            flagged = items.Restrict(keyword_filter)
            try:
                flagged.Sort("[ReceivedTime]", True)
            except Exception:
                pass
            flagged_count = int(flagged.Count)
        except Exception as exc:
            logger.debug("Keyword Restrict failed for %s/%s: %s", account_smtp, folder, exc)
            flagged = None
            flagged_count = 0
        extra_cap = min(max_take, 80)
        extra_taken = 0
        for idx in range(1, (flagged_count if flagged is not None else 0) + 1):
            if extra_taken >= extra_cap or len(results) >= max_take:
                break
            try:
                item = flagged.Item(idx)
            except Exception:
                continue
            mail = _keep(item)
            if mail is None:
                continue
            results.append(mail)
            extra_taken += 1
        if extra_taken:
            logger.info(
                "Keyword pass: +%d message(s) from %s/%s",
                extra_taken,
                account_smtp,
                folder,
            )
    return results


def _walk_conversation_items(
    conv,
    *,
    store_id: str,
    account_smtp: str,
    cfg: AesAccountScanConfig,
    since: datetime,
    folder: str,
    body_limit: int,
    max_items: int,
    seen_entry_ids: set,
) -> List[ScrapedMail]:
    """Walk Outlook Conversation.GetRootItems / GetChildren (capped)."""
    collected: List[ScrapedMail] = []

    def walk(simple_items, depth: int = 0) -> None:
        if simple_items is None or depth > 8 or len(collected) >= max_items:
            return
        try:
            count = int(simple_items.Count)
        except Exception:
            return
        for i in range(1, count + 1):
            if len(collected) >= max_items:
                return
            try:
                item = simple_items.Item(i)
            except Exception:
                continue
            mail = _mail_from_item(
                item,
                store_id=store_id,
                account_smtp=account_smtp,
                cfg=cfg,
                since=since,
                unread_only=False,
                include_body=True,
                folder=folder,
                body_limit=body_limit,
                ignore_since=True,
            )
            if mail is not None and mail.entry_id not in seen_entry_ids:
                seen_entry_ids.add(mail.entry_id)
                collected.append(mail)
            try:
                children = conv.GetChildren(item)
            except Exception:
                children = None
            if children is not None:
                walk(children, depth + 1)

    try:
        roots = conv.GetRootItems()
    except Exception:
        roots = None
    walk(roots)
    return collected


def expand_conversations(
    seeds: List[ScrapedMail],
    *,
    max_per_conversation: int = 25,
    max_conversations: int = 20,
    progress: Optional[Any] = None,
) -> List[ScrapedMail]:
    """Pull remaining thread members for seed messages via GetConversation."""
    if not seeds:
        return []
    with _OutlookComScope():
        return _expand_conversations_impl(
            seeds,
            max_per_conversation=max_per_conversation,
            max_conversations=max_conversations,
            progress=progress,
        )


def _expand_conversations_impl(
    seeds: List[ScrapedMail],
    *,
    max_per_conversation: int = 25,
    max_conversations: int = 20,
    progress: Optional[Any] = None,
) -> List[ScrapedMail]:
    try:
        app = _outlook_app()
        ns = app.GetNamespace("MAPI")
    except Exception as exc:
        raise RuntimeError(f"Outlook is not available: {exc}") from exc

    extra: List[ScrapedMail] = []
    seen_conv: set = set()
    seen_ids: set = {m.entry_id for m in seeds if m.entry_id}
    since = datetime.now() - timedelta(days=3650)

    for mail in seeds[: max(1, int(max_conversations))]:
        key = (mail.conversation_id or mail.conversation or mail.entry_id).lower()
        if not key or key in seen_conv:
            continue
        seen_conv.add(key)
        if not mail.entry_id:
            continue
        _emit_progress(
            progress,
            f"Expanding thread: {(mail.conversation or mail.subject or '')[:60]}",
        )
        try:
            if mail.store_id:
                item = ns.GetItemFromID(mail.entry_id, mail.store_id)
            else:
                item = ns.GetItemFromID(mail.entry_id)
        except Exception:
            continue
        try:
            conv = item.GetConversation()
        except Exception:
            conv = None
        if conv is None:
            continue
        cfg = AesAccountScanConfig(
            store_id=mail.store_id,
            smtp=mail.account_smtp,
            in_cc=True,
        )
        try:
            walked = _walk_conversation_items(
                conv,
                store_id=mail.store_id,
                account_smtp=mail.account_smtp,
                cfg=cfg,
                since=since,
                folder=mail.folder or "Inbox",
                body_limit=2800,
                max_items=max(1, int(max_per_conversation)),
                seen_entry_ids=seen_ids,
            )
        except Exception as exc:
            logger.debug("Conversation expand failed: %s", exc)
            continue
        extra.extend(walked)
    return extra


def _scrape_inbox_for_search_impl(
    *,
    store_id: str = "",
    smtp: str = "",
    days: int = 30,
    max_messages: int = 150,
    include_sent: bool = True,
    unread_only: bool = False,
    keywords: Optional[List[str]] = None,
    progress: Optional[Any] = None,
) -> List[ScrapedMail]:
    try:
        app = _outlook_app()
        ns = app.GetNamespace("MAPI")
    except Exception as exc:
        raise RuntimeError(f"Outlook is not available: {exc}") from exc

    found = _find_store(ns, store_id=store_id, smtp=smtp)
    if found is None:
        raise RuntimeError(
            "Could not find that Outlook mailbox. Open Outlook and pick the "
            "account again on the Inbox Search tab."
        )
    store, resolved_id = found
    account_smtp = smtp or _account_smtp_for_store(ns, store)
    saved = {c.store_id: c for c in load_aes_account_configs()}
    cfg = saved.get(resolved_id) or AesAccountScanConfig(
        store_id=resolved_id,
        smtp=account_smtp,
        enabled=True,
        in_cc=True,
    )
    # Search should include Cc so conversational context is not dropped
    cfg = AesAccountScanConfig(
        store_id=cfg.store_id,
        smtp=cfg.smtp or account_smtp,
        display=cfg.display,
        enabled=True,
        responses=True,
        in_cc=True,
    )

    since = datetime.now() - timedelta(days=max(1, int(days)))
    max_take = max(10, min(400, int(max_messages)))
    keyword_filter = _outlook_keyword_subject_filter(since, keywords)
    seen_entry_ids: set = set()
    results: List[ScrapedMail] = []

    folders: List[Tuple[int, str]] = [(OL_FOLDER_INBOX, "Inbox")]
    if include_sent:
        folders.append((OL_FOLDER_SENTMAIL, "Sent"))

    per_folder = max_take if not include_sent else max(20, max_take // 2 + max_take // 4)

    for ol_id, folder_name in folders:
        _emit_progress(progress, f"Scanning {folder_name} ({account_smtp or folder_name})…")
        try:
            folder = store.GetDefaultFolder(ol_id)
            items = folder.Items
        except Exception as exc:
            logger.warning("Skip folder %s: %s", folder_name, exc)
            continue
        batch = _scrape_folder_items(
            items,
            store_id=resolved_id,
            account_smtp=account_smtp,
            cfg=cfg,
            since=since,
            unread_only=unread_only,
            folder=folder_name,
            max_take=per_folder,
            seen_entry_ids=seen_entry_ids,
            body_limit=2800,
            keyword_filter=keyword_filter,
        )
        results.extend(batch)

    results.sort(key=lambda m: m.received_dt or datetime.min, reverse=True)
    if len(results) > max_take:
        results = results[:max_take]
    logger.info(
        "Inbox search scrape: %d messages from %s (days=%s, sent=%s)",
        len(results),
        account_smtp or resolved_id[:20],
        days,
        include_sent,
    )
    _emit_progress(progress, f"Scraped {len(results)} message(s)")
    return results


def fetch_mail_body(entry_id: str, store_id: str = "", *, limit: int = 20000) -> str:
    """Load plain-text body for one message (on-demand preview — avoids bulk Body reads)."""
    if not entry_id:
        return ""
    with _OutlookComScope():
        try:
            app = _outlook_app()
            ns = app.GetNamespace("MAPI")
            if store_id:
                item = ns.GetItemFromID(entry_id, store_id)
            else:
                item = ns.GetItemFromID(entry_id)
        except Exception as exc:
            logger.debug("fetch_mail_body open failed: %s", exc)
            return ""
        body = ""
        try:
            body = _safe_str(
                item.PropertyAccessor.GetProperty(
                    "urn:schemas:httpmail:textdescription"
                ),
                limit=limit,
            )
        except Exception:
            body = ""
        if not body:
            try:
                body = _safe_str(getattr(item, "Body", ""), limit=limit)
            except Exception:
                body = ""
        return body
