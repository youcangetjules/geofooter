#!/usr/bin/env python3
"""AES action protocol handler.

Handles aes:// links clicked in AES footers / reports, e.g.
    aes://block-attachments?sender=foo@bar.com&domain=bar.com&guri=...
    aes://block-beacons?sender=...&domain=...
    aes://trust-sender?sender=...&domain=...
    aes://untrust-sender?sender=...&domain=...
    aes://broker-removal?sender=...&domain=...&broker=spokeo&subject=...
    aes://restore-html?id=r20260923120000...
    aes://open-report?path=C:/GeoFooter/output/links/links_report_….html
    aes://set-link-risk?report=links_report_…&id=…&level=low|medium|high|clear

Rules are stored in %LOCALAPPDATA%\\GeoFooter\\aes_sender_rules.json and are
consumed by the Outlook VBA engine on the next scan of mail from that sender.

restore-html reloads the HTML backup saved when a message was force-converted
to text-only (score > 70 mitigation).

Run once with --register to create the HKCU aes:// protocol registration.
"""

from __future__ import annotations

import json
import logging
import os
import sys
import urllib.parse
from datetime import datetime, timezone
from pathlib import Path
from typing import Any, Dict, Optional

# Launched by path from the aes:// registration: put the install root (not aes\) on sys.path.
_PKG_DIR = Path(__file__).resolve().parent
sys.path[:] = [p for p in sys.path if Path(p or ".").resolve() != _PKG_DIR]
if str(_PKG_DIR.parent) not in sys.path:
    sys.path.insert(0, str(_PKG_DIR.parent))


def _action_log_path() -> Path:
    try:
        from geofooter.paths import debug_log_path

        return debug_log_path("aes_action_handler.log")
    except Exception:
        return Path("debuglog") / "aes_action_handler.log"


LOG_PATH = _action_log_path()
try:
    from geofooter.paths import user_data_dir as _user_data

    RULES_PATH = _user_data() / "aes_sender_rules.json"
except Exception:
    RULES_PATH = Path(os.environ.get("LOCALAPPDATA", "")) / "GeoFooter" / "aes_sender_rules.json"
_REPORT_ROOTS = []
try:
    from geofooter.paths import get_install_root, user_data_dir

    _REPORT_ROOTS = [
        get_install_root() / "output",
        user_data_dir() / "output",
    ]
except Exception:
    _REPORT_ROOTS = [Path("output")]

MITIGATED_DIRS = []
try:
    from geofooter.paths import get_install_root, user_data_dir

    MITIGATED_DIRS = [
        user_data_dir() / "mitigated_html",
        get_install_root() / "mitigated_html",
        get_install_root() / "VBA" / "mitigated_html",
    ]
except Exception:
    MITIGATED_DIRS = [Path("mitigated_html")]

ACTIONS = {
    "block-attachments": ("block_attachments", "Attachments from {who} will be quarantined by AES."),
    "block-beacons": ("block_beacons", "Tracking beacons in mail from {who} will be neutralised by AES."),
    "trust-sender": ("trusted", "{who} is now on the AES trusted-sender list."),
    "untrust-sender": ("untrusted", "{who} is marked not trusted in AES."),
}


def _identities(sender: str, domain: str) -> list[str]:
    """Email and/or domain forms that must stay consistent across rule lists."""
    out: list[str] = []
    for value in (sender, domain):
        v = (value or "").strip().lower()
        if not v or v == "unknown":
            continue
        if v not in out:
            out.append(v)
    return out


def _purge_identities(entries: list, identities: list[str]) -> list:
    doomed = set(identities)
    return [e for e in entries if str(e).strip().lower() not in doomed]


def _ensure_on_list(entries: list, who: str) -> list:
    who = (who or "").strip().lower()
    if not who:
        return list(entries)
    cleaned = [e for e in entries if str(e).strip().lower() != who]
    cleaned.append(who)
    return cleaned



def _setup_logging() -> logging.Logger:
    logging.basicConfig(
        filename=str(LOG_PATH),
        level=logging.INFO,
        format="%(asctime)s | %(levelname)s | %(message)s",
    )
    return logging.getLogger("aes_action_handler")


def load_rules() -> dict:
    empty = {
        "block_attachments": [],
        "block_beacons": [],
        "allow_beacons": [],
        "full_no_trust": [],
        "trusted": [],
        "untrusted": [],
    }
    try:
        if RULES_PATH.is_file():
            data = json.loads(RULES_PATH.read_text(encoding="utf-8"))
            if isinstance(data, dict):
                for key in empty:
                    if not isinstance(data.get(key), list):
                        data[key] = []
                return data
    except Exception:
        pass
    return empty


def _global_beacon_blocks(domain: str) -> bool:
    """Default from AES Settings → Beacon Blocking. Footer buttons override this."""
    path = RULES_PATH.parent / "aes_beacon_blocking.json"
    try:
        if not path.is_file():
            return False
        data = json.loads(path.read_text(encoding="utf-8"))
    except Exception:
        return False
    if not isinstance(data, dict) or not data.get("enabled"):
        return False
    dom = (domain or "").strip().lower()
    whitelist = {
        str(item).strip().lower()
        for item in (data.get("whitelist_domains") or [])
        if str(item).strip()
    }
    if dom and dom in whitelist:
        return False
    return bool(data.get("block_all", True))


def _listed(rules: dict, key: str, identities: list[str]) -> bool:
    entries = [str(e).strip().lower() for e in (rules.get(key) or [])]
    return any(ident in entries for ident in identities)


def _beacons_blocked_now(rules: dict, identities: list[str], domain: str) -> bool:
    """Effective beacon block: button override, else the global default."""
    if _listed(rules, "trusted", identities):
        return False
    # FNT and an explicit block win. An allow override beats NT's beacon block.
    if _listed(rules, "full_no_trust", identities) or _listed(rules, "block_beacons", identities):
        return True
    if _listed(rules, "allow_beacons", identities):
        return False
    if _listed(rules, "untrusted", identities):
        return True
    return _global_beacon_blocks(domain)


def _nt_level(rules: dict, identities: list[str]) -> int:
    """0 neutral, 1 NT (trust off, beacons blocked), 2 FNT (text-only)."""
    if _listed(rules, "trusted", identities):
        return 0
    if _listed(rules, "full_no_trust", identities):
        return 2
    if _listed(rules, "untrusted", identities):
        return 1
    return 0


def save_rules(rules: dict) -> None:
    for key in (
        "block_attachments",
        "block_beacons",
        "allow_beacons",
        "full_no_trust",
        "trusted",
        "untrusted",
    ):
        if not isinstance(rules.get(key), list):
            rules[key] = []
    rules["updated"] = datetime.now(timezone.utc).strftime("%Y-%m-%d %H:%M:%S UTC")
    RULES_PATH.parent.mkdir(parents=True, exist_ok=True)
    tmp = RULES_PATH.with_suffix(".json.tmp")
    tmp.write_text(json.dumps(rules, indent=2), encoding="utf-8")
    os.replace(tmp, RULES_PATH)


def show_message(title: str, body: str, error: bool = False) -> None:
    try:
        import tkinter as tk
        from tkinter import messagebox

        root = tk.Tk()
        root.withdraw()
        root.attributes("-topmost", True)
        if error:
            messagebox.showerror(title, body, parent=root)
        else:
            messagebox.showinfo(title, body, parent=root)
        root.destroy()
    except Exception:
        pass


def _find_restore_meta(restore_id: str) -> Optional[Dict[str, Any]]:
    rid = (restore_id or "").strip()
    if not rid or ".." in rid or "/" in rid or "\\" in rid:
        return None
    for base in MITIGATED_DIRS:
        meta_path = base / f"{rid}.json"
        if not meta_path.is_file():
            continue
        try:
            data = json.loads(meta_path.read_text(encoding="utf-8"))
        except Exception:
            continue
        if isinstance(data, dict):
            data.setdefault("html_path", str(base / f"{rid}.html"))
            return data
    return None


def handle_restore_html(params: dict, logger: logging.Logger) -> int:
    """Restore a text-only mitigated message to its saved HTML body via Outlook COM."""
    restore_id = (params.get("id") or [""])[0].strip()
    meta = _find_restore_meta(restore_id)
    if not meta:
        logger.error("restore-html: unknown id=%s", restore_id)
        show_message(
            "AES",
            f"No HTML backup found for id {restore_id or '(missing)'}.",
            error=True,
        )
        return 1

    html_path = Path(str(meta.get("html_path") or ""))
    if not html_path.is_file():
        logger.error("restore-html: missing file %s", html_path)
        show_message("AES", f"HTML backup file missing:\n{html_path}", error=True)
        return 1

    entry_id = str(meta.get("entry_id") or "").strip()
    store_id = str(meta.get("store_id") or "").strip()
    if not entry_id:
        logger.error("restore-html: no entry_id in meta for %s", restore_id)
        show_message("AES", "Backup is missing the Outlook EntryID.", error=True)
        return 1

    try:
        html_body = html_path.read_text(encoding="utf-8", errors="replace")
    except Exception as exc:
        logger.error("restore-html: read failed: %s", exc)
        show_message("AES", f"Could not read HTML backup:\n{exc}", error=True)
        return 1

    try:
        import win32com.client  # type: ignore

        outlook = win32com.client.Dispatch("Outlook.Application")
        ns = outlook.GetNamespace("MAPI")
        if store_id:
            item = ns.GetItemFromID(entry_id, store_id)
        else:
            item = ns.GetItemFromID(entry_id)
        # olFormatHTML = 2
        item.BodyFormat = 2
        item.HTMLBody = html_body
        item.Save()
        logger.info(
            "restore-html: restored id=%s subject=%s",
            restore_id,
            meta.get("subject"),
        )
        show_message(
            "Aliniant Email Scanner",
            "Original HTML format restored for this message.\n"
            "Attachments that were quarantined are not put back automatically.",
        )
        return 0
    except Exception as exc:
        logger.exception("restore-html: Outlook COM failed")
        show_message(
            "AES",
            "Could not restore HTML via Outlook.\n"
            f"{exc}\n\nBackup file:\n{html_path}",
            error=True,
        )
        return 1


def handle_broker_removal(params: dict, logger: logging.Logger) -> int:
    """Queue a data-broker removal item for the GURI Data Brokers tab (no send)."""
    try:
        from aura.pending import enqueue_from_aes
    except Exception as exc:
        logger.exception("broker-removal: import failed")
        show_message("AES", f"Could not load broker-removal module:\n{exc}", error=True)
        return 1

    sender = (params.get("sender") or [""])[0].strip().lower()
    domain = (params.get("domain") or [""])[0].strip().lower()
    guri = (params.get("guri") or [""])[0].strip()
    broker_id = (params.get("broker") or [""])[0].strip()
    subject = (params.get("subject") or [""])[0].strip()

    if not sender and not domain:
        show_message("AES", "Broker-removal link is missing sender details.", error=True)
        return 1

    try:
        item = enqueue_from_aes(
            sender=sender,
            domain=domain,
            subject=subject,
            guri=guri,
            broker_id=broker_id,
        )
    except Exception as exc:
        logger.exception("broker-removal: enqueue failed")
        show_message("AES", f"Could not queue removal request:\n{exc}", error=True)
        return 1
    if item is None:
        show_message("AES", "This message is outside the Aura date window; not queued.")
        return 0

    logger.info(
        "broker-removal queued id=%s broker=%s sender=%s domain=%s",
        item.get("id"),
        item.get("broker_name"),
        sender,
        domain,
    )
    show_message(
        "Aliniant Email Scanner",
        "Queued for data-broker removal:\n"
        f"  {item.get('broker_name')}\n"
        f"  {sender or domain}\n\n"
        "Open GURI → Data Brokers → AES inbox to Accept (creates a draft request)\n"
        "or Dismiss. Templates are filled manually — nothing is sent automatically.",
    )
    return 0


def _safe_report_path(raw: str) -> Optional[Path]:
    """Resolve a local report path; only allow under known GeoFooter output trees."""
    text = (raw or "").strip().strip('"')
    if not text:
        return None
    if text.lower().startswith("file:"):
        text = urllib.parse.urlparse(text).path
        if text.startswith("/") and len(text) > 2 and text[2] == ":":
            text = text[1:]  # /C:/... → C:/...
        text = urllib.parse.unquote(text)
    try:
        path = Path(text).resolve()
    except Exception:
        return None
    if not path.is_file():
        return None
    for root in _REPORT_ROOTS:
        try:
            root_res = root.resolve()
        except Exception:
            continue
        if not root_res.exists():
            continue
        try:
            path.relative_to(root_res)
            return path
        except ValueError:
            continue
    return None


def _links_pair(stem: str) -> Optional[tuple]:
    """Return (html path, sidecar path) for a links report under a known output root."""
    from aes.link_ratings import valid_report_stem

    if not valid_report_stem(stem):
        return None
    for root in _REPORT_ROOTS:
        try:
            root_res = root.resolve()
        except Exception:
            continue
        html_path = (root_res / "links" / f"{stem}.html").resolve()
        side_path = (root_res / "links" / f"{stem}.json").resolve()
        try:
            html_path.relative_to(root_res)
            side_path.relative_to(root_res)
        except ValueError:
            continue
        if html_path.is_file() and side_path.is_file():
            return html_path, side_path
    return None


def handle_set_link_risk(params: dict, logger: logging.Logger) -> int:
    """Save a manual link rating and refresh that links report."""
    from aes.link_ratings import (
        finding_url,
        rewrite_report,
        save_rating,
        valid_link_id,
    )

    stem = (params.get("report") or [""])[0].strip()
    item_id = (params.get("id") or [""])[0].strip().lower()
    level = (params.get("level") or [""])[0].strip().lower()
    if level not in {"low", "medium", "high", "clear"} or not valid_link_id(item_id):
        logger.error("set-link-risk: bad args report=%s id=%s level=%s", stem, item_id, level)
        show_message("AES", "That link-rating action is not valid.", error=True)
        return 1
    pair = _links_pair(stem)
    if pair is None:
        logger.error("set-link-risk: report not found stem=%s", stem)
        show_message(
            "AES",
            "That links page has no saved link list. Rescan the message and open Show links again.",
            error=True,
        )
        return 1
    html_path, side_path = pair
    try:
        sidecar = json.loads(side_path.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError) as exc:
        logger.error("set-link-risk: sidecar unreadable %s (%s)", side_path, exc)
        show_message("AES", "Could not read the saved link list for that page.", error=True)
        return 1
    url = finding_url(sidecar if isinstance(sidecar, dict) else {}, item_id)
    if not url:
        logger.error("set-link-risk: id not in report id=%s", item_id)
        show_message("AES", "That row is not on this links page.", error=True)
        return 1
    try:
        save_rating(url, level)
        rewrite_report(html_path, side_path)
    except Exception as exc:
        logger.exception("set-link-risk: save failed")
        show_message("AES", f"Could not save that rating.\n\n{exc}", error=True)
        return 1
    logger.info("set-link-risk: %s -> %s report=%s", item_id, level, stem)
    if level == "clear":
        blurb = "Cleared. Later scans use the scanner's rating again."
    else:
        blurb = f"Saved as {level.upper()}. Later scans of this URL keep that rating."
    show_message("AES", blurb + "\n\nRefresh the links page.")
    return 0


def handle_open_report(params: dict, logger: logging.Logger) -> int:
    """Open a local AES HTML report (attachments / beacons / links / full scan)."""
    raw = (params.get("path") or [""])[0]
    path = _safe_report_path(raw)
    if path is None:
        logger.error("open-report: refused or missing path=%s", raw)
        show_message(
            "AES",
            "Could not open that report (missing file or path not under GeoFooter output).",
            error=True,
        )
        return 1
    try:
        os.startfile(str(path))  # type: ignore[attr-defined]
        logger.info("open-report: opened %s", path)
        return 0
    except Exception as exc:
        logger.exception("open-report: startfile failed")
        show_message("AES", f"Could not open report:\n{path}\n\n{exc}", error=True)
        return 1


def handle_url(url: str, logger: logging.Logger) -> int:
    parsed = urllib.parse.urlparse(url)
    if parsed.scheme.lower() != "aes":
        logger.error("Not an aes:// URL: %s", url)
        show_message("AES", f"Unrecognised link: {url}", error=True)
        return 1

    # aes://block-attachments?... puts the action in netloc; aes:block-... in path.
    action = (parsed.netloc or parsed.path.lstrip("/")).strip().lower()
    params = urllib.parse.parse_qs(parsed.query)

    if action == "restore-html":
        return handle_restore_html(params, logger)

    if action == "broker-removal":
        return handle_broker_removal(params, logger)

    if action == "open-report":
        return handle_open_report(params, logger)

    if action == "set-link-risk":
        return handle_set_link_risk(params, logger)

    sender = (params.get("sender") or [""])[0].strip().lower()
    domain = (params.get("domain") or [""])[0].strip().lower()
    guri = (params.get("guri") or [""])[0].strip()

    who = sender or domain
    if action not in ACTIONS or not who:
        logger.error("Bad action/target: action=%s sender=%s domain=%s", action, sender, domain)
        show_message("AES", "This AES action link is missing its sender details.", error=True)
        return 1

    list_key, blurb = ACTIONS[action]
    rules = load_rules()
    identities = _identities(sender, domain)
    who = identities[0] if identities else ""

    already = who in [str(e).strip().lower() for e in rules[list_key]]

    # Keep trust / untrust / blocks mutually consistent.
    # Trust clears BOTH the email and the domain from every block list — otherwise
    # a domain-level beacon block survives after trusting the mailbox address
    # (attachments appeared to "unblock" while Beacons BLOCKED stayed on).
    beacons_on = None
    convert_fnt = False
    replace_footer = True
    if list_key == "trusted":
        rules["trusted"] = _ensure_on_list(rules["trusted"], who)
        rules["untrusted"] = _purge_identities(rules["untrusted"], identities)
        rules["block_attachments"] = _purge_identities(rules["block_attachments"], identities)
        rules["block_beacons"] = _purge_identities(rules["block_beacons"], identities)
        rules["allow_beacons"] = _purge_identities(rules["allow_beacons"], identities)
        rules["full_no_trust"] = _purge_identities(rules["full_no_trust"], identities)
    elif action == "untrust-sender":
        # Three states: neutral NT, first click (trust off + beacons blocked),
        # second click FNT (attachments blocked, this mail becomes text).
        nt_level = _nt_level(rules, identities)
        rules["trusted"] = _purge_identities(rules["trusted"], identities)
        rules["allow_beacons"] = _purge_identities(rules["allow_beacons"], identities)
        if nt_level >= 2:
            replace_footer = False
            blurb = (
                "{who} is already Full No Trust. "
                "The message stays text until you click the restore link."
            )
        elif nt_level == 1:
            convert_fnt = True
            replace_footer = False
            rules["untrusted"] = _ensure_on_list(rules["untrusted"], who)
            rules["full_no_trust"] = _ensure_on_list(rules["full_no_trust"], who)
            rules["block_beacons"] = _ensure_on_list(rules["block_beacons"], who)
            rules["block_attachments"] = _ensure_on_list(rules["block_attachments"], who)
            blurb = (
                "{who} is Full No Trust. Attachments are blocked and this message "
                "was converted to text. The restore link at the bottom does nothing "
                "until you click it."
            )
        else:
            rules["untrusted"] = _ensure_on_list(rules["untrusted"], who)
            rules["full_no_trust"] = _purge_identities(rules["full_no_trust"], identities)
            rules["block_beacons"] = _ensure_on_list(rules["block_beacons"], who)
            blurb = (
                "Trust removed for {who}. Beacons from this sender are blocked. "
                "NT stays the neutral chip. Click NT again for Full No Trust."
            )
    elif action == "block-beacons":
        # Toggle against the effective state so the button overrides the global default.
        beacons_on = not _beacons_blocked_now(rules, identities, domain)
        if beacons_on:
            rules["block_beacons"] = _ensure_on_list(rules["block_beacons"], who)
            rules["allow_beacons"] = _purge_identities(rules["allow_beacons"], identities)
            rules["trusted"] = _purge_identities(rules["trusted"], identities)
            blurb = "Tracking beacons in mail from {who} will be neutralised by AES."
        else:
            rules["allow_beacons"] = _ensure_on_list(rules["allow_beacons"], who)
            rules["block_beacons"] = _purge_identities(rules["block_beacons"], identities)
            blurb = (
                "Beacons from {who} will not be blocked. "
                "This overrides the global beacon rule for this sender."
            )
    else:
        # Blocking attachments/beacons clears trusted (not untrusted).
        rules[list_key] = _ensure_on_list(rules[list_key], who)
        rules["trusted"] = _purge_identities(rules["trusted"], identities)

    save_rules(rules)
    logger.info(
        "Action %s applied to %s identities=%s (guri=%s, already=%s) "
        "block_attachments=%s block_beacons=%s trusted=%s",
        action,
        who,
        identities,
        guri,
        already,
        rules.get("block_attachments"),
        rules.get("block_beacons"),
        rules.get("trusted"),
    )

    refreshed = False
    converted = False
    replaced = False
    locked = False
    if convert_fnt:
        converted = _apply_full_no_trust_text(logger)
    else:
        try:
            _outlook, open_item = _outlook_mail_item()
            locked = _mail_is_locked_text(open_item)
        except Exception:
            locked = False
        if locked:
            logger.info("Leaving text-only mail unchanged after %s", action)
        elif replace_footer:
            # One writer. Painting HTML and scanning at the same time flips the
            # message between text and HTML ("message has been changed").
            replaced = _replace_open_footer(logger)
            if not replaced:
                refreshed = _refresh_open_mail_action_buttons(
                    action=action,
                    sender=sender,
                    domain=domain,
                    logger=logger,
                    beacons_on=beacons_on,
                )
        else:
            refreshed = _refresh_open_mail_action_buttons(
                action=action,
                sender=sender,
                domain=domain,
                logger=logger,
                beacons_on=beacons_on,
            )

    state = "was already set — rule refreshed" if already else "rule saved"
    if converted or locked:
        extra = "\nThis message stays text. The restore link is not opened unless you click it."
    elif replaced:
        extra = "\nThe footer on this message is being replaced."
    elif refreshed:
        extra = "\nFooter buttons on the open message were updated."
    else:
        extra = "\nRescan the message (Short Scan) to refresh the footer."
    show_message(
        "Aliniant Email Scanner",
        f"{blurb.format(who=who)}\n\n({state}.){extra}\n"
        f"Rules file: {RULES_PATH}",
    )
    return 0


def _mail_is_locked_text(item) -> bool:
    """True when this mail was converted to text and must not be put back to HTML."""
    if item is None:
        return False
    try:
        body = str(getattr(item, "Body", "") or "")
    except Exception:
        return False
    upper = body.upper()
    return "AES HIGH RISK MITIGATION" in upper or "AES FULL NO TRUST" in upper


def _outlook_mail_item():
    """Open inspector message, or the explorer selection."""
    import win32com.client  # type: ignore

    outlook = win32com.client.Dispatch("Outlook.Application")
    item = None
    insp = outlook.ActiveInspector
    if insp is not None:
        try:
            item = insp.CurrentItem
        except Exception:
            item = None
    if item is None:
        exp = outlook.ActiveExplorer
        if exp is not None and exp.Selection.Count > 0:
            item = exp.Selection.Item(1)
    return outlook, item


def _replace_open_footer(logger: logging.Logger) -> bool:
    """Run AES Short Scan so the open mail's footer is replaced, not left stale."""
    try:
        outlook, _item = _outlook_mail_item()
        explorer = outlook.ActiveExplorer
        if explorer is None:
            return False
        ctl = explorer.CommandBars.FindControl(Tag="AES_SHORTSCAN")
        if ctl is None:
            logger.warning("Footer replace: AES Short Scan button not found")
            return False
        ctl.Execute()
        logger.info("Footer replace: AES Short Scan started")
        return True
    except Exception as exc:
        logger.warning("Footer replace failed: %s", exc)
        return False


def _save_html_backup(item, html: str) -> str:
    rid = "r" + datetime.now().strftime("%Y%m%d%H%M%S%f")[:17]
    base = RULES_PATH.parent / "mitigated_html"
    base.mkdir(parents=True, exist_ok=True)
    html_path = base / f"{rid}.html"
    html_path.write_text(html, encoding="utf-8")
    store_id = ""
    try:
        store_id = str(item.Parent.StoreID)
    except Exception:
        pass
    meta = {
        "id": rid,
        "entry_id": str(getattr(item, "EntryID", "") or ""),
        "store_id": store_id,
        "html_path": str(html_path),
        "subject": str(getattr(item, "Subject", "") or ""),
    }
    (base / f"{rid}.json").write_text(json.dumps(meta), encoding="utf-8")
    return rid


def _quarantine_attachments(item, logger: logging.Logger) -> int:
    try:
        count = int(item.Attachments.Count)
    except Exception:
        return 0
    if count <= 0:
        return 0
    qdir = RULES_PATH.parent / "Quarantine" / datetime.now().strftime("%Y%m%d_%H%M%S")
    removed = 0
    for i in range(count, 0, -1):
        try:
            att = item.Attachments.Item(i)
            name = str(getattr(att, "FileName", "") or "attachment.bin")
            for bad in '\\/:*?"<>|':
                name = name.replace(bad, "_")
            if not name:
                name = "attachment.bin"
            qdir.mkdir(parents=True, exist_ok=True)
            att.SaveAsFile(str(qdir / name))
            att.Delete()
            removed += 1
        except Exception as exc:
            logger.warning("FNT quarantine skipped an attachment: %s", exc)
    return removed


def _apply_full_no_trust_text(logger: logging.Logger) -> bool:
    """Convert the open mail to text. The restore link is only a link."""
    try:
        _outlook, item = _outlook_mail_item()
        if item is None:
            return False
        plain = str(getattr(item, "Body", "") or "")
        if "AES FULL NO TRUST" in plain.upper():
            logger.info("FNT: message is already text-only")
            return True
        html = str(getattr(item, "HTMLBody", "") or "")
        rid = _save_html_backup(item, html) if html.strip() else ""
        removed = _quarantine_attachments(item, logger)
        try:
            plain = str(item.Body or "")
        except Exception:
            pass
        notice = (
            "AES FULL NO TRUST\r\n"
            "Full No Trust: attachments blocked"
            f" ({removed}) and this message converted to text-only.\r\n"
            "Original HTML is not restored unless you open the link below.\r\n\r\n"
        )
        if rid:
            notice += (
                "Restore original HTML format:\r\n"
                f"aes://restore-html?id={rid}\r\n\r\n"
            )
        notice += "----- Original message -----\r\n"
        # olFormatPlain = 1. Do not open the restore link.
        item.BodyFormat = 1
        item.Body = notice + plain
        item.Save()
        logger.info("FNT: converted open mail to text restore_id=%s quarantined=%s", rid, removed)
        return True
    except Exception as exc:
        logger.exception("FNT text conversion failed")
        show_message("AES", f"Could not convert this message to text:\n{exc}", error=True)
        return False


def _refresh_open_mail_action_buttons(
    *,
    action: str,
    sender: str,
    domain: str,
    logger: logging.Logger,
    beacons_on: bool | None = None,
) -> bool:
    """Rewrite AES action-button labels in the open/selected mail after a rule change.

    Footer HTML is baked in at scan time; without this, Trust looks like it
    failed to clear Beacons BLOCKED even though the rules file was updated.
    """
    try:
        import win32com.client  # type: ignore
        import re as _re

        outlook = win32com.client.Dispatch("Outlook.Application")
        item = None
        insp = outlook.ActiveInspector
        if insp is not None:
            try:
                item = insp.CurrentItem
            except Exception:
                item = None
        if item is None:
            exp = outlook.ActiveExplorer
            if exp is not None and exp.Selection.Count > 0:
                item = exp.Selection.Item(1)
        if item is None:
            return False

        html = str(getattr(item, "HTMLBody", "") or "")
        if "AES actions:" not in html and "aes://trust-sender" not in html.lower():
            return False

        # Only touch mail that looks like it belongs to this sender/domain.
        hay = html.lower()
        if sender and sender.lower() not in hay and domain and domain.lower() not in hay:
            # Links encode sender=; accept either.
            if f"sender={urllib.parse.quote(sender)}" not in hay and (
                not domain or f"domain={urllib.parse.quote(domain)}" not in hay
            ):
                if sender and f"sender={sender.lower()}" not in hay:
                    return False

        new_html = html
        if action == "trust-sender":
            new_html = _re.sub(
                r">Attachments BLOCKED(?:\s*&#10003;|\s*✓)?<",
                ">Block attachments from sender<",
                new_html,
                flags=_re.IGNORECASE,
            )
            new_html = _re.sub(
                r">Beacons BLOCKED(?:\s*&#10003;|\s*✓)?<",
                ">Block beacons from sender<",
                new_html,
                flags=_re.IGNORECASE,
            )
            new_html = _re.sub(
                r">Trust sender(?: \(undoes blocks\))?(?:\s*&#10003;|\s*✓)?<",
                ">Sender Trusted &#10003;<",
                new_html,
                flags=_re.IGNORECASE,
            )
            new_html = _re.sub(
                r">Sender TRUSTED(?:\s*&#10003;|\s*✓)?<",
                ">Sender Trusted &#10003;<",
                new_html,
                flags=_re.IGNORECASE,
            )
            new_html = _paint_action_cell(new_html, "aes://trust-sender", "#2e7d32", "#ffffff")
            new_html = _paint_action_cell(new_html, "aes://block-beacons", "#f2f8fa", "#0f6b7c")
            new_html = _paint_action_cell(new_html, "aes://block-attachments", "#f2f8fa", "#0f6b7c")
            new_html = _re.sub(r">TS<", ">ST<", new_html)
            new_html = _paint_short_chip(new_html, "aes://trust-sender", CHIP_TRUSTED_BG, CHIP_TRUSTED_FG)
            new_html = _paint_short_chip(new_html, "aes://block-beacons", CHIP_BG, CHIP_FG)
            new_html = _paint_short_chip(new_html, "aes://block-attachments", CHIP_BG, CHIP_FG)
        elif action == "block-beacons" and beacons_on is False:
            new_html = _re.sub(
                r">Beacons BLOCKED(?:\s*&#10003;|\s*✓)?<",
                ">Block beacons from sender<",
                new_html,
                flags=_re.IGNORECASE,
            )
            new_html = _paint_action_cell(new_html, "aes://block-beacons", "#f2f8fa", "#0f6b7c")
            new_html = _paint_short_chip(new_html, "aes://block-beacons", CHIP_BG, CHIP_FG)
        elif action == "block-beacons":
            new_html = _re.sub(
                r">Block beacons from sender<",
                ">Beacons BLOCKED &#10003;<",
                new_html,
                flags=_re.IGNORECASE,
            )
            new_html = _paint_action_cell(new_html, "aes://block-beacons", "#b71c1c", "#ffffff")
            new_html = _re.sub(r">ST<", ">TS<", new_html)
            new_html = _paint_short_chip(new_html, "aes://block-beacons", CHIP_BLOCKED_BG, "#ffffff")
            new_html = _paint_short_chip(new_html, "aes://trust-sender", CHIP_BG, CHIP_FG)
            new_html = _re.sub(
                r">Sender Trusted(?:\s*&#10003;|\s*✓)?<",
                ">Trust sender<",
                new_html,
                flags=_re.IGNORECASE,
            )
            new_html = _paint_action_cell(new_html, "aes://trust-sender", "#f2f8fa", "#0f6b7c")
        elif action == "block-attachments":
            new_html = _re.sub(
                r">Block attachments from sender<",
                ">Attachments BLOCKED &#10003;<",
                new_html,
                flags=_re.IGNORECASE,
            )
            new_html = _paint_action_cell(new_html, "aes://block-attachments", "#b71c1c", "#ffffff")
            new_html = _re.sub(r">ST<", ">TS<", new_html)
            new_html = _paint_short_chip(new_html, "aes://block-attachments", CHIP_BLOCKED_BG, "#ffffff")
            new_html = _paint_short_chip(new_html, "aes://trust-sender", CHIP_BG, CHIP_FG)
            new_html = _re.sub(
                r">Sender Trusted(?:\s*&#10003;|\s*✓)?<",
                ">Trust sender<",
                new_html,
                flags=_re.IGNORECASE,
            )
            new_html = _paint_action_cell(new_html, "aes://trust-sender", "#f2f8fa", "#0f6b7c")
        elif action == "untrust-sender":
            new_html = _re.sub(
                r">Mark not trusted<",
                ">Sender Not Trusted &#10003;<",
                new_html,
                flags=_re.IGNORECASE,
            )
            new_html = _re.sub(
                r">Sender NOT TRUSTED(?:\s*&#10003;|\s*✓)?<",
                ">Sender Not Trusted &#10003;<",
                new_html,
                flags=_re.IGNORECASE,
            )
            new_html = _re.sub(
                r">Sender Trusted(?:\s*&#10003;|\s*✓)?<",
                ">Trust sender<",
                new_html,
                flags=_re.IGNORECASE,
            )
            new_html = _paint_action_cell(new_html, "aes://untrust-sender", "#f2f8fa", "#0f6b7c")
            new_html = _paint_action_cell(new_html, "aes://trust-sender", "#f2f8fa", "#0f6b7c")
            new_html = _paint_action_cell(new_html, "aes://block-beacons", "#b71c1c", "#ffffff")
            new_html = _re.sub(r">ST<", ">TS<", new_html)
            new_html = _paint_short_chip(new_html, "aes://trust-sender", CHIP_BG, CHIP_FG)
            # NT stays the neutral chip. Beacons blocked by this click turn BB red.
            new_html = _paint_short_chip(new_html, "aes://untrust-sender", CHIP_BG, CHIP_FG)
            new_html = _paint_short_chip(new_html, "aes://block-beacons", CHIP_BLOCKED_BG, "#ffffff")

        if new_html == html:
            return False
        item.HTMLBody = new_html
        item.Save()
        logger.info("Refreshed AES action buttons on open mail after %s", action)
        return True
    except Exception as exc:
        logger.warning("Could not refresh open-mail action buttons: %s", exc)
        return False


# Quick Action chip colours; mirror AES_CHIP_* in geolocate_headers.py.
CHIP_BG = "#ffffff"
CHIP_FG = "#0f1f2a"
CHIP_BLOCKED_BG = "#d92d20"
CHIP_TRUSTED_BG = "#5dce8a"
CHIP_TRUSTED_FG = "#0f1f2a"


def _paint_action_cell(html: str, href_prefix: str, bg: str, fg: str) -> str:
    """Restyle the nested <td> wrapping an aes:// action link."""
    import re as _re

    pattern = _re.compile(
        r"(<td[^>]*bgcolor='[^']*'[^>]*style='background:[^;']+; border:1px solid [^;']+;"
        r"[^']*'>\s*<a href='" + _re.escape(href_prefix) + r"[^']*'\s*style='color:[^;']+;)",
        _re.IGNORECASE,
    )

    def repl(match) -> str:
        chunk = match.group(1)
        chunk = _re.sub(r"bgcolor='[^']*'", f"bgcolor='{bg}'", chunk, count=1)
        chunk = _re.sub(r"background:[^;']+", f"background:{bg}", chunk, count=1)
        chunk = _re.sub(r"border:1px solid [^;']+", f"border:1px solid {bg if bg != '#f2f8fa' else '#0f6b7c'}", chunk, count=1)
        chunk = _re.sub(r"style='color:[^;']+", f"style='color:{fg}", chunk, count=1)
        return chunk

    return pattern.sub(repl, html, count=1)


def _paint_short_chip(html: str, href_prefix: str, bg: str, fg: str) -> str:
    """Restyle a two-letter Quick Action chip (<SL>, <BA>, <BB>, <TS>, <ST>, <NT>)."""
    import re as _re

    border = fg if bg == "#ffffff" else bg
    pattern = _re.compile(
        r"(<td[^>]*bgcolor=')([^']*)('[^>]*style='background:)([^;']+)(; border:1px solid )([^;']+)"
        r"(;[^']*'[^>]*>)"
        r"(\s*<!--\[if mso\]>.*?<!\[endif\]-->\s*)?"
        r"(<!--\[if !mso\]><!-->\s*)?"
        r"(<a href='" + _re.escape(href_prefix) + r"[^']*' style='color:)"
        r"([^;']+)(;[^']*'>)([A-Za-z]{2})(</a>)",
        _re.IGNORECASE | _re.DOTALL,
    )

    def repl(match) -> str:
        mso = match.group(8) or ""
        if mso:
            mso = _re.sub(r"fillcolor='[^']*'", f"fillcolor='{bg}'", mso, count=1)
            mso = _re.sub(r"strokecolor='[^']*'", f"strokecolor='{border}'", mso, count=1)
            mso = _re.sub(r"(<center style='color:)[^;']+", rf"\g<1>{fg}", mso, count=1)
        return (
            f"{match.group(1)}{bg}{match.group(3)}{bg}{match.group(5)}{border}"
            f"{match.group(7)}{mso}{match.group(9) or ''}"
            f"{match.group(10)}{fg}{match.group(12)}{match.group(13)}{match.group(14)}"
        )

    return pattern.sub(repl, html)


def register_protocol(logger: logging.Logger) -> int:
    """Register the aes:// URL protocol for the current user."""
    import winreg

    pythonw = Path(sys.executable).with_name("pythonw.exe")
    if not pythonw.exists():
        pythonw = Path(sys.executable)
    handler = Path(__file__).resolve()
    command = f'"{pythonw}" "{handler}" "%1"'

    root = winreg.CreateKey(winreg.HKEY_CURRENT_USER, r"Software\Classes\aes")
    winreg.SetValueEx(root, None, 0, winreg.REG_SZ, "URL:AES Action Protocol")
    winreg.SetValueEx(root, "URL Protocol", 0, winreg.REG_SZ, "")
    icon = winreg.CreateKey(root, "DefaultIcon")
    winreg.SetValueEx(icon, None, 0, winreg.REG_SZ, f"{pythonw},0")
    cmd = winreg.CreateKey(root, r"shell\open\command")
    winreg.SetValueEx(cmd, None, 0, winreg.REG_SZ, command)
    winreg.CloseKey(cmd)
    winreg.CloseKey(icon)
    winreg.CloseKey(root)
    logger.info("aes:// protocol registered -> %s", command)
    print(f"aes:// protocol registered: {command}")
    return 0


def main() -> int:
    logger = _setup_logging()
    args = sys.argv[1:]
    if not args:
        print(__doc__)
        return 0
    if args[0] == "--register":
        return register_protocol(logger)
    return handle_url(args[0], logger)


if __name__ == "__main__":
    sys.exit(main())
