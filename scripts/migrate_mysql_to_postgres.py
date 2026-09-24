#!/usr/bin/env python3
"""
Migrate GURI data from MySQL (or SQLite) into PostgreSQL.

Usage:
  python migrate_mysql_to_postgres.py
  python migrate_mysql_to_postgres.py --sqlite C:\\GeoFooter\\guri_records.db

Reads source credentials from guri_mysql_config.json (MySQL) unless --sqlite is set.
Writes/uses destination credentials from guri_postgres_config.json.
"""

from __future__ import annotations

import argparse
import json
import os
import sqlite3
import sys
from pathlib import Path
from typing import Any, Dict, Iterable, List, Optional, Sequence, Tuple

BASE = Path(r"C:\GeoFooter")


def _load_json(path: Path) -> Dict[str, Any]:
    with path.open("r", encoding="utf-8") as fh:
        data = json.load(fh)
    if not isinstance(data, dict):
        raise ValueError(f"Invalid JSON object in {path}")
    return data


def _pg_connect(cfg: Dict[str, Any]):
    import psycopg2

    return psycopg2.connect(
        host=cfg["host"],
        port=int(cfg.get("port", 5432)),
        user=cfg["user"],
        password=cfg.get("password") or "",
        dbname=cfg["database"],
    )


def _ensure_pg_database(cfg: Dict[str, Any]) -> None:
    import psycopg2
    from psycopg2 import sql

    dbname = str(cfg["database"])
    admin = dict(cfg)
    admin["database"] = "postgres"
    conn = _pg_connect(admin)
    conn.autocommit = True
    try:
        cur = conn.cursor()
        cur.execute("SELECT 1 FROM pg_database WHERE datname = %s", (dbname,))
        if cur.fetchone() is None:
            cur.execute(sql.SQL("CREATE DATABASE {}").format(sql.Identifier(dbname)))
            print(f"Created PostgreSQL database {dbname}")
        cur.close()
    finally:
        conn.close()


def _init_pg_schema(cfg: Dict[str, Any]) -> None:
    # Reuse application DDL
    from guri import GURIDatabase
    from guri_component_library import GURIComponentLibrary

    db = GURIDatabase(db_type="postgres", pg_config=cfg, base_path=str(BASE))
    GURIComponentLibrary(database=db)
    print("PostgreSQL schema ready")


def _fetch_mysql(cfg: Dict[str, Any], query: str) -> List[Tuple[Any, ...]]:
    import mysql.connector

    conn = mysql.connector.connect(
        host=cfg["host"],
        port=int(cfg.get("port", 3306)),
        user=cfg["user"],
        password=cfg.get("password") or "",
        database=cfg["database"],
    )
    try:
        cur = conn.cursor()
        cur.execute(query)
        rows = list(cur.fetchall())
        cur.close()
        return rows
    finally:
        conn.close()


def _fetch_sqlite(db_path: Path, query: str) -> List[Tuple[Any, ...]]:
    conn = sqlite3.connect(str(db_path))
    try:
        cur = conn.cursor()
        cur.execute(query)
        return list(cur.fetchall())
    finally:
        conn.close()


def _copy_table(
    pg_cfg: Dict[str, Any],
    table: str,
    columns: Sequence[str],
    rows: Iterable[Tuple[Any, ...]],
    conflict: Optional[str] = None,
) -> int:
    cols = ", ".join(columns)
    placeholders = ", ".join(["%s"] * len(columns))
    sql = f"INSERT INTO {table} ({cols}) VALUES ({placeholders})"
    if conflict:
        sql += f" ON CONFLICT {conflict} DO NOTHING"

    conn = _pg_connect(pg_cfg)
    try:
        cur = conn.cursor()
        count = 0
        for row in rows:
            cur.execute(sql, tuple(row))
            count += cur.rowcount if cur.rowcount and cur.rowcount > 0 else 0
        # Prefer counting attempted inserts when DO NOTHING yields 0
        conn.commit()
        cur.close()
        return count
    finally:
        conn.close()


def _reset_serial(pg_cfg: Dict[str, Any], table: str, id_col: str = "id") -> None:
    from psycopg2 import sql

    conn = _pg_connect(pg_cfg)
    try:
        cur = conn.cursor()
        cur.execute(
            sql.SQL(
                "SELECT setval(pg_get_serial_sequence(%s, %s), "
                "COALESCE((SELECT MAX({}) FROM {}), 1), true)"
            ).format(sql.Identifier(id_col), sql.Identifier(table)),
            (table, id_col),
        )
        conn.commit()
        cur.close()
    finally:
        conn.close()


def migrate(
    *,
    mysql_cfg: Optional[Dict[str, Any]] = None,
    sqlite_path: Optional[Path] = None,
    pg_cfg: Dict[str, Any],
    log: Optional[Any] = None,
) -> Dict[str, Any]:
    """Copy GURI tables from MySQL or SQLite into PostgreSQL.

    Returns a report dict with per-table counts and a ``messages`` list.
    ``log`` may be a callable(str) for live progress (defaults to print).
    """
    emit = log if callable(log) else print
    report: Dict[str, Any] = {
        "ok": False,
        "source": "sqlite" if sqlite_path is not None else "mysql",
        "tables": {},
        "messages": [],
        "error": "",
    }

    def note(msg: str) -> None:
        report["messages"].append(msg)
        emit(msg)

    if sqlite_path is None and mysql_cfg is None:
        raise ValueError("Provide mysql_cfg or sqlite_path")

    _ensure_pg_database(pg_cfg)
    note("Ensured PostgreSQL database exists")
    _init_pg_schema(pg_cfg)
    note("PostgreSQL schema ready")

    def fetch(query: str) -> List[Tuple[Any, ...]]:
        if sqlite_path is not None:
            return _fetch_sqlite(sqlite_path, query)
        assert mysql_cfg is not None
        return _fetch_mysql(mysql_cfg, query)

    guri_cols = (
        "id",
        "guri",
        "sender",
        "recipients",
        "subject",
        "datetime",
        "avg_risk",
        "document_type",
        "sensitivity",
        "company_name",
        "created_at",
    )
    # Some older DBs may lack company_name — probe with a safe select list
    try:
        guri_rows = fetch(
            "SELECT id, guri, sender, recipients, subject, datetime, avg_risk, "
            "document_type, sensitivity, company_name, created_at FROM guri_records"
        )
    except Exception:
        raw = fetch(
            "SELECT id, guri, sender, recipients, subject, datetime, avg_risk, "
            "document_type, sensitivity, created_at FROM guri_records"
        )
        guri_rows = [
            (r[0], r[1], r[2], r[3], r[4], r[5], r[6], r[7], r[8], None, r[9])
            for r in raw
        ]

    _copy_table(
        pg_cfg,
        "guri_records",
        guri_cols,
        guri_rows,
        conflict="(guri)",
    )
    report["tables"]["guri_records"] = {"source_rows": len(guri_rows)}
    note(f"guri_records: migrated {len(guri_rows)} row(s) (unique on guri)")
    _reset_serial(pg_cfg, "guri_records")

    try:
        supplier_rows = fetch(
            "SELECT id, company_name, company_domain, supplier_code, contract_code, "
            "contact_person, contact_email, notes, created_at, updated_at "
            "FROM supplier_codes"
        )
    except Exception as exc:
        note(f"supplier_codes skipped: {exc}")
        supplier_rows = []
    if supplier_rows:
        _copy_table(
            pg_cfg,
            "supplier_codes",
            (
                "id",
                "company_name",
                "company_domain",
                "supplier_code",
                "contract_code",
                "contact_person",
                "contact_email",
                "notes",
                "created_at",
                "updated_at",
            ),
            supplier_rows,
            conflict="(company_name)",
        )
        report["tables"]["supplier_codes"] = {"source_rows": len(supplier_rows)}
        note(f"supplier_codes: migrated {len(supplier_rows)} row(s)")
        _reset_serial(pg_cfg, "supplier_codes")

    try:
        comp_rows = fetch(
            "SELECT id, component_index, component_hex, real_value, guri_id, "
            "created_at, updated_at FROM component_library"
        )
    except Exception as exc:
        note(f"component_library skipped: {exc}")
        comp_rows = []
    if comp_rows:
        _copy_table(
            pg_cfg,
            "component_library",
            (
                "id",
                "component_index",
                "component_hex",
                "real_value",
                "guri_id",
                "created_at",
                "updated_at",
            ),
            comp_rows,
            conflict="(component_index, component_hex)",
        )
        report["tables"]["component_library"] = {"source_rows": len(comp_rows)}
        note(f"component_library: migrated {len(comp_rows)} row(s)")
        _reset_serial(pg_cfg, "component_library")

    note("Migration complete.")
    note(f"Point GURI at: {BASE / 'guri_postgres_config.json'}")
    report["ok"] = True
    return report


def load_mysql_config(path: Optional[Path] = None) -> Dict[str, Any]:
    return _load_json(path or (BASE / "guri_mysql_config.json"))


def load_postgres_config(path: Optional[Path] = None) -> Dict[str, Any]:
    return _load_json(path or (BASE / "guri_postgres_config.json"))


def save_postgres_config(cfg: Dict[str, Any], path: Optional[Path] = None) -> Path:
    out = path or (BASE / "guri_postgres_config.json")
    out.write_text(json.dumps(cfg, indent=4) + "\n", encoding="utf-8")
    return out


def main() -> int:
    parser = argparse.ArgumentParser(description="Migrate GURI MySQL/SQLite → PostgreSQL")
    parser.add_argument(
        "--sqlite",
        type=str,
        default="",
        help="Migrate from a SQLite file instead of MySQL",
    )
    parser.add_argument(
        "--write-pg-config-from-mysql",
        action="store_true",
        help="Create guri_postgres_config.json from MySQL host/user/password/database (port 5432)",
    )
    args = parser.parse_args()

    pg_path = BASE / "guri_postgres_config.json"
    mysql_path = BASE / "guri_mysql_config.json"

    mysql_cfg: Optional[Dict[str, Any]] = None
    sqlite_path: Optional[Path] = Path(args.sqlite) if args.sqlite else None

    if sqlite_path is None:
        if not mysql_path.is_file():
            print(f"Missing {mysql_path}. Pass --sqlite path instead.", file=sys.stderr)
            return 1
        mysql_cfg = _load_json(mysql_path)

    if args.write_pg_config_from_mysql or not pg_path.is_file():
        if mysql_cfg is None:
            print("Cannot invent PostgreSQL config without MySQL config.", file=sys.stderr)
            return 1
        pg_cfg = {
            "host": mysql_cfg.get("host", "localhost"),
            "user": mysql_cfg.get("user", "postgres"),
            "password": mysql_cfg.get("password", ""),
            "database": mysql_cfg.get("database", "guri_db"),
            "port": 5432,
        }
        save_postgres_config(pg_cfg, pg_path)
        print(f"Wrote {pg_path} (review user/password for your PostgreSQL server)")
    else:
        pg_cfg = _load_json(pg_path)

    for key in ("host", "user", "password", "database"):
        if key not in pg_cfg:
            print(f"PostgreSQL config missing '{key}'", file=sys.stderr)
            return 1

    try:
        migrate(mysql_cfg=mysql_cfg, sqlite_path=sqlite_path, pg_cfg=pg_cfg)
    except Exception as exc:
        print(f"Migration failed: {exc}", file=sys.stderr)
        return 1
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
