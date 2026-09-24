#!/usr/bin/env python3
"""
GURI Component Library Module
Stores and retrieves translations from GURI component hex values to their real values.
Excludes Component 4 (datetime - decodable) and Component 6 (document type - actual code).

Invariant: for a given component_index, the same real_value ALWAYS maps to the same
component_hex (and vice versa for the canonical mapping). Hex codes are stable tags
in the GURI string — they are not re-randomized per document.
"""

from __future__ import annotations

import logging
import random
from typing import Any, Dict, List, Optional, Tuple


class GURIComponentLibrary:
    """Manages translations from GURI component hex values to real values."""

    # Component indices (0-based) that need mapping
    COMPONENT_INDICES = {
        0: "sender",  # Component 1: Sender (5 hex)
        1: "recipients",  # Component 2: Recipients (5 hex)
        2: "subject",  # Component 3: Subject (8 hex)
        4: "risk_location",  # Component 5: Risk/Location (3 hex)
    }

    # Expected hex length per mappable component
    COMPONENT_HEX_LENGTHS = {
        0: 5,
        1: 5,
        2: 8,
        4: 3,
    }

    # Component indices excluded from mapping
    EXCLUDED_COMPONENTS = {
        3: "datetime",  # Component 4: Decodable timer
        5: "document_type",  # Component 6: Actual code, not random
    }

    def __init__(self, database=None):
        """
        Initialize the component library.

        Args:
            database: GURIDatabase instance for storing/retrieving mappings
        """
        self.logger = logging.getLogger(__name__)
        self.database = database
        if database:
            self._init_library_table()

    def _ph(self) -> str:
        return "%s" if self.database.db_type in {"postgres", "mysql"} else "?"

    def _init_library_table(self):
        """Initialize the component_library table in the database."""
        try:
            conn = self.database._get_connection()
            cursor = conn.cursor()

            if self.database.db_type == "postgres":
                cursor.execute(
                    """
                    CREATE TABLE IF NOT EXISTS component_library (
                        id SERIAL PRIMARY KEY,
                        component_index INT NOT NULL,
                        component_hex VARCHAR(20) NOT NULL,
                        real_value TEXT NOT NULL,
                        guri_id INT,
                        created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                        updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                        UNIQUE (component_index, component_hex)
                    )
                    """
                )
                cursor.execute(
                    "CREATE INDEX IF NOT EXISTS idx_component_index ON component_library (component_index)"
                )
                cursor.execute(
                    "CREATE INDEX IF NOT EXISTS idx_component_hex ON component_library (component_hex)"
                )
                cursor.execute(
                    "CREATE INDEX IF NOT EXISTS idx_guri_id ON component_library (guri_id)"
                )
                # Stable value → hex: one real_value per component slot
                try:
                    cursor.execute(
                        """
                        CREATE UNIQUE INDEX IF NOT EXISTS uq_component_real_value
                        ON component_library (component_index, real_value)
                        """
                    )
                except Exception as idx_exc:
                    self.logger.warning(
                        "Could not create uq_component_real_value (run Repair): %s",
                        idx_exc,
                    )
            elif self.database.db_type == "mysql":
                cursor.execute(
                    """
                    CREATE TABLE IF NOT EXISTS component_library (
                        id INT AUTO_INCREMENT PRIMARY KEY,
                        component_index INT NOT NULL,
                        component_hex VARCHAR(20) NOT NULL,
                        real_value TEXT NOT NULL,
                        guri_id INT,
                        created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                        updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP ON UPDATE CURRENT_TIMESTAMP,
                        INDEX idx_component_index (component_index),
                        INDEX idx_component_hex (component_hex),
                        INDEX idx_guri_id (guri_id),
                        UNIQUE KEY unique_component_hex (component_index, component_hex)
                    ) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4
                    """
                )
                try:
                    cursor.execute(
                        """
                        CREATE UNIQUE INDEX uq_component_real_value
                        ON component_library (component_index, real_value(255))
                        """
                    )
                except Exception:
                    pass  # may already exist or collide until dedupe runs
            else:  # sqlite
                cursor.execute(
                    """
                    CREATE TABLE IF NOT EXISTS component_library (
                        id INTEGER PRIMARY KEY AUTOINCREMENT,
                        component_index INTEGER NOT NULL,
                        component_hex TEXT NOT NULL,
                        real_value TEXT NOT NULL,
                        guri_id INTEGER,
                        created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                        updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                        UNIQUE(component_index, component_hex)
                    )
                    """
                )
                cursor.execute(
                    "CREATE INDEX IF NOT EXISTS idx_component_index ON component_library(component_index)"
                )
                cursor.execute(
                    "CREATE INDEX IF NOT EXISTS idx_component_hex ON component_library(component_hex)"
                )
                cursor.execute(
                    "CREATE INDEX IF NOT EXISTS idx_guri_id ON component_library(guri_id)"
                )
                try:
                    cursor.execute(
                        """
                        CREATE UNIQUE INDEX IF NOT EXISTS uq_component_real_value
                        ON component_library (component_index, real_value)
                        """
                    )
                except Exception as idx_exc:
                    self.logger.warning(
                        "Could not create uq_component_real_value (run Repair): %s",
                        idx_exc,
                    )

            conn.commit()
            cursor.close()
            conn.close()
            self.logger.info("Component library table initialized")

        except Exception as e:
            self.logger.error(f"Error initializing component library table: {e}")
            raise

    @staticmethod
    def _norm_value(real_value: Optional[str]) -> str:
        return str(real_value if real_value is not None else "").strip()

    def get_hex_for_real_value(
        self, component_index: int, real_value: str
    ) -> Optional[str]:
        """Return the canonical component_hex for a real value, if already mapped."""
        if not self.database or component_index not in self.COMPONENT_INDICES:
            return None
        value = self._norm_value(real_value)
        ph = self._ph()
        try:
            conn = self.database._get_connection()
            cursor = conn.cursor()
            cursor.execute(
                f"""
                SELECT component_hex FROM component_library
                WHERE component_index = {ph} AND real_value = {ph}
                ORDER BY id ASC
                LIMIT 1
                """,
                (component_index, value),
            )
            row = cursor.fetchone()
            cursor.close()
            conn.close()
            return str(row[0]) if row else None
        except Exception as e:
            self.logger.error(f"Error looking up hex for value: {e}")
            return None

    def _hex_in_use(self, component_index: int, component_hex: str) -> bool:
        ph = self._ph()
        try:
            conn = self.database._get_connection()
            cursor = conn.cursor()
            cursor.execute(
                f"""
                SELECT 1 FROM component_library
                WHERE component_index = {ph} AND component_hex = {ph}
                LIMIT 1
                """,
                (component_index, component_hex),
            )
            row = cursor.fetchone()
            cursor.close()
            conn.close()
            return bool(row)
        except Exception:
            return False

    def _new_unique_hex(self, component_index: int, length: int) -> str:
        for _ in range(64):
            candidate = "".join(random.choices("0123456789abcdef", k=length))
            if not self._hex_in_use(component_index, candidate):
                return candidate
        # Extremely unlikely fallback
        return "".join(random.choices("0123456789abcdef", k=length))

    def allocate_hex_for_value(
        self,
        component_index: int,
        real_value: str,
        *,
        length: Optional[int] = None,
        guri_id: Optional[int] = None,
    ) -> str:
        """Return stable hex for this value — reuse existing, else allocate once."""
        if component_index not in self.COMPONENT_INDICES:
            raise ValueError(f"Component {component_index} is not mappable")
        value = self._norm_value(real_value)
        length = int(length or self.COMPONENT_HEX_LENGTHS[component_index])

        existing = self.get_hex_for_real_value(component_index, value)
        if existing:
            if guri_id is not None:
                self._touch_guri_id(component_index, existing, guri_id)
            return existing

        component_hex = self._new_unique_hex(component_index, length)
        self.store_component_mapping(component_index, component_hex, value, guri_id)
        # Re-read in case a concurrent insert won the race
        confirmed = self.get_hex_for_real_value(component_index, value)
        return confirmed or component_hex

    def _touch_guri_id(
        self, component_index: int, component_hex: str, guri_id: int
    ) -> None:
        """Update last-seen guri_id on an existing mapping (does not change hex)."""
        if not self.database:
            return
        ph = self._ph()
        try:
            conn = self.database._get_connection()
            cursor = conn.cursor()
            if self.database.db_type == "sqlite":
                cursor.execute(
                    f"""
                    UPDATE component_library
                    SET guri_id = {ph}, updated_at = datetime('now')
                    WHERE component_index = {ph} AND component_hex = {ph}
                    """,
                    (guri_id, component_index, component_hex),
                )
            else:
                cursor.execute(
                    f"""
                    UPDATE component_library
                    SET guri_id = {ph}, updated_at = CURRENT_TIMESTAMP
                    WHERE component_index = {ph} AND component_hex = {ph}
                    """,
                    (guri_id, component_index, component_hex),
                )
            conn.commit()
            cursor.close()
            conn.close()
        except Exception as e:
            self.logger.debug(f"touch guri_id failed: {e}")

    def store_component_mapping(
        self,
        component_index: int,
        component_hex: str,
        real_value: str,
        guri_id: Optional[int] = None,
    ) -> bool:
        """
        Store a mapping from component hex value to real value.

        If this real_value already has a hex for the component slot, that existing
        hex is kept (same value → same GURI component). A mismatched hex argument
        is ignored in favour of the canonical mapping.
        """
        if not self.database:
            self.logger.warning("No database connection for storing component mapping")
            return False

        if component_index not in self.COMPONENT_INDICES:
            if component_index in self.EXCLUDED_COMPONENTS:
                self.logger.debug(
                    f"Component {component_index} ({self.EXCLUDED_COMPONENTS[component_index]}) is excluded from mapping"
                )
            else:
                self.logger.warning(f"Invalid component index: {component_index}")
            return False

        value = self._norm_value(real_value)
        component_hex = str(component_hex or "").strip().lower()
        existing_hex = self.get_hex_for_real_value(component_index, value)
        if existing_hex:
            if existing_hex.lower() != component_hex:
                self.logger.info(
                    "Reusing canonical hex %s for component %s value %r (ignored %s)",
                    existing_hex,
                    component_index,
                    value[:80],
                    component_hex,
                )
            if guri_id is not None:
                self._touch_guri_id(component_index, existing_hex, guri_id)
            return True

        try:
            conn = self.database._get_connection()
            cursor = conn.cursor()
            ph = self._ph()

            if self.database.db_type == "postgres":
                query = f"""
                    INSERT INTO component_library
                    (component_index, component_hex, real_value, guri_id)
                    VALUES ({ph}, {ph}, {ph}, {ph})
                    ON CONFLICT (component_index, component_hex) DO UPDATE SET
                    real_value = EXCLUDED.real_value,
                    guri_id = COALESCE(EXCLUDED.guri_id, component_library.guri_id),
                    updated_at = CURRENT_TIMESTAMP
                """
            elif self.database.db_type == "mysql":
                query = f"""
                    INSERT INTO component_library
                    (component_index, component_hex, real_value, guri_id)
                    VALUES ({ph}, {ph}, {ph}, {ph})
                    ON DUPLICATE KEY UPDATE
                    real_value = VALUES(real_value),
                    guri_id = COALESCE(VALUES(guri_id), guri_id),
                    updated_at = CURRENT_TIMESTAMP
                """
            else:  # sqlite
                query = f"""
                    INSERT OR REPLACE INTO component_library
                    (component_index, component_hex, real_value, guri_id, updated_at)
                    VALUES ({ph}, {ph}, {ph}, {ph}, datetime('now'))
                """

            cursor.execute(query, (component_index, component_hex, value, guri_id))
            conn.commit()
            cursor.close()
            conn.close()

            self.logger.debug(
                f"Stored mapping: Component {component_index} '{component_hex}' -> '{value[:50]}...'"
            )
            return True

        except Exception as e:
            # Unique on real_value may have won a race — treat as success if mapped
            again = self.get_hex_for_real_value(component_index, value)
            if again:
                return True
            self.logger.error(f"Error storing component mapping: {e}")
            return False

    def get_component_value(
        self, component_index: int, component_hex: str
    ) -> Optional[str]:
        """Get the real value for a component hex value."""
        if not self.database:
            return None

        if component_index in self.EXCLUDED_COMPONENTS:
            return None

        try:
            conn = self.database._get_connection()
            cursor = conn.cursor()
            ph = self._ph()

            cursor.execute(
                f"""
                SELECT real_value FROM component_library
                WHERE component_index = {ph} AND component_hex = {ph}
                """,
                (component_index, str(component_hex or "").strip().lower()),
            )
            row = cursor.fetchone()
            # case-insensitive fallback
            if not row:
                cursor.execute(
                    f"""
                    SELECT real_value FROM component_library
                    WHERE component_index = {ph} AND LOWER(component_hex) = {ph}
                    """,
                    (component_index, str(component_hex or "").strip().lower()),
                )
                row = cursor.fetchone()
            cursor.close()
            conn.close()

            if row:
                return row[0]
            return None

        except Exception as e:
            self.logger.error(f"Error getting component value: {e}")
            return None

    def store_guri_components(
        self,
        guri: str,
        sender: str,
        recipients: str,
        subject: str,
        avg_risk: str,
        guri_id: Optional[int] = None,
    ) -> bool:
        """
        Store all component mappings for a GURI.

        Prefer the canonical hex already allocated for each real value so the
        same field value never receives a second GURI component code.
        """
        try:
            components = guri.split("x")
            if len(components) != 6:
                self.logger.warning(
                    f"Invalid GURI format: expected 6 components, got {len(components)}"
                )
                return False

            pairs = (
                (0, sender, components[0]),
                (1, recipients, components[1]),
                (2, subject, components[2]),
                (4, avg_risk, components[4]),
            )
            ok = True
            for idx, value, hex_from_guri in pairs:
                # allocate/reuse by value; ignore stray hex if value already known
                allocated = self.allocate_hex_for_value(
                    idx, value, guri_id=guri_id
                )
                if allocated.lower() != str(hex_from_guri).lower():
                    self.logger.warning(
                        "GURI component %s hex %s does not match canonical %s for %r",
                        idx,
                        hex_from_guri,
                        allocated,
                        self._norm_value(value)[:80],
                    )
                ok = (
                    self.store_component_mapping(idx, allocated, value, guri_id)
                    and ok
                )
            return ok

        except Exception as e:
            self.logger.error(f"Error storing GURI components: {e}")
            return False

    def get_all_mappings_for_component(
        self, component_index: int, limit: int = 100
    ) -> List[Dict[str, Any]]:
        """Get all mappings for a specific component index."""
        if not self.database:
            return []

        if component_index not in self.COMPONENT_INDICES:
            return []

        try:
            conn = self.database._get_connection()
            cursor = conn.cursor()
            ph = self._ph()

            cursor.execute(
                f"""
                SELECT component_hex, real_value, guri_id, created_at, updated_at
                FROM component_library
                WHERE component_index = {ph}
                ORDER BY updated_at DESC
                LIMIT {ph}
                """,
                (component_index, limit),
            )
            rows = cursor.fetchall()
            cursor.close()
            conn.close()

            mappings = []
            for row in rows:
                mappings.append(
                    {
                        "component_hex": row[0],
                        "real_value": row[1],
                        "guri_id": row[2],
                        "created_at": row[3],
                        "updated_at": row[4] if len(row) > 4 else row[3],
                    }
                )

            return mappings

        except Exception as e:
            self.logger.error(f"Error getting component mappings: {e}")
            return []

    def get_component_statistics(self) -> Dict[str, int]:
        """Get statistics about stored component mappings."""
        if not self.database:
            return {}

        stats = {}
        try:
            conn = self.database._get_connection()
            cursor = conn.cursor()
            ph = self._ph()

            for component_index in self.COMPONENT_INDICES.keys():
                cursor.execute(
                    f"""
                    SELECT COUNT(*) FROM component_library
                    WHERE component_index = {ph}
                    """,
                    (component_index,),
                )
                count = cursor.fetchone()[0]
                component_name = self.COMPONENT_INDICES[component_index]
                stats[component_name] = count

            cursor.close()
            conn.close()

        except Exception as e:
            self.logger.error(f"Error getting component statistics: {e}")

        return stats

    def lookup_guri_components(self, guri: str) -> Dict[int, Optional[str]]:
        """Look up all component values for a GURI."""
        result: Dict[int, Optional[str]] = {}

        try:
            components = guri.split("x")
            if len(components) != 6:
                return result

            for component_index in self.COMPONENT_INDICES.keys():
                component_hex = components[component_index]
                real_value = self.get_component_value(component_index, component_hex)
                result[component_index] = real_value

            for component_index in self.EXCLUDED_COMPONENTS.keys():
                result[component_index] = None

        except Exception as e:
            self.logger.error(f"Error looking up GURI components: {e}")

        return result

    def build_stable_guri_parts(
        self,
        sender: str,
        recipients: str,
        subject: str,
        avg_risk: str,
    ) -> Tuple[str, str, str, str]:
        """Allocate/reuse hex parts for the four mappable GURI fields."""
        c0 = self.allocate_hex_for_value(0, sender)
        c1 = self.allocate_hex_for_value(1, recipients)
        c2 = self.allocate_hex_for_value(2, subject)
        c4 = self.allocate_hex_for_value(4, avg_risk)
        return c0, c1, c2, c4

    def dedupe_library(self) -> Dict[str, int]:
        """Collapse duplicate (component_index, real_value) rows to the oldest hex.

        Keeps the lowest ``id`` mapping as canonical; deletes later duplicates.
        Alias hex→value rows are removed so each value has one hex.
        """
        report = {"groups": 0, "deleted": 0, "canonical": 0}
        if not self.database:
            return report
        ph = self._ph()
        try:
            conn = self.database._get_connection()
            cursor = conn.cursor()
            cursor.execute(
                """
                SELECT component_index, real_value, COUNT(*) AS n
                FROM component_library
                GROUP BY component_index, real_value
                HAVING COUNT(*) > 1
                """
            )
            dup_groups = cursor.fetchall() or []
            report["groups"] = len(dup_groups)

            for component_index, real_value, _n in dup_groups:
                cursor.execute(
                    f"""
                    SELECT id, component_hex FROM component_library
                    WHERE component_index = {ph} AND real_value = {ph}
                    ORDER BY id ASC
                    """,
                    (component_index, real_value),
                )
                rows = cursor.fetchall() or []
                if len(rows) < 2:
                    continue
                keep_id = rows[0][0]
                report["canonical"] += 1
                drop_ids = [r[0] for r in rows[1:]]
                for did in drop_ids:
                    cursor.execute(
                        f"DELETE FROM component_library WHERE id = {ph}",
                        (did,),
                    )
                    report["deleted"] += 1
                self.logger.info(
                    "Dedupe component %s value %r → keep id %s, removed %s",
                    component_index,
                    str(real_value)[:60],
                    keep_id,
                    drop_ids,
                )

            conn.commit()
            cursor.close()
            conn.close()
        except Exception as e:
            self.logger.error(f"Error deduping component library: {e}")
        return report

    def rewrite_guri_records_to_canonical(self, *, limit: int = 0) -> Dict[str, int]:
        """Rewrite guri_records.guri strings so mappable parts use canonical hexes."""
        report = {"scanned": 0, "updated": 0, "skipped": 0, "errors": 0}
        if not self.database:
            return report
        ph = self._ph()
        try:
            conn = self.database._get_connection()
            cursor = conn.cursor()
            sql = """
                SELECT id, guri, sender, recipients, subject, avg_risk
                FROM guri_records
                ORDER BY id ASC
            """
            if limit and limit > 0:
                sql += f" LIMIT {int(limit)}"
            cursor.execute(sql)
            rows = cursor.fetchall() or []

            for row in rows:
                report["scanned"] += 1
                record_id, guri, sender, recipients, subject, avg_risk = row
                parts = str(guri or "").split("x")
                if len(parts) != 6:
                    report["skipped"] += 1
                    continue
                try:
                    c0, c1, c2, c4 = self.build_stable_guri_parts(
                        str(sender or ""),
                        str(recipients or ""),
                        str(subject or ""),
                        str(avg_risk or ""),
                    )
                    new_guri = f"{c0}x{c1}x{c2}x{parts[3]}x{c4}x{parts[5]}"
                    if new_guri == guri:
                        report["skipped"] += 1
                        continue
                    cursor.execute(
                        f"UPDATE guri_records SET guri = {ph} WHERE id = {ph}",
                        (new_guri, record_id),
                    )
                    # Keep library guri_id references useful
                    self.store_guri_components(
                        new_guri,
                        str(sender or ""),
                        str(recipients or ""),
                        str(subject or ""),
                        str(avg_risk or ""),
                        record_id,
                    )
                    report["updated"] += 1
                except Exception as exc:
                    report["errors"] += 1
                    self.logger.warning(
                        "Failed rewriting GURI id=%s: %s", record_id, exc
                    )

            conn.commit()
            cursor.close()
            conn.close()
        except Exception as e:
            self.logger.error(f"Error rewriting GURI records: {e}")
        return report

    def repair_stable_components(self, *, rewrite_records: bool = True) -> Dict[str, Any]:
        """Dedupe library then optionally rewrite guri_records to canonical hexes."""
        dedupe = self.dedupe_library()
        rewrite = (
            self.rewrite_guri_records_to_canonical()
            if rewrite_records
            else {"scanned": 0, "updated": 0, "skipped": 0, "errors": 0}
        )
        return {"dedupe": dedupe, "rewrite": rewrite}
