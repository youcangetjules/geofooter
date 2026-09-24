#!/usr/bin/env python3
"""
GURI (Globally Unique Record Identifier) Generator and Database Manager
Handles creation, storage, and retrieval of GURI records for email analysis.
Supports both SQLite and PostgreSQL databases.
"""

import json
import os
import sqlite3
import random
import logging
from datetime import datetime
from typing import Optional, Dict, Any, Tuple, List

# Try to import PostgreSQL connector
try:
    import psycopg2
    from psycopg2 import Error as PostgresError
    from psycopg2 import sql as pg_sql
    POSTGRES_AVAILABLE = True
except ImportError:
    psycopg2 = None  # type: ignore[assignment]
    pg_sql = None  # type: ignore[assignment]
    POSTGRES_AVAILABLE = False
    PostgresError = Exception

# MySQL kept as transitional fallback until Postgres cutover is complete
try:
    import mysql.connector
    from mysql.connector import Error as MySQLError
    MYSQL_AVAILABLE = True
except ImportError:
    mysql = None  # type: ignore[assignment]
    MYSQL_AVAILABLE = False
    MySQLError = Exception


class GURIDatabase:
    """Handles GURI database operations for storing and retrieving email analysis records.
    
    PostgreSQL is the supported engine. MySQL and SQLite remain only for legacy
    migration / emergency read paths and are deprecated.
    """
    
    def __init__(self, 
                 db_type: str = "postgres",
                 db_path: Optional[str] = None, 
                 base_path: str = "C:/GeoFooter",
                 pg_config: Optional[Dict[str, Any]] = None,
                 mysql_config: Optional[Dict[str, Any]] = None):
        """
        Initialize GURI database.
        
        Args:
            db_type: "postgres" (preferred), or deprecated "sqlite" / "mysql"
            db_path: Optional custom database path for SQLite. If None, uses default location.
            base_path: Base directory for default SQLite database location.
            pg_config: PostgreSQL connection configuration dict
            mysql_config: MySQL connection configuration dict (legacy / deprecated)
        """
        self.logger = logging.getLogger(__name__)
        raw_type = (db_type or "postgres").lower()
        if raw_type in {"postgresql", "pgsql", "pg"}:
            raw_type = "postgres"
        self.db_type = raw_type
        self.connection = None
        self.db_path: Optional[str] = None
        self.pg_config: Optional[Dict[str, Any]] = None
        self.mysql_config: Optional[Dict[str, Any]] = None
        
        if self.db_type == "postgres":
            if not POSTGRES_AVAILABLE or psycopg2 is None:
                raise ImportError(
                    "PostgreSQL support requires psycopg2. "
                    "Install with: pip install psycopg2-binary"
                )
            if not pg_config:
                raise ValueError("pg_config is required when db_type is 'postgres'")
            
            self.pg_config = dict(pg_config)
            if "port" not in self.pg_config:
                self.pg_config["port"] = 5432
            self.logger.info(
                "GURI database type: PostgreSQL at %s:%s",
                self.pg_config.get("host"),
                self.pg_config.get("port", 5432),
            )

        elif self.db_type == "mysql":
            self.logger.warning(
                "MySQL is deprecated for GURI — migrate to PostgreSQL "
                "(see migrate_mysql_to_postgres.py)"
            )
            if not MYSQL_AVAILABLE or mysql is None:
                raise ImportError(
                    "MySQL support requires mysql-connector-python. "
                    "Install with: pip install mysql-connector-python"
                )
            if not mysql_config:
                raise ValueError("mysql_config is required when db_type is 'mysql'")
            self.mysql_config = dict(mysql_config)
            if "port" not in self.mysql_config:
                self.mysql_config["port"] = 3306
            if "charset" not in self.mysql_config:
                self.mysql_config["charset"] = "utf8mb4"
            self.logger.info(
                "GURI database type: MySQL (deprecated) at %s:%s",
                self.mysql_config.get("host"),
                self.mysql_config.get("port", 3306),
            )
            
        elif self.db_type == "sqlite":
            self.logger.warning(
                "SQLite is deprecated for GURI — configure guri_postgres_config.json"
            )
            if db_path is None:
                base_path = os.path.abspath(base_path)
                self.db_path = os.path.join(base_path, "guri_records.db")
            else:
                self.db_path = db_path
            self.logger.info(f"GURI database type: SQLite (deprecated) at {self.db_path}")
            
        else:
            raise ValueError(
                f"Invalid db_type: {db_type}. Must be 'postgres' "
                "(or deprecated 'sqlite' / 'mysql')"
            )
        
        self._init_database()

    def _uses_pyformat(self) -> bool:
        """True when the driver uses %s placeholders (Postgres / MySQL)."""
        return self.db_type in {"postgres", "mysql"}

    @staticmethod
    def _pg_connect_kwargs(cfg: Dict[str, Any]) -> Dict[str, Any]:
        """Map our config keys to psycopg2.connect kwargs."""
        return {
            "host": cfg["host"],
            "port": int(cfg.get("port", 5432)),
            "user": cfg["user"],
            "password": cfg.get("password") or "",
            "dbname": cfg["database"],
        }

    def _ensure_postgres_database(self) -> None:
        """Create the target database if it does not exist yet."""
        if not self.pg_config or psycopg2 is None or pg_sql is None:
            return
        dbname = str(self.pg_config.get("database") or "").strip()
        if not dbname:
            raise ValueError("PostgreSQL database name is required")
        admin_cfg = dict(self.pg_config)
        admin_cfg["database"] = "postgres"
        conn = psycopg2.connect(**self._pg_connect_kwargs(admin_cfg))
        conn.autocommit = True
        try:
            cursor = conn.cursor()
            cursor.execute("SELECT 1 FROM pg_database WHERE datname = %s", (dbname,))
            exists = cursor.fetchone() is not None
            if not exists:
                cursor.execute(
                    pg_sql.SQL("CREATE DATABASE {}").format(pg_sql.Identifier(dbname))
                )
                self.logger.info("Created PostgreSQL database %s", dbname)
            cursor.close()
        finally:
            conn.close()
    
    def _get_connection(self):
        """Get database connection based on database type."""
        if self.db_type == "postgres":
            if not POSTGRES_AVAILABLE or psycopg2 is None:
                raise ImportError(
                    "PostgreSQL support requires psycopg2. "
                    "Install with: pip install psycopg2-binary"
                )
            if not self.pg_config:
                raise ValueError("pg_config is not set")
            return psycopg2.connect(**self._pg_connect_kwargs(self.pg_config))
        if self.db_type == "mysql":
            if not MYSQL_AVAILABLE or mysql is None:
                raise ImportError(
                    "MySQL support requires mysql-connector-python. "
                    "Install with: pip install mysql-connector-python"
                )
            if not self.mysql_config:
                raise ValueError("mysql_config is not set")
            return mysql.connector.connect(**self.mysql_config)
        if not self.db_path:
            raise ValueError("SQLite db_path is not set")
        return sqlite3.connect(self.db_path)

    @staticmethod
    def _as_row(row: Any) -> Optional[Tuple[Any, ...]]:
        """Normalize a DB fetchone() result to a tuple for integer indexing."""
        if row is None:
            return None
        return tuple(row)

    @staticmethod
    def _as_rows(rows: Any) -> List[Tuple[Any, ...]]:
        """Normalize a DB fetchall() result to a list of tuples."""
        return [tuple(row) for row in rows]

    @staticmethod
    def _record_from_row(row: Tuple[Any, ...]) -> Dict[str, Any]:
        """Map a guri_records SELECT row to a dict (handles older schemas)."""
        cols = list(row)
        if len(cols) > 9:
            sensitivity = cols[8]
            created_at = cols[9]
        else:
            sensitivity = '[SEC1:(U)EXTERNAL/UNRATED]'
            created_at = cols[8] if len(cols) > 8 else None
        return {
            'id': cols[0],
            'guri': cols[1],
            'sender': cols[2],
            'recipients': cols[3],
            'subject': cols[4],
            'datetime': cols[5],
            'avg_risk': cols[6],
            'document_type': cols[7],
            'sensitivity': sensitivity,
            'created_at': created_at,
        }
    
    def _init_database(self):
        """Initialize the database and create tables if they don't exist."""
        try:
            if self.db_type == "postgres":
                self._ensure_postgres_database()
                conn = self._get_connection()
                cursor = conn.cursor()
                
                # PostgreSQL syntax
                # avg_risk also stores document file paths for non-email types — must be TEXT.
                cursor.execute('''
                    CREATE TABLE IF NOT EXISTS guri_records (
                        id SERIAL PRIMARY KEY,
                        guri VARCHAR(50) UNIQUE NOT NULL,
                        sender VARCHAR(255) NOT NULL,
                        recipients TEXT,
                        subject TEXT,
                        datetime VARCHAR(50),
                        avg_risk TEXT,
                        document_type VARCHAR(10) DEFAULT '01',
                        sensitivity VARCHAR(100) DEFAULT '[SEC1:(U)EXTERNAL/UNRATED]',
                        company_name VARCHAR(255),
                        created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
                    )
                ''')
                cursor.execute('CREATE INDEX IF NOT EXISTS idx_guri ON guri_records (guri)')
                cursor.execute('CREATE INDEX IF NOT EXISTS idx_sender ON guri_records (sender)')
                cursor.execute('CREATE INDEX IF NOT EXISTS idx_created_at ON guri_records (created_at)')
                cursor.execute('CREATE INDEX IF NOT EXISTS idx_company_name ON guri_records (company_name)')
                self._migrate_avg_risk_to_text(cursor)
                
                cursor.execute('''
                    CREATE TABLE IF NOT EXISTS supplier_codes (
                        id SERIAL PRIMARY KEY,
                        company_name VARCHAR(255) UNIQUE NOT NULL,
                        company_domain VARCHAR(255),
                        supplier_code VARCHAR(100),
                        contract_code VARCHAR(100),
                        contact_person VARCHAR(255),
                        contact_email VARCHAR(255),
                        notes TEXT,
                        created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                        updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
                    )
                ''')
                cursor.execute(
                    'CREATE INDEX IF NOT EXISTS idx_supplier_company_name ON supplier_codes (company_name)'
                )
                cursor.execute(
                    'CREATE INDEX IF NOT EXISTS idx_supplier_company_domain ON supplier_codes (company_domain)'
                )
                conn.commit()
                cursor.close()
                conn.close()
                self.logger.info("GURI PostgreSQL database initialized")

            elif self.db_type == "mysql":
                conn = self._get_connection()
                cursor = conn.cursor()
                cursor.execute('''
                    CREATE TABLE IF NOT EXISTS guri_records (
                        id INT AUTO_INCREMENT PRIMARY KEY,
                        guri VARCHAR(50) UNIQUE NOT NULL,
                        sender VARCHAR(255) NOT NULL,
                        recipients TEXT,
                        subject TEXT,
                        datetime VARCHAR(50),
                        avg_risk TEXT,
                        document_type VARCHAR(10) DEFAULT '01',
                        sensitivity VARCHAR(100) DEFAULT '[SEC1:(U)EXTERNAL/UNRATED]',
                        company_name VARCHAR(255),
                        created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                        INDEX idx_guri (guri),
                        INDEX idx_sender (sender),
                        INDEX idx_created_at (created_at),
                        INDEX idx_company_name (company_name)
                    ) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4
                ''')
                try:
                    cursor.execute('ALTER TABLE guri_records MODIFY COLUMN avg_risk TEXT')
                except Exception as mig_exc:
                    self.logger.debug("MySQL avg_risk widen skipped: %s", mig_exc)
                cursor.execute('''
                    CREATE TABLE IF NOT EXISTS supplier_codes (
                        id INT AUTO_INCREMENT PRIMARY KEY,
                        company_name VARCHAR(255) UNIQUE NOT NULL,
                        company_domain VARCHAR(255),
                        supplier_code VARCHAR(100),
                        contract_code VARCHAR(100),
                        contact_person VARCHAR(255),
                        contact_email VARCHAR(255),
                        notes TEXT,
                        created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                        updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP ON UPDATE CURRENT_TIMESTAMP,
                        INDEX idx_company_name (company_name),
                        INDEX idx_company_domain (company_domain)
                    ) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4
                ''')
                conn.commit()
                cursor.close()
                conn.close()
                self.logger.info("GURI MySQL database initialized")
                
            else:  # sqlite
                if not self.db_path:
                    raise ValueError("SQLite db_path is not set")
                with sqlite3.connect(self.db_path) as conn:
                    cursor = conn.cursor()
                    cursor.execute('''
                        CREATE TABLE IF NOT EXISTS guri_records (
                            id INTEGER PRIMARY KEY AUTOINCREMENT,
                            guri TEXT UNIQUE NOT NULL,
                            sender TEXT NOT NULL,
                            recipients TEXT,
                            subject TEXT,
                            datetime TEXT,
                            avg_risk TEXT,
                            document_type TEXT DEFAULT '01',
                            sensitivity TEXT DEFAULT '[SEC1:(U)EXTERNAL/UNRATED]',
                            company_name TEXT,
                            created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
                        )
                    ''')
                    # Create indexes for SQLite
                    cursor.execute('CREATE INDEX IF NOT EXISTS idx_guri ON guri_records(guri)')
                    cursor.execute('CREATE INDEX IF NOT EXISTS idx_sender ON guri_records(sender)')
                    cursor.execute('CREATE INDEX IF NOT EXISTS idx_company_name ON guri_records(company_name)')
                    
                    # Create supplier_codes table
                    cursor.execute('''
                        CREATE TABLE IF NOT EXISTS supplier_codes (
                            id INTEGER PRIMARY KEY AUTOINCREMENT,
                            company_name TEXT UNIQUE NOT NULL,
                            company_domain TEXT,
                            supplier_code TEXT,
                            contract_code TEXT,
                            contact_person TEXT,
                            contact_email TEXT,
                            notes TEXT,
                            created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                            updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
                        )
                    ''')
                    cursor.execute('CREATE INDEX IF NOT EXISTS idx_supplier_company ON supplier_codes(company_name)')
                    cursor.execute('CREATE INDEX IF NOT EXISTS idx_supplier_domain ON supplier_codes(company_domain)')
                    
                    conn.commit()
                    self.logger.info(f"GURI SQLite database initialized at: {self.db_path}")
                    
                    # Auto-migrate: Add missing columns if they don't exist
                    try:
                        cursor.execute("PRAGMA table_info(guri_records)")
                        columns = [col[1] for col in self._as_rows(cursor.fetchall())]
                        
                        if 'sensitivity' not in columns:
                            self.logger.info("Adding sensitivity column...")
                            cursor.execute('ALTER TABLE guri_records ADD COLUMN sensitivity TEXT DEFAULT \'[SEC1:(U)EXTERNAL/UNRATED]\'')
                            conn.commit()
                        
                        if 'company_name' not in columns:
                            self.logger.info("Adding company_name column...")
                            cursor.execute('ALTER TABLE guri_records ADD COLUMN company_name TEXT')
                            conn.commit()
                            self.logger.info("company_name column added successfully")
                    except Exception as e:
                        self.logger.warning(f"Could not add columns (may already exist): {e}")
                    
        except Exception as e:
            self.logger.error(f"Failed to initialize GURI database: {e}")
            raise

    def _migrate_avg_risk_to_text(self, cursor) -> None:
        """Widen avg_risk so document file paths fit (was VARCHAR(50))."""
        if self.db_type != "postgres":
            return
        try:
            cursor.execute(
                """
                SELECT character_maximum_length
                FROM information_schema.columns
                WHERE table_schema = current_schema()
                  AND table_name = 'guri_records'
                  AND column_name = 'avg_risk'
                """
            )
            row = self._as_row(cursor.fetchone())
            max_len = row[0] if row else None
            if max_len is not None:
                cursor.execute("ALTER TABLE guri_records ALTER COLUMN avg_risk TYPE TEXT")
                self.logger.info("Migrated guri_records.avg_risk VARCHAR(%s) → TEXT", max_len)
        except Exception as exc:
            self.logger.warning("avg_risk TEXT migration skipped: %s", exc)
    
    def get_guri(self, sender: str, recipients: str, subject: str, dt_str: str, avg_risk: str) -> Optional[str]:
        """
        Retrieve existing GURI for given parameters, or None if not found.
        
        Args:
            sender: Email sender address
            recipients: Recipient email addresses
            subject: Email subject line
            dt_str: DateTime string
            avg_risk: Average risk assessment string
            
        Returns:
            Existing GURI string if found, None otherwise
        """
        try:
            conn = self._get_connection()
            cursor = conn.cursor()
            
            if self._uses_pyformat():
                query = '''
                    SELECT guri FROM guri_records 
                    WHERE sender = %s AND recipients = %s AND subject = %s AND datetime = %s AND avg_risk = %s
                '''
            else:  # sqlite
                query = '''
                    SELECT guri FROM guri_records 
                    WHERE sender = ? AND recipients = ? AND subject = ? AND datetime = ? AND avg_risk = ?
                '''
            
            cursor.execute(query, (sender, recipients, subject, dt_str, avg_risk))
            result = self._as_row(cursor.fetchone())
            cursor.close()
            conn.close()
            
            if result:
                guri = str(result[0])
                self.logger.info(f"Found existing GURI: {guri}")
                return guri
            return None
            
        except Exception as e:
            self.logger.error(f"Error retrieving GURI: {e}")
            return None
    
    def insert_guri(self, guri: str, sender: str, recipients: str, subject: str, dt_str: str, avg_risk: str, document_type: str = "01", sensitivity: str = "[SEC1:(U)EXTERNAL/UNRATED]") -> Optional[int]:
        """
        Insert a new GURI record into the database.
        
        Args:
            guri: The generated GURI string
            sender: Email sender address
            recipients: Recipient email addresses
            subject: Email subject line
            dt_str: DateTime string
            avg_risk: Average risk assessment string
            document_type: Document type code (01=Email, 02=Word, 03=Excel, etc.)
            sensitivity: Security classification level
            
        Returns:
            Record ID if successful, None otherwise
        """
        try:
            conn = self._get_connection()
            cursor = conn.cursor()
            
            if self.db_type == "postgres":
                # PostgreSQL: ON CONFLICT DO NOTHING
                query = '''
                    INSERT INTO guri_records 
                    (guri, sender, recipients, subject, datetime, avg_risk, document_type, sensitivity)
                    VALUES (%s, %s, %s, %s, %s, %s, %s, %s)
                    ON CONFLICT (guri) DO NOTHING
                    RETURNING id
                '''
            elif self.db_type == "mysql":
                query = '''
                    INSERT IGNORE INTO guri_records 
                    (guri, sender, recipients, subject, datetime, avg_risk, document_type, sensitivity)
                    VALUES (%s, %s, %s, %s, %s, %s, %s, %s)
                '''
            else:  # sqlite
                # SQLite: INSERT OR IGNORE
                query = '''
                    INSERT OR IGNORE INTO guri_records 
                    (guri, sender, recipients, subject, datetime, avg_risk, document_type, sensitivity)
                    VALUES (?, ?, ?, ?, ?, ?, ?, ?)
                '''
            
            cursor.execute(query, (guri, sender, recipients, subject, dt_str, avg_risk, document_type, sensitivity))
            
            record_id = None
            if self.db_type == "postgres":
                row = self._as_row(cursor.fetchone())
                if row:
                    record_id = int(row[0])
                else:
                    cursor.execute(
                        'SELECT id FROM guri_records WHERE guri = %s LIMIT 1',
                        (guri,),
                    )
                    row = self._as_row(cursor.fetchone())
                    if row:
                        record_id = int(row[0])
                conn.commit()
            elif self.db_type == "mysql":
                conn.commit()
                record_id = cursor.lastrowid
                if record_id == 0:
                    cursor.execute(
                        'SELECT id FROM guri_records WHERE guri = %s LIMIT 1',
                        (guri,),
                    )
                    row = self._as_row(cursor.fetchone())
                    if row:
                        record_id = int(row[0])
            else:
                conn.commit()
                record_id = cursor.lastrowid
                if record_id == 0:
                    cursor.execute(
                        'SELECT id FROM guri_records WHERE guri = ? LIMIT 1',
                        (guri,),
                    )
                    row = self._as_row(cursor.fetchone())
                    if row:
                        record_id = int(row[0])
            
            cursor.close()
            conn.close()
            
            if record_id:
                self.logger.info(f"Inserted/Found GURI record: {guri} (ID: {record_id}, Document Type: {document_type}, Sensitivity: {sensitivity})")
            else:
                self.logger.warning(f"Could not determine record ID for GURI: {guri}")
            
            return int(record_id) if record_id else None
            
        except Exception as e:
            self.logger.error(f"Error inserting GURI: {e}")
            raise
    
    def generate_guri(self, sender: str, recipients: str, subject: str, dt_str: str, avg_risk: str, document_type: str = "01", sensitivity: str = "[SEC1:(U)EXTERNAL/UNRATED]") -> str:
        """
        Generate a new GURI or retrieve existing one from database.
        
        The GURI format is: {5hex}x{5hex}x{8hex}x{8hex}x{3hex}x{2hex}
        Example: a1b2cx3d4e5xf6g7h8i9xj0k1l2m3xn4oxp5
        
        Args:
            sender: Email sender address (or julian.garrett@aliniant.com for non-emails)
            recipients: Recipient email addresses (or "FFFFF" for non-emails)
            subject: Email subject line (or document name)
            dt_str: DateTime string
            avg_risk: Average risk assessment string (or "FFFFF" for non-emails)
            document_type: Document type code (01=Email, 02=Word, 03=Excel, etc.)
            sensitivity: Security classification level
            
        Returns:
            GURI string (either existing or newly generated)
        """
        # First, try to get existing GURI
        existing_guri = self.get_guri(sender, recipients, subject, dt_str, avg_risk)
        if existing_guri:
            return existing_guri
        
        # Generate new GURI — mappable fields reuse stable hex for the same real value
        def randhex(n):
            """Generate random hexadecimal string of length n."""
            return ''.join(random.choices('0123456789abcdef', k=n))
        
        def encode_datetime_component(dt_str: str) -> str:
            """
            Encode datetime as 8-digit hex representing seconds since midnight.
            Reference: Midnight of the current date = 00000000
            
            Args:
                dt_str: DateTime string (e.g., "2025-11-02 14:30:45")
            
            Returns:
                8-character hex string
            """
            formats = (
                "%Y-%m-%d %H:%M:%S",
                "%Y/%m/%d %H:%M:%S",
                "%Y-%m-%d %H:%M",
                "%Y/%m/%d %H:%M",
                "%d/%m/%Y %H:%M:%S",
                "%d-%m-%Y %H:%M:%S",
            )
            last_err: Optional[Exception] = None
            for fmt in formats:
                try:
                    dt = datetime.strptime(dt_str.strip(), fmt)
                    midnight = dt.replace(hour=0, minute=0, second=0, microsecond=0)
                    seconds_since_midnight = int((dt - midnight).total_seconds())
                    return format(seconds_since_midnight, "08x")
                except Exception as e:
                    last_err = e
            self.logger.warning(
                "Error encoding datetime '%s': %s. Using random component.",
                dt_str,
                last_err,
            )
            return randhex(8)
        
        from guri_component_library import GURIComponentLibrary
        library = GURIComponentLibrary(database=self)
        component1, component2, component3, component5 = library.build_stable_guri_parts(
            sender, recipients, subject, avg_risk
        )
        component4 = encode_datetime_component(dt_str)  # DateTime (seconds since midnight)
        component6 = document_type   # Document Type (actual code, not random!)
        
        guri = f"{component1}x{component2}x{component3}x{component4}x{component5}x{component6}"
        
        # Store in database — creation is not successful unless the row lands.
        record_id = self.insert_guri(guri, sender, recipients, subject, dt_str, avg_risk, document_type, sensitivity)
        if not record_id:
            raise RuntimeError(
                f"GURI was generated ({guri}) but could not be saved to the database."
            )
        
        # Store / refresh component mappings (same value → same hex)
        try:
            library.store_guri_components(guri, sender, recipients, subject, avg_risk, record_id)
        except Exception as e:
            self.logger.warning(f"Failed to store component mappings: {e}")
        
        return guri
    
    def get_all_records(self, limit: int = 100, offset: int = 0) -> list:
        """
        Retrieve all GURI records from the database.
        
        Args:
            limit: Maximum number of records to return
            offset: Number of records to skip
            
        Returns:
            List of dictionaries containing all record fields
        """
        try:
            conn = self._get_connection()
            cursor = conn.cursor()
            
            if self._uses_pyformat():
                query = '''
                    SELECT id, guri, sender, recipients, subject, datetime, avg_risk, 
                           document_type, sensitivity, created_at
                    FROM guri_records
                    ORDER BY created_at DESC
                    LIMIT %s OFFSET %s
                '''
            else:  # sqlite
                query = '''
                    SELECT id, guri, sender, recipients, subject, datetime, avg_risk, 
                           document_type, sensitivity, created_at
                    FROM guri_records
                    ORDER BY created_at DESC
                    LIMIT ? OFFSET ?
                '''
            
            cursor.execute(query, (limit, offset))
            rows = self._as_rows(cursor.fetchall())
            cursor.close()
            conn.close()
            
            records = []
            for row in rows:
                records.append(self._record_from_row(row))
            
            return records
            
        except Exception as e:
            self.logger.error(f"Error retrieving records: {e}")
            return []
    
    def get_record_count(self) -> int:
        """
        Get total count of GURI records in the database.
        
        Returns:
            Total number of records
        """
        try:
            conn = self._get_connection()
            cursor = conn.cursor()
            
            query = "SELECT COUNT(*) FROM guri_records"
            cursor.execute(query)
            row = self._as_row(cursor.fetchone())
            cursor.close()
            conn.close()
            
            return int(row[0]) if row else 0
            
        except Exception as e:
            self.logger.error(f"Error getting record count: {e}")
            return 0
    
    def get_record_by_guri(self, guri: str) -> Optional[Dict[str, Any]]:
        """
        Get a record by GURI string.
        
        Args:
            guri: GURI string to search for
            
        Returns:
            Dictionary containing record data or None if not found
        """
        try:
            conn = self._get_connection()
            cursor = conn.cursor()
            
            if self._uses_pyformat():
                query = '''
                    SELECT id, guri, sender, recipients, subject, datetime, avg_risk, 
                           document_type, sensitivity, created_at
                    FROM guri_records
                    WHERE guri = %s
                    LIMIT 1
                '''
            else:  # sqlite
                query = '''
                    SELECT id, guri, sender, recipients, subject, datetime, avg_risk, 
                           document_type, sensitivity, created_at
                    FROM guri_records
                    WHERE guri = ?
                    LIMIT 1
                '''
            
            cursor.execute(query, (guri,))
            row = self._as_row(cursor.fetchone())
            cursor.close()
            conn.close()
            
            if row:
                return self._record_from_row(row)
            return None
            
        except Exception as e:
            self.logger.error(f"Error getting record by GURI: {e}")
            return None
    
    def get_subjects_by_recipient(self, recipient: str, limit: int = 50) -> list:
        """
        Get distinct subjects for a specific recipient.
        
        Args:
            recipient: Recipient email address to search for
            limit: Maximum number of subjects to return
            
        Returns:
            List of distinct subject strings for that recipient
        """
        try:
            conn = self._get_connection()
            cursor = conn.cursor()
            
            if self._uses_pyformat():
                query = '''
                    SELECT DISTINCT subject
                    FROM guri_records
                    WHERE recipients = %s
                    ORDER BY created_at DESC
                    LIMIT %s
                '''
                cursor.execute(query, (recipient, limit))
            else:  # sqlite
                query = '''
                    SELECT DISTINCT subject
                    FROM guri_records
                    WHERE recipients = ?
                    ORDER BY created_at DESC
                    LIMIT ?
                '''
                cursor.execute(query, (recipient, limit))
            
            rows = self._as_rows(cursor.fetchall())
            cursor.close()
            conn.close()
            
            subjects = [row[0] for row in rows if row[0]]
            return subjects
            
        except Exception as e:
            self.logger.error(f"Error getting subjects by recipient: {e}")
            return []
    
    def search_by_sender(self, sender: str, limit: int = 50) -> list:
        """
        Search for GURI records by sender email address.
        
        Args:
            sender: Email sender to search for (partial match supported)
            limit: Maximum number of records to return
            
        Returns:
            List of matching records
        """
        try:
            conn = self._get_connection()
            cursor = conn.cursor()
            
            if self._uses_pyformat():
                query = '''
                    SELECT id, guri, sender, recipients, subject, datetime, avg_risk, 
                           document_type, sensitivity, created_at
                    FROM guri_records
                    WHERE sender LIKE %s
                    ORDER BY created_at DESC
                    LIMIT %s
                '''
                cursor.execute(query, (f"%{sender}%", limit))
            else:  # sqlite
                query = '''
                    SELECT id, guri, sender, recipients, subject, datetime, avg_risk, 
                           document_type, sensitivity, created_at
                    FROM guri_records
                    WHERE sender LIKE ?
                    ORDER BY created_at DESC
                    LIMIT ?
                '''
                cursor.execute(query, (f"%{sender}%", limit))
            
            rows = self._as_rows(cursor.fetchall())
            cursor.close()
            conn.close()
            
            records = []
            for row in rows:
                records.append(self._record_from_row(row))
            
            return records
            
        except Exception as e:
            self.logger.error(f"Error searching records: {e}")
            return []

    def search_by_sender_domains(
        self,
        domains: list,
        *,
        limit_per_domain: int = 40,
        document_type: str = "01",
    ) -> list:
        """Find email GURI records whose sender address is at one of the domains.

        Uses targeted LIKE queries (not a full-table dump) so Welcome
        Important Emails can stay current without scanning 10k rows.
        """
        out: list = []
        seen: set = set()
        for raw in domains or []:
            domain = str(raw or "").strip().lower().lstrip("@")
            if not domain or "@" in domain:
                continue
            try:
                # Prefer "@domain" so we match the address host, not substrings
                # in the local-part. Cap per domain so UI stays snappy.
                chunk = self.search_by_sender(f"@{domain}", limit=int(limit_per_domain))
            except Exception as e:
                self.logger.error(f"Error searching domain {domain}: {e}")
                continue
            for record in chunk:
                if document_type and str(record.get("document_type") or "") != document_type:
                    continue
                sender = str(record.get("sender") or "").lower()
                # Confirm real domain match (handles Name <addr@domain> forms)
                host = ""
                if "@" in sender:
                    host = sender.rsplit("@", 1)[-1].strip(">")
                if not host or not (
                    host == domain or host.endswith("." + domain)
                ):
                    continue
                key = str(record.get("guri") or record.get("id") or "")
                if key and key in seen:
                    continue
                if key:
                    seen.add(key)
                out.append(record)
        return out
    
    @staticmethod
    def extract_company_from_domain(email: str) -> Optional[str]:
        """
        Extract company name from email domain.
        
        Args:
            email: Email address
            
        Returns:
            Company name (domain without TLD) or None
        """
        try:
            if '@' not in email:
                return None
            
            domain = email.split('@')[1].lower()
            # Remove common TLDs
            parts = domain.split('.')
            if len(parts) >= 2:
                # Return second-level domain (e.g., 'google' from 'google.com')
                return parts[-2].capitalize()
            return domain.capitalize()
        except:
            return None
    
    def add_supplier_code(self, company_name: str, company_domain: str = "", supplier_code: str = "",
                         contract_code: str = "", contact_person: str = "", contact_email: str = "",
                         notes: str = "") -> bool:
        """
        Add or update supplier code information for a company.
        
        Args:
            company_name: Company name
            company_domain: Company domain
            supplier_code: Supplier code
            contract_code: Contract code
            contact_person: Contact person name
            contact_email: Contact person email
            notes: Additional notes
            
        Returns:
            True if successful, False otherwise
        """
        try:
            conn = self._get_connection()
            cursor = conn.cursor()
            
            if self.db_type == "postgres":
                query = '''
                    INSERT INTO supplier_codes 
                    (company_name, company_domain, supplier_code, contract_code, 
                     contact_person, contact_email, notes)
                    VALUES (%s, %s, %s, %s, %s, %s, %s)
                    ON CONFLICT (company_name) DO UPDATE SET
                    company_domain = EXCLUDED.company_domain,
                    supplier_code = EXCLUDED.supplier_code,
                    contract_code = EXCLUDED.contract_code,
                    contact_person = EXCLUDED.contact_person,
                    contact_email = EXCLUDED.contact_email,
                    notes = EXCLUDED.notes,
                    updated_at = CURRENT_TIMESTAMP
                '''
            elif self.db_type == "mysql":
                query = '''
                    INSERT INTO supplier_codes 
                    (company_name, company_domain, supplier_code, contract_code, 
                     contact_person, contact_email, notes)
                    VALUES (%s, %s, %s, %s, %s, %s, %s)
                    ON DUPLICATE KEY UPDATE
                    company_domain = VALUES(company_domain),
                    supplier_code = VALUES(supplier_code),
                    contract_code = VALUES(contract_code),
                    contact_person = VALUES(contact_person),
                    contact_email = VALUES(contact_email),
                    notes = VALUES(notes),
                    updated_at = CURRENT_TIMESTAMP
                '''
            else:  # sqlite
                query = '''
                    INSERT OR REPLACE INTO supplier_codes 
                    (company_name, company_domain, supplier_code, contract_code,
                     contact_person, contact_email, notes, updated_at)
                    VALUES (?, ?, ?, ?, ?, ?, ?, datetime('now'))
                '''
            
            cursor.execute(query, (company_name, company_domain, supplier_code, contract_code,
                                 contact_person, contact_email, notes))
            conn.commit()
            cursor.close()
            conn.close()
            
            self.logger.info(f"Added/Updated supplier code for: {company_name}")
            return True
            
        except Exception as e:
            self.logger.error(f"Error adding supplier code: {e}")
            return False
    
    def get_all_suppliers(self) -> list:
        """
        Get all supplier codes.
        
        Returns:
            List of dictionaries containing supplier information
        """
        try:
            conn = self._get_connection()
            cursor = conn.cursor()
            
            query = '''
                SELECT id, company_name, company_domain, supplier_code, contract_code,
                       contact_person, contact_email, notes, created_at, updated_at
                FROM supplier_codes
                ORDER BY company_name
            '''
            
            cursor.execute(query)
            rows = self._as_rows(cursor.fetchall())
            cursor.close()
            conn.close()
            
            suppliers = []
            for row in rows:
                suppliers.append({
                    'id': row[0],
                    'company_name': row[1],
                    'company_domain': row[2],
                    'supplier_code': row[3],
                    'contract_code': row[4],
                    'contact_person': row[5],
                    'contact_email': row[6],
                    'notes': row[7],
                    'created_at': row[8],
                    'updated_at': row[9]
                })
            
            return suppliers
            
        except Exception as e:
            self.logger.error(f"Error retrieving suppliers: {e}")
            return []
    
    def get_supplier_by_company(self, company_name: str) -> Optional[dict]:
        """
        Get supplier information by company name.
        
        Args:
            company_name: Company name to search for
            
        Returns:
            Dictionary containing supplier information or None
        """
        try:
            conn = self._get_connection()
            cursor = conn.cursor()
            
            if self._uses_pyformat():
                query = '''
                    SELECT id, company_name, company_domain, supplier_code, contract_code,
                           contact_person, contact_email, notes, created_at, updated_at
                    FROM supplier_codes
                    WHERE company_name = %s
                '''
            else:  # sqlite
                query = '''
                    SELECT id, company_name, company_domain, supplier_code, contract_code,
                           contact_person, contact_email, notes, created_at, updated_at
                    FROM supplier_codes
                    WHERE company_name = ?
                '''
            
            cursor.execute(query, (company_name,))
            row = self._as_row(cursor.fetchone())
            cursor.close()
            conn.close()
            
            if row:
                return {
                    'id': row[0],
                    'company_name': row[1],
                    'company_domain': row[2],
                    'supplier_code': row[3],
                    'contract_code': row[4],
                    'contact_person': row[5],
                    'contact_email': row[6],
                    'notes': row[7],
                    'created_at': row[8],
                    'updated_at': row[9]
                }
            return None
            
        except Exception as e:
            self.logger.error(f"Error retrieving supplier: {e}")
            return None
    
    def delete_supplier_code(self, company_name: str) -> bool:
        """
        Delete supplier code by company name.
        
        Args:
            company_name: Company name
            
        Returns:
            True if successful, False otherwise
        """
        try:
            conn = self._get_connection()
            cursor = conn.cursor()
            
            if self._uses_pyformat():
                query = 'DELETE FROM supplier_codes WHERE company_name = %s'
            else:  # sqlite
                query = 'DELETE FROM supplier_codes WHERE company_name = ?'
            
            cursor.execute(query, (company_name,))
            conn.commit()
            cursor.close()
            conn.close()
            
            self.logger.info(f"Deleted supplier code for: {company_name}")
            return True
            
        except Exception as e:
            self.logger.error(f"Error deleting supplier code: {e}")
            return False

    # ------------------------------------------------------------------
    # Introspection (Database browser tab)
    # ------------------------------------------------------------------

    def is_deprecated_engine(self) -> bool:
        return self.db_type in {"mysql", "sqlite"}

    def get_connection_info(self) -> Dict[str, Any]:
        """Summary of the active connection for UI display."""
        info: Dict[str, Any] = {
            "engine": self.db_type,
            "deprecated": self.is_deprecated_engine(),
            "label": self.db_type.upper(),
            "host": "",
            "port": "",
            "database": "",
            "user": "",
            "path": "",
            "server_version": "",
        }
        try:
            if self.db_type == "postgres" and self.pg_config:
                info.update(
                    {
                        "label": "PostgreSQL",
                        "host": str(self.pg_config.get("host") or ""),
                        "port": str(self.pg_config.get("port") or 5432),
                        "database": str(self.pg_config.get("database") or ""),
                        "user": str(self.pg_config.get("user") or ""),
                    }
                )
                conn = self._get_connection()
                try:
                    cur = conn.cursor()
                    cur.execute("SELECT version()")
                    row = self._as_row(cur.fetchone())
                    info["server_version"] = str(row[0]) if row else ""
                    cur.close()
                finally:
                    conn.close()
            elif self.db_type == "mysql" and self.mysql_config:
                info.update(
                    {
                        "label": "MySQL (deprecated)",
                        "host": str(self.mysql_config.get("host") or ""),
                        "port": str(self.mysql_config.get("port") or 3306),
                        "database": str(self.mysql_config.get("database") or ""),
                        "user": str(self.mysql_config.get("user") or ""),
                    }
                )
            elif self.db_type == "sqlite":
                info.update(
                    {
                        "label": "SQLite (deprecated)",
                        "database": os.path.basename(self.db_path or "") or "sqlite",
                        "path": str(self.db_path or ""),
                    }
                )
        except Exception as exc:
            self.logger.warning("get_connection_info failed: %s", exc)
        return info

    def list_schemas(self) -> List[str]:
        """List non-system schemas (Postgres) or database name (MySQL/SQLite)."""
        try:
            conn = self._get_connection()
            cur = conn.cursor()
            if self.db_type == "postgres":
                cur.execute(
                    """
                    SELECT nspname
                    FROM pg_namespace
                    WHERE nspname NOT LIKE 'pg\\_%'
                      AND nspname <> 'information_schema'
                    ORDER BY nspname
                    """
                )
                schemas = [str(r[0]) for r in self._as_rows(cur.fetchall())]
            elif self.db_type == "mysql" and self.mysql_config:
                schemas = [str(self.mysql_config.get("database") or "")]
            else:
                schemas = ["main"]
            cur.close()
            conn.close()
            return [s for s in schemas if s]
        except Exception as exc:
            self.logger.error("list_schemas failed: %s", exc)
            return []

    def list_tables(self, schema: Optional[str] = None) -> List[Dict[str, Any]]:
        """List tables with approximate row counts where available."""
        out: List[Dict[str, Any]] = []
        try:
            conn = self._get_connection()
            cur = conn.cursor()
            if self.db_type == "postgres":
                if schema:
                    cur.execute(
                        """
                        SELECT n.nspname, c.relname,
                               COALESCE(s.n_live_tup, c.reltuples)::bigint
                        FROM pg_class c
                        JOIN pg_namespace n ON n.oid = c.relnamespace
                        LEFT JOIN pg_stat_all_tables s
                          ON s.relid = c.oid
                        WHERE c.relkind IN ('r', 'p')
                          AND n.nspname = %s
                        ORDER BY c.relname
                        """,
                        (schema,),
                    )
                else:
                    cur.execute(
                        """
                        SELECT n.nspname, c.relname,
                               COALESCE(s.n_live_tup, c.reltuples)::bigint
                        FROM pg_class c
                        JOIN pg_namespace n ON n.oid = c.relnamespace
                        LEFT JOIN pg_stat_all_tables s
                          ON s.relid = c.oid
                        WHERE c.relkind IN ('r', 'p')
                          AND n.nspname NOT LIKE 'pg\\_%'
                          AND n.nspname <> 'information_schema'
                        ORDER BY n.nspname, c.relname
                        """
                    )
                for schema_name, name, rows in self._as_rows(cur.fetchall()):
                    out.append(
                        {
                            "schema": str(schema_name),
                            "name": str(name),
                            "rows": int(rows or 0),
                            "full_name": f"{schema_name}.{name}",
                        }
                    )
            elif self.db_type == "mysql":
                dbname = (self.mysql_config or {}).get("database") or schema or ""
                cur.execute(
                    """
                    SELECT table_schema, table_name, table_rows
                    FROM information_schema.tables
                    WHERE table_schema = %s AND table_type = 'BASE TABLE'
                    ORDER BY table_name
                    """,
                    (dbname,),
                )
                for schema_name, name, rows in self._as_rows(cur.fetchall()):
                    out.append(
                        {
                            "schema": str(schema_name),
                            "name": str(name),
                            "rows": int(rows or 0),
                            "full_name": str(name),
                        }
                    )
            else:
                cur.execute(
                    "SELECT name FROM sqlite_master WHERE type='table' "
                    "AND name NOT LIKE 'sqlite_%' ORDER BY name"
                )
                for (name,) in self._as_rows(cur.fetchall()):
                    count = 0
                    try:
                        cur.execute(f'SELECT COUNT(*) FROM "{name}"')
                        row = self._as_row(cur.fetchone())
                        count = int(row[0]) if row else 0
                    except Exception:
                        count = 0
                    out.append(
                        {
                            "schema": "main",
                            "name": str(name),
                            "rows": count,
                            "full_name": str(name),
                        }
                    )
            cur.close()
            conn.close()
        except Exception as exc:
            self.logger.error("list_tables failed: %s", exc)
        return out

    def describe_table(
        self, table: str, schema: Optional[str] = None
    ) -> List[Dict[str, Any]]:
        """Column metadata for a table."""
        table = (table or "").strip()
        if not table:
            return []
        cols: List[Dict[str, Any]] = []
        try:
            conn = self._get_connection()
            cur = conn.cursor()
            if self.db_type == "postgres":
                schema = schema or "public"
                cur.execute(
                    """
                    SELECT column_name, data_type, is_nullable, column_default
                    FROM information_schema.columns
                    WHERE table_schema = %s AND table_name = %s
                    ORDER BY ordinal_position
                    """,
                    (schema, table),
                )
                for name, dtype, nullable, default in self._as_rows(cur.fetchall()):
                    cols.append(
                        {
                            "name": str(name),
                            "type": str(dtype),
                            "nullable": str(nullable).upper() == "YES",
                            "default": "" if default is None else str(default),
                        }
                    )
            elif self.db_type == "mysql":
                schema = schema or (self.mysql_config or {}).get("database") or ""
                cur.execute(
                    """
                    SELECT column_name, data_type, is_nullable, column_default
                    FROM information_schema.columns
                    WHERE table_schema = %s AND table_name = %s
                    ORDER BY ordinal_position
                    """,
                    (schema, table),
                )
                for name, dtype, nullable, default in self._as_rows(cur.fetchall()):
                    cols.append(
                        {
                            "name": str(name),
                            "type": str(dtype),
                            "nullable": str(nullable).upper() == "YES",
                            "default": "" if default is None else str(default),
                        }
                    )
            else:
                cur.execute(f'PRAGMA table_info("{table}")')
                for _cid, name, dtype, notnull, default, _pk in self._as_rows(
                    cur.fetchall()
                ):
                    cols.append(
                        {
                            "name": str(name),
                            "type": str(dtype or ""),
                            "nullable": not bool(notnull),
                            "default": "" if default is None else str(default),
                        }
                    )
            cur.close()
            conn.close()
        except Exception as exc:
            self.logger.error("describe_table failed: %s", exc)
        return cols

    def preview_table(
        self,
        table: str,
        schema: Optional[str] = None,
        *,
        limit: int = 50,
    ) -> Dict[str, Any]:
        """Return column names + sample rows for the Database browser."""
        table = (table or "").strip()
        limit = max(1, min(500, int(limit or 50)))
        empty = {"columns": [], "rows": [], "error": ""}
        if not table:
            return empty
        try:
            conn = self._get_connection()
            cur = conn.cursor()
            if self.db_type == "postgres":
                schema = schema or "public"
                if pg_sql is None:
                    raise RuntimeError("psycopg2.sql unavailable")
                query = pg_sql.SQL("SELECT * FROM {}.{} LIMIT %s").format(
                    pg_sql.Identifier(schema),
                    pg_sql.Identifier(table),
                )
                cur.execute(query, (limit,))
            elif self.db_type == "mysql":
                # Identifiers from introspection only
                safe = "".join(ch if ch.isalnum() or ch == "_" else "" for ch in table)
                cur.execute(f"SELECT * FROM `{safe}` LIMIT %s", (limit,))
            else:
                safe = table.replace('"', "")
                cur.execute(f'SELECT * FROM "{safe}" LIMIT ?', (limit,))
            colnames = [d[0] for d in (cur.description or [])]
            raw_rows = self._as_rows(cur.fetchall())
            rows = [
                {colnames[i]: row[i] for i in range(len(colnames))}
                for row in raw_rows
            ]
            cur.close()
            conn.close()
            return {"columns": colnames, "rows": rows, "error": ""}
        except Exception as exc:
            self.logger.error("preview_table failed: %s", exc)
            return {"columns": [], "rows": [], "error": str(exc)}

    def get_guri_record_count(self) -> int:
        """Count rows in guri_records when the table exists."""
        try:
            return int(self.get_record_count())
        except Exception:
            return 0


def connect_guri_database(
    base_path: str = "C:/GeoFooter",
    logger: Optional[logging.Logger] = None,
    *,
    allow_legacy_fallback: bool = False,
) -> Tuple[GURIDatabase, str]:
    """
    Connect to the GURI database.

    PostgreSQL via ``guri_postgres_config.json`` is required.

    MySQL / SQLite are deprecated. They are only used when
    ``allow_legacy_fallback=True`` (emergency / migration tooling).

    Returns:
        Tuple of (database instance, db_type label: "postgres", "mysql", or "sqlite")
    """
    log = logger or logging.getLogger(__name__)
    pg_config_path = os.path.join(base_path, "guri_postgres_config.json")

    if os.path.exists(pg_config_path):
        try:
            with open(pg_config_path, "r", encoding="utf-8") as config_file:
                pg_config = json.load(config_file)

            required_fields = ["host", "user", "password", "database"]
            missing_fields = [
                field for field in required_fields if field not in pg_config
            ]
            if missing_fields:
                raise ValueError(
                    f"PostgreSQL config missing required fields: {', '.join(missing_fields)}"
                )

            if not POSTGRES_AVAILABLE:
                raise ImportError(
                    "psycopg2 is not installed "
                    "(pip install psycopg2-binary)"
                )

            if "port" not in pg_config:
                pg_config["port"] = 5432

            db = GURIDatabase(
                db_type="postgres", pg_config=pg_config, base_path=base_path
            )
            log.info(
                "Connected to GURI PostgreSQL database at %s:%s/%s",
                pg_config["host"],
                pg_config.get("port", 5432),
                pg_config["database"],
            )
            return db, "postgres"
        except json.JSONDecodeError as exc:
            log.error("Failed to parse GURI PostgreSQL config file: %s", exc)
            if not allow_legacy_fallback:
                raise
            log.warning("Falling back to deprecated MySQL/SQLite (legacy).")
        except Exception as exc:
            log.error("Failed to connect to GURI PostgreSQL: %s", exc)
            if not allow_legacy_fallback:
                raise
            log.warning("Falling back to deprecated MySQL/SQLite (legacy).")
    elif not allow_legacy_fallback:
        raise FileNotFoundError(
            f"PostgreSQL config not found: {pg_config_path}. "
            "Copy guri_postgres_config.example.json to guri_postgres_config.json."
        )

    if not allow_legacy_fallback:
        raise RuntimeError(
            "PostgreSQL is required. MySQL and SQLite auto-fallback are disabled."
        )

    # Deprecated legacy paths (explicit opt-in only)
    mysql_config_path = os.path.join(base_path, "guri_mysql_config.json")
    if os.path.exists(mysql_config_path):
        try:
            with open(mysql_config_path, "r", encoding="utf-8") as config_file:
                mysql_config = json.load(config_file)
            required_fields = ["host", "user", "password", "database"]
            missing_fields = [
                field for field in required_fields if field not in mysql_config
            ]
            if missing_fields:
                raise ValueError(
                    f"MySQL config missing required fields: {', '.join(missing_fields)}"
                )
            if not MYSQL_AVAILABLE:
                raise ImportError(
                    "mysql-connector-python is not installed "
                    "(pip install mysql-connector-python)"
                )
            if "port" not in mysql_config:
                mysql_config["port"] = 3306
            if "charset" not in mysql_config:
                mysql_config["charset"] = "utf8mb4"
            db = GURIDatabase(
                db_type="mysql", mysql_config=mysql_config, base_path=base_path
            )
            log.warning(
                "Connected to deprecated GURI MySQL at %s:%s/%s",
                mysql_config["host"],
                mysql_config.get("port", 3306),
                mysql_config["database"],
            )
            return db, "mysql"
        except Exception as exc:
            log.warning(
                "Failed to connect to deprecated MySQL: %s. Trying SQLite.", exc
            )

    db = GURIDatabase(db_type="sqlite", base_path=base_path)
    log.warning("Connected to deprecated GURI SQLite at %s", db.db_path)
    return db, "sqlite"


def print_help():
    """Display comprehensive help information about the GURI module."""
    help_text = """
╔══════════════════════════════════════════════════════════════════════════════╗
║                    GURI - Globally Unique Record Identifier                  ║
║                         Generator and Database Manager                        ║
╚══════════════════════════════════════════════════════════════════════════════╝

DESCRIPTION:
    The GURI module provides a system for generating and managing unique 
    identifiers for email analysis records. Each GURI is a unique hexadecimal
    string that can be used to track and reference email security assessments.

GURI FORMAT:
    Pattern: {5hex}x{5hex}x{8hex}x{8hex}x{3hex}x{2hex}
    Example: a1b2cx3d4e5xf6g7h8i9xj0k1l2m3xn4oxp5
    
    Where:
    - {5hex} = 5 random hexadecimal characters
    - x = separator character
    - Total length: 37 characters

DATABASE SCHEMA:
    Table: guri_records
    Fields:
    - id              : Primary key (auto-increment)
    - guri            : Unique GURI string
    - sender          : Email sender address
    - recipients      : Recipient email addresses
    - subject         : Email subject line
    - datetime        : Timestamp of the email/analysis
    - avg_risk        : Risk assessment string (e.g., "LOW (10/100)")
    - document_type   : Document type code (01=Email, 02=Word, 03=Excel, etc.)
    - created_at      : Record creation timestamp (auto-generated)

USAGE EXAMPLES:

    1. Basic SQLite Usage (Default):
    ┌──────────────────────────────────────────────────────────────────────────┐
    │ from guri import GURIDatabase                                            │
    │                                                                          │
    │ # Initialize SQLite database (default)                                   │
    │ db = GURIDatabase()                                                      │
    │                                                                          │
    │ # Generate a GURI                                                        │
    │ guri = db.generate_guri(                                                 │
    │     sender="user@example.com",                                           │
    │     recipients="recipient@example.com",                                  │
    │     subject="Important Email",                                           │
    │     dt_str="2025-10-31 12:00:00",                                        │
    │     avg_risk="LOW (10/100)"                                              │
    │ )                                                                        │
    │ print(f"Generated GURI: {guri}")                                         │
    └──────────────────────────────────────────────────────────────────────────┘

    2. PostgreSQL Usage:
    ┌──────────────────────────────────────────────────────────────────────────┐
    │ from guri import GURIDatabase                                            │
    │                                                                          │
    │ # PostgreSQL configuration                                                    │
    │ pg_config = {                                                         │
    │     "host": "localhost",                                                 │
    │     "user": "root",                                                      │
    │     "password": "your_password",                                         │
    │     "database": "guri_db",                                               │
    │     "port": 5432,  # optional, default is 3306                           │
    │     "charset": "utf8mb4"  # optional                                     │
    │ }                                                                        │
    │                                                                          │
    │ # Initialize PostgreSQL database                                              │
    │ db = GURIDatabase(db_type="postgres", pg_config=pg_config)           │
    │                                                                          │
    │ # Use the same way as SQLite                                             │
    │ guri = db.generate_guri(...)                                             │
    └──────────────────────────────────────────────────────────────────────────┘

    3. Custom SQLite Database Path:
    ┌──────────────────────────────────────────────────────────────────────────┐
    │ db = GURIDatabase(db_type="sqlite", db_path="/custom/path/my_guri.db")  │
    └──────────────────────────────────────────────────────────────────────────┘

    4. Retrieve Existing GURI:
    ┌──────────────────────────────────────────────────────────────────────────┐
    │ # If the same parameters are used, it will return the existing GURI     │
    │ existing_guri = db.get_guri(                                             │
    │     sender="user@example.com",                                           │
    │     recipients="recipient@example.com",                                  │
    │     subject="Important Email",                                           │
    │     dt_str="2025-10-31 12:00:00",                                        │
    │     avg_risk="LOW (10/100)"                                              │
    │ )                                                                        │
    │ if existing_guri:                                                        │
    │     print(f"Found existing GURI: {existing_guri}")                       │
    └──────────────────────────────────────────────────────────────────────────┘

    5. Direct Insert (Advanced):
    ┌──────────────────────────────────────────────────────────────────────────┐
    │ success = db.insert_guri(                                                │
    │     guri="a1b2cx3d4e5xf6g7h8i9xj0k1l2m3xn4oxp5",                        │
    │     sender="user@example.com",                                           │
    │     recipients="recipient@example.com",                                  │
    │     subject="Important Email",                                           │
    │     dt_str="2025-10-31 12:00:00",                                        │
    │     avg_risk="LOW (10/100)",                                             │
    │     document_type="01"  # 01=Email, 02=Word, etc.                        │
    │ )                                                                        │
    └──────────────────────────────────────────────────────────────────────────┘

COMMAND LINE USAGE:
    python guri.py              - Run example demonstration
    python guri.py --help       - Display this help message

CLASS METHODS:
    GURIDatabase.__init__(db_type="sqlite", db_path=None, base_path="C:/GeoFooter",
                         mysql_config=None)
        Initialize the GURI database (SQLite or PostgreSQL)
        
    GURIDatabase.generate_guri(sender, recipients, subject, dt_str, avg_risk)
        Generate new GURI or retrieve existing one
        Returns: GURI string
        
    GURIDatabase.get_guri(sender, recipients, subject, dt_str, avg_risk)
        Retrieve existing GURI from database
        Returns: GURI string or None
        
    GURIDatabase.insert_guri(guri, sender, recipients, subject, dt_str, 
                             avg_risk, document_type="01")
        Insert a new GURI record
        Returns: bool (success/failure)

KEY FEATURES:
    ✓ Automatic database creation and initialization
    ✓ Dual database support (SQLite and PostgreSQL)
    ✓ Duplicate prevention (returns existing GURI for same parameters)
    ✓ Unique constraint on GURI field
    ✓ Comprehensive logging
    ✓ Thread-safe database operations
    ✓ Configurable database location
    ✓ Automatic indexing for performance
    ✓ Connection support (PostgreSQL)

DEPENDENCIES:
    Required:
    - sqlite3 (built-in)
    - random (built-in)
    - logging (built-in)
    - os (built-in)
    - typing (built-in)
    
    Optional (for PostgreSQL support):
    - psycopg2-binary (install with: pip install psycopg2-binary)

VERSION: 1.0.0
AUTHOR: Extracted from geolocate_headers.py
LICENSE: Use as needed

For more information or support, refer to the main documentation.
"""
    print(help_text)


def main():
    """Example usage of the GURI database."""
    import sys
    
    # Check for help flag
    if len(sys.argv) > 1 and sys.argv[1] in ['--help', '-h', 'help']:
        print_help()
        return
    
    # Setup basic logging
    logging.basicConfig(
        level=logging.INFO,
        format='%(asctime)s - %(name)s - %(levelname)s - %(message)s'
    )
    
    print("=" * 80)
    print("GURI Database - Example Demonstration")
    print("=" * 80)
    print()
    
    # Initialize database
    print("1. Initializing GURI database...")
    db = GURIDatabase()
    print(f"   ✓ Database initialized at: {db.db_path}")
    print()
    
    # Example: Generate a GURI
    print("2. Generating a new GURI...")
    sender = "example@example.com"
    recipients = "recipient@example.com"
    subject = "Test Email"
    dt_str = "2025-10-31 12:00:00"
    avg_risk = "LOW (10/100)"
    
    guri = db.generate_guri(sender, recipients, subject, dt_str, avg_risk)
    print(f"   ✓ Generated GURI: {guri}")
    print()
    
    # Try to retrieve the same GURI (should return the same one)
    print("3. Attempting to retrieve existing GURI with same parameters...")
    guri2 = db.generate_guri(sender, recipients, subject, dt_str, avg_risk)
    print(f"   ✓ Retrieved GURI: {guri2}")
    print(f"   ✓ GURIs match: {guri == guri2}")
    print()
    
    # Generate a different GURI with different parameters
    print("4. Generating a different GURI with different parameters...")
    guri3 = db.generate_guri(
        sender="different@example.com",
        recipients=recipients,
        subject=subject,
        dt_str=dt_str,
        avg_risk="HIGH (85/100)"
    )
    print(f"   ✓ Generated new GURI: {guri3}")
    print(f"   ✓ Different from first: {guri != guri3}")
    print()
    
    print("=" * 80)
    print("Demonstration complete!")
    print("Run 'python guri.py --help' for detailed usage information.")
    print("=" * 80)


if __name__ == "__main__":
    main()

