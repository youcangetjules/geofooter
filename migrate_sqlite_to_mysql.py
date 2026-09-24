#!/usr/bin/env python3
"""
SQLite to MySQL Migration Tool for GURI Database

This module migrates all records from a SQLite database to a MySQL database,
including both guri_records and supplier_codes tables.

Usage:
    python migrate_sqlite_to_mysql.py
    
Or import and use programmatically:
    from migrate_sqlite_to_mysql import migrate_database
    migrate_database(sqlite_path, mysql_config, skip_duplicates=True)
"""

import os
import sys
import sqlite3
import logging
from typing import Dict, Any, Optional
from datetime import datetime

# Try to import MySQL connector
try:
    import mysql.connector
    from mysql.connector import Error as MySQLError
    MYSQL_AVAILABLE = True
except ImportError:
    MYSQL_AVAILABLE = False
    MySQLError = Exception


def setup_logging():
    """Setup logging configuration."""
    logging.basicConfig(
        level=logging.INFO,
        format='%(asctime)s - %(levelname)s - %(message)s',
        datefmt='%Y-%m-%d %H:%M:%S'
    )
    return logging.getLogger(__name__)


def migrate_guri_records(sqlite_conn, mysql_conn, skip_duplicates: bool = True, logger=None):
    """
    Migrate guri_records from SQLite to MySQL.
    
    Args:
        sqlite_conn: SQLite connection object
        mysql_conn: MySQL connection object
        skip_duplicates: If True, skip records that already exist (based on GURI)
        logger: Logger instance
        
    Returns:
        Tuple of (total_count, migrated_count, skipped_count, error_count)
    """
    if logger is None:
        logger = logging.getLogger(__name__)
    
    sqlite_cursor = sqlite_conn.cursor()
    mysql_cursor = mysql_conn.cursor()
    
    try:
        # Get total count from SQLite
        sqlite_cursor.execute("SELECT COUNT(*) FROM guri_records")
        total_count = sqlite_cursor.fetchone()[0]
        logger.info(f"Found {total_count} GURI records in SQLite database")
        
        if total_count == 0:
            logger.info("No records to migrate")
            return (0, 0, 0, 0)
        
        # Check which columns exist in SQLite
        sqlite_cursor.execute("PRAGMA table_info(guri_records)")
        columns_info = sqlite_cursor.fetchall()
        column_names = [col[1] for col in columns_info]
        
        # Build SELECT query based on available columns
        base_columns = ['guri', 'sender', 'recipients', 'subject', 'datetime', 'avg_risk']
        select_columns = base_columns.copy()
        
        if 'document_type' in column_names:
            select_columns.append('document_type')
        else:
            select_columns.append("'01' as document_type")
        
        if 'sensitivity' in column_names:
            select_columns.append('sensitivity')
        else:
            select_columns.append("'[SEC1:(U)EXTERNAL/UNRATED]' as sensitivity")
        
        if 'company_name' in column_names:
            select_columns.append('company_name')
        else:
            select_columns.append('NULL as company_name')
        
        if 'created_at' in column_names:
            select_columns.append('created_at')
        else:
            select_columns.append("datetime('now') as created_at")
        
        query = f'''
            SELECT {', '.join(select_columns)}
            FROM guri_records
            ORDER BY created_at
        '''
        
        # Fetch all records from SQLite
        sqlite_cursor.execute(query)
        
        migrated_count = 0
        skipped_count = 0
        error_count = 0
        
        # Process records in batches
        batch_size = 100
        batch = []
        
        for row in sqlite_cursor.fetchall():
            # Unpack row based on available columns
            guri = row[0]
            sender = row[1]
            recipients = row[2]
            subject = row[3]
            dt_str = row[4]
            avg_risk = row[5]
            
            # Handle optional columns
            idx = 6
            document_type = row[idx] if idx < len(row) else '01'
            idx += 1
            sensitivity = row[idx] if idx < len(row) else '[SEC1:(U)EXTERNAL/UNRATED]'
            idx += 1
            company_name = row[idx] if idx < len(row) else None
            idx += 1
            created_at = row[idx] if idx < len(row) else None
            
            # Handle None values
            document_type = document_type or '01'
            sensitivity = sensitivity or '[SEC1:(U)EXTERNAL/UNRATED]'
            company_name = company_name if company_name else None
            
            batch.append((guri, sender, recipients, subject, dt_str, avg_risk, 
                         document_type, sensitivity, company_name, created_at))
            
            if len(batch) >= batch_size:
                result = _insert_batch(mysql_cursor, batch, skip_duplicates, logger)
                migrated_count += result[0]
                skipped_count += result[1]
                error_count += result[2]
                batch = []
                mysql_conn.commit()
        
        # Insert remaining records
        if batch:
            result = _insert_batch(mysql_cursor, batch, skip_duplicates, logger)
            migrated_count += result[0]
            skipped_count += result[1]
            error_count += result[2]
            mysql_conn.commit()
        
        logger.info(f"GURI Records Migration Complete:")
        logger.info(f"  Total: {total_count}")
        logger.info(f"  Migrated: {migrated_count}")
        logger.info(f"  Skipped (duplicates): {skipped_count}")
        logger.info(f"  Errors: {error_count}")
        
        return (total_count, migrated_count, skipped_count, error_count)
        
    except Exception as e:
        logger.error(f"Error migrating GURI records: {e}")
        mysql_conn.rollback()
        raise
    finally:
        sqlite_cursor.close()


def _insert_batch(mysql_cursor, batch, skip_duplicates, logger):
    """Insert a batch of records into MySQL."""
    migrated = 0
    skipped = 0
    errors = 0
    
    if skip_duplicates:
        query = '''
            INSERT IGNORE INTO guri_records 
            (guri, sender, recipients, subject, datetime, avg_risk, 
             document_type, sensitivity, company_name, created_at)
            VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s, %s)
        '''
    else:
        query = '''
            INSERT INTO guri_records 
            (guri, sender, recipients, subject, datetime, avg_risk, 
             document_type, sensitivity, company_name, created_at)
            VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s, %s)
            ON DUPLICATE KEY UPDATE
                sender = VALUES(sender),
                recipients = VALUES(recipients),
                subject = VALUES(subject),
                datetime = VALUES(datetime),
                avg_risk = VALUES(avg_risk),
                document_type = VALUES(document_type),
                sensitivity = VALUES(sensitivity),
                company_name = VALUES(company_name)
        '''
    
    try:
        mysql_cursor.executemany(query, batch)
        migrated = mysql_cursor.rowcount
        if skip_duplicates and migrated < len(batch):
            skipped = len(batch) - migrated
    except Exception as e:
        logger.error(f"Error inserting batch: {e}")
        errors = len(batch)
    
    return (migrated, skipped, errors)


def migrate_supplier_codes(sqlite_conn, mysql_conn, skip_duplicates: bool = True, logger=None):
    """
    Migrate supplier_codes from SQLite to MySQL.
    
    Args:
        sqlite_conn: SQLite connection object
        mysql_conn: MySQL connection object
        skip_duplicates: If True, skip records that already exist (based on company_name)
        logger: Logger instance
        
    Returns:
        Tuple of (total_count, migrated_count, skipped_count, error_count)
    """
    if logger is None:
        logger = logging.getLogger(__name__)
    
    sqlite_cursor = sqlite_conn.cursor()
    mysql_cursor = mysql_conn.cursor()
    
    try:
        # Check if supplier_codes table exists in SQLite
        sqlite_cursor.execute('''
            SELECT name FROM sqlite_master 
            WHERE type='table' AND name='supplier_codes'
        ''')
        if not sqlite_cursor.fetchone():
            logger.info("supplier_codes table does not exist in SQLite, skipping")
            return (0, 0, 0, 0)
        
        # Get total count from SQLite
        sqlite_cursor.execute("SELECT COUNT(*) FROM supplier_codes")
        total_count = sqlite_cursor.fetchone()[0]
        logger.info(f"Found {total_count} supplier records in SQLite database")
        
        if total_count == 0:
            logger.info("No supplier records to migrate")
            return (0, 0, 0, 0)
        
        # Fetch all records from SQLite
        sqlite_cursor.execute('''
            SELECT company_name, company_domain, supplier_code, contract_code,
                   contact_person, contact_email, notes, created_at, updated_at
            FROM supplier_codes
            ORDER BY created_at
        ''')
        
        migrated_count = 0
        skipped_count = 0
        error_count = 0
        
        # Process records
        for row in sqlite_cursor.fetchall():
            company_name, company_domain, supplier_code, contract_code, \
            contact_person, contact_email, notes, created_at, updated_at = row
            
            # Handle None values
            company_domain = company_domain or None
            supplier_code = supplier_code or None
            contract_code = contract_code or None
            contact_person = contact_person or None
            contact_email = contact_email or None
            notes = notes or None
            updated_at = updated_at or created_at
            
            try:
                if skip_duplicates:
                    query = '''
                        INSERT IGNORE INTO supplier_codes
                        (company_name, company_domain, supplier_code, contract_code,
                         contact_person, contact_email, notes, created_at, updated_at)
                        VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s)
                    '''
                else:
                    query = '''
                        INSERT INTO supplier_codes
                        (company_name, company_domain, supplier_code, contract_code,
                         contact_person, contact_email, notes, created_at, updated_at)
                        VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s)
                        ON DUPLICATE KEY UPDATE
                            company_domain = VALUES(company_domain),
                            supplier_code = VALUES(supplier_code),
                            contract_code = VALUES(contract_code),
                            contact_person = VALUES(contact_person),
                            contact_email = VALUES(contact_email),
                            notes = VALUES(notes),
                            updated_at = VALUES(updated_at)
                    '''
                
                mysql_cursor.execute(query, (
                    company_name, company_domain, supplier_code, contract_code,
                    contact_person, contact_email, notes, created_at, updated_at
                ))
                
                if mysql_cursor.rowcount > 0:
                    migrated_count += 1
                else:
                    skipped_count += 1
                    
            except Exception as e:
                logger.error(f"Error inserting supplier record '{company_name}': {e}")
                error_count += 1
        
        mysql_conn.commit()
        
        logger.info(f"Supplier Codes Migration Complete:")
        logger.info(f"  Total: {total_count}")
        logger.info(f"  Migrated: {migrated_count}")
        logger.info(f"  Skipped (duplicates): {skipped_count}")
        logger.info(f"  Errors: {error_count}")
        
        return (total_count, migrated_count, skipped_count, error_count)
        
    except Exception as e:
        logger.error(f"Error migrating supplier codes: {e}")
        mysql_conn.rollback()
        raise
    finally:
        sqlite_cursor.close()


def migrate_database(sqlite_path: str, 
                    mysql_config: Dict[str, Any],
                    skip_duplicates: bool = True,
                    logger: Optional[logging.Logger] = None) -> Dict[str, Any]:
    """
    Migrate all data from SQLite to MySQL database.
    
    Args:
        sqlite_path: Path to SQLite database file
        mysql_config: MySQL connection configuration dict with keys:
                     host, user, password, database (optional: port, charset)
        skip_duplicates: If True, skip records that already exist
        logger: Optional logger instance
        
    Returns:
        Dictionary with migration statistics
        
    Example:
        mysql_config = {
            "host": "localhost",
            "user": "root",
            "password": "your_password",
            "database": "guri_db",
            "port": 3306
        }
        result = migrate_database("guri_records.db", mysql_config)
    """
    if logger is None:
        logger = setup_logging()
    
    if not MYSQL_AVAILABLE:
        raise ImportError(
            "MySQL support requires mysql-connector-python. "
            "Install with: pip install mysql-connector-python"
        )
    
    if not os.path.exists(sqlite_path):
        raise FileNotFoundError(f"SQLite database not found: {sqlite_path}")
    
    logger.info("=" * 60)
    logger.info("Starting SQLite to MySQL Migration")
    logger.info("=" * 60)
    logger.info(f"SQLite Database: {sqlite_path}")
    logger.info(f"MySQL Database: {mysql_config.get('host')}:{mysql_config.get('port', 3306)}/{mysql_config.get('database')}")
    logger.info(f"Skip Duplicates: {skip_duplicates}")
    logger.info("")
    
    sqlite_conn = None
    mysql_conn = None
    
    try:
        # Connect to SQLite
        logger.info("Connecting to SQLite database...")
        sqlite_conn = sqlite3.connect(sqlite_path)
        logger.info("✓ Connected to SQLite")
        
        # Connect to MySQL
        logger.info("Connecting to MySQL database...")
        mysql_conn = mysql.connector.connect(**mysql_config)
        logger.info("✓ Connected to MySQL")
        
        # Ensure MySQL tables exist (they should, but just in case)
        logger.info("Ensuring MySQL tables exist...")
        mysql_cursor = mysql_conn.cursor()
        
        # Create guri_records table if needed
        mysql_cursor.execute('''
            CREATE TABLE IF NOT EXISTS guri_records (
                id INT AUTO_INCREMENT PRIMARY KEY,
                guri VARCHAR(50) UNIQUE NOT NULL,
                sender VARCHAR(255) NOT NULL,
                recipients TEXT,
                subject TEXT,
                datetime VARCHAR(50),
                avg_risk VARCHAR(50),
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
        
        # Create supplier_codes table if needed
        mysql_cursor.execute('''
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
        
        mysql_conn.commit()
        mysql_cursor.close()
        logger.info("✓ MySQL tables verified/created")
        logger.info("")
        
        # Migrate GURI records
        logger.info("Migrating GURI records...")
        guri_stats = migrate_guri_records(sqlite_conn, mysql_conn, skip_duplicates, logger)
        logger.info("")
        
        # Migrate supplier codes
        logger.info("Migrating supplier codes...")
        supplier_stats = migrate_supplier_codes(sqlite_conn, mysql_conn, skip_duplicates, logger)
        logger.info("")
        
        # Summary
        logger.info("=" * 60)
        logger.info("Migration Complete!")
        logger.info("=" * 60)
        logger.info(f"GURI Records: {guri_stats[1]} migrated, {guri_stats[2]} skipped, {guri_stats[3]} errors")
        logger.info(f"Supplier Codes: {supplier_stats[1]} migrated, {supplier_stats[2]} skipped, {supplier_stats[3]} errors")
        logger.info("")
        
        return {
            'guri_records': {
                'total': guri_stats[0],
                'migrated': guri_stats[1],
                'skipped': guri_stats[2],
                'errors': guri_stats[3]
            },
            'supplier_codes': {
                'total': supplier_stats[0],
                'migrated': supplier_stats[1],
                'skipped': supplier_stats[2],
                'errors': supplier_stats[3]
            },
            'success': True
        }
        
    except Exception as e:
        logger.error(f"Migration failed: {e}")
        if mysql_conn:
            mysql_conn.rollback()
        return {
            'success': False,
            'error': str(e)
        }
        
    finally:
        if sqlite_conn:
            sqlite_conn.close()
        if mysql_conn:
            mysql_conn.close()
        logger.info("Database connections closed")


def main():
    """Main function for command-line usage."""
    logger = setup_logging()
    
    print("=" * 60)
    print("GURI Database Migration Tool")
    print("SQLite to MySQL Migration")
    print("=" * 60)
    print()
    
    # Get SQLite path
    default_sqlite = os.path.join("C:/GeoFooter", "guri_records.db")
    sqlite_path = input(f"SQLite database path [{default_sqlite}]: ").strip()
    if not sqlite_path:
        sqlite_path = default_sqlite
    
    if not os.path.exists(sqlite_path):
        print(f"Error: SQLite database not found: {sqlite_path}")
        sys.exit(1)
    
    # Get MySQL configuration
    print()
    print("MySQL Connection Configuration:")
    mysql_config = {}
    mysql_config['host'] = input("Host [localhost]: ").strip() or "localhost"
    mysql_config['port'] = int(input("Port [3306]: ").strip() or "3306")
    mysql_config['user'] = input("User [root]: ").strip() or "root"
    mysql_config['password'] = input("Password: ").strip()
    mysql_config['database'] = input("Database [guri_db]: ").strip() or "guri_db"
    mysql_config['charset'] = 'utf8mb4'
    
    # Ask about duplicates
    print()
    skip_duplicates_input = input("Skip duplicate records? [Y/n]: ").strip().lower()
    skip_duplicates = skip_duplicates_input != 'n'
    
    print()
    print("Starting migration...")
    print()
    
    # Run migration
    result = migrate_database(sqlite_path, mysql_config, skip_duplicates, logger)
    
    if result['success']:
        print()
        print("✓ Migration completed successfully!")
        sys.exit(0)
    else:
        print()
        print(f"✗ Migration failed: {result.get('error', 'Unknown error')}")
        sys.exit(1)


if __name__ == "__main__":
    main()

