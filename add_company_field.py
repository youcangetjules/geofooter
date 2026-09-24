#!/usr/bin/env python3
"""
Database Migration Script: Add Company Name Field to GURI Records
Adds company_name column and creates supplier/contract codes table.
"""

import sqlite3
import os
import logging

logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

def migrate_sqlite_database(db_path="C:/GeoFooter/guri_records.db"):
    """Add company_name column and create supplier codes table."""
    if not os.path.exists(db_path):
        logger.warning(f"Database not found: {db_path}")
        return False
    
    try:
        conn = sqlite3.connect(db_path)
        cursor = conn.cursor()
        
        # Check if company_name column already exists
        cursor.execute("PRAGMA table_info(guri_records)")
        columns = [col[1] for col in cursor.fetchall()]
        
        if 'company_name' not in columns:
            logger.info("Adding company_name column to guri_records table...")
            cursor.execute('''
                ALTER TABLE guri_records 
                ADD COLUMN company_name TEXT
            ''')
            logger.info("✓ Added company_name column")
        else:
            logger.info("company_name column already exists")
        
        # Create supplier_codes table
        logger.info("Creating supplier_codes table...")
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
        
        cursor.execute('CREATE INDEX IF NOT EXISTS idx_company_name ON supplier_codes(company_name)')
        cursor.execute('CREATE INDEX IF NOT EXISTS idx_company_domain ON supplier_codes(company_domain)')
        
        conn.commit()
        cursor.close()
        conn.close()
        
        logger.info("✓ Successfully migrated SQLite database")
        return True
        
    except Exception as e:
        logger.error(f"Failed to migrate database: {e}")
        return False

def migrate_mysql_database(host, user, password, database, port=3306):
    """Add company_name column and create supplier codes table for MySQL."""
    try:
        import mysql.connector
        
        conn = mysql.connector.connect(
            host=host,
            user=user,
            password=password,
            database=database,
            port=port
        )
        
        cursor = conn.cursor()
        
        # Check if company_name column already exists
        cursor.execute(f"""
            SELECT COUNT(*) 
            FROM information_schema.COLUMNS 
            WHERE TABLE_SCHEMA = '{database}' 
            AND TABLE_NAME = 'guri_records' 
            AND COLUMN_NAME = 'company_name'
        """)
        
        if cursor.fetchone()[0] == 0:
            logger.info("Adding company_name column to guri_records table...")
            cursor.execute('''
                ALTER TABLE guri_records 
                ADD COLUMN company_name VARCHAR(255)
            ''')
            logger.info("✓ Added company_name column")
        else:
            logger.info("company_name column already exists")
        
        # Create supplier_codes table
        logger.info("Creating supplier_codes table...")
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
        
        logger.info("✓ Successfully migrated MySQL database")
        return True
        
    except ImportError:
        logger.error("MySQL connector not available. Install with: pip install mysql-connector-python")
        return False
    except Exception as e:
        logger.error(f"Failed to migrate MySQL database: {e}")
        return False

if __name__ == "__main__":
    print("=" * 60)
    print("GURI Database Migration: Add Company Name & Supplier Codes")
    print("=" * 60)
    print()
    
    # Migrate SQLite database
    print("Migrating SQLite database...")
    migrate_sqlite_database()
    print()
    
    # Optionally migrate MySQL
    migrate_mysql = input("Do you want to migrate a MySQL database? (y/n): ").lower()
    if migrate_mysql == 'y':
        print("\nMySQL Connection Details:")
        host = input("Host (default: localhost): ") or "localhost"
        port = input("Port (default: 3306): ") or "3306"
        user = input("User (default: root): ") or "root"
        password = input("Password: ")
        database = input("Database (default: guri_db): ") or "guri_db"
        
        migrate_mysql_database(host, user, password, database, int(port))
    
    print("\n" + "=" * 60)
    print("Migration complete!")
    print("=" * 60)

