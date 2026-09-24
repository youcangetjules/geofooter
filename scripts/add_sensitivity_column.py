#!/usr/bin/env python3
"""
Database Migration Script: Add Sensitivity Column to GURI Records
Adds sensitivity/security classification column to existing GURI databases.
"""

import sqlite3
import os
import logging

logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

def migrate_sqlite_database(db_path="C:/GeoFooter/guri_records.db"):
    """Add sensitivity column to SQLite database."""
    if not os.path.exists(db_path):
        logger.warning(f"Database not found: {db_path}")
        return False
    
    try:
        conn = sqlite3.connect(db_path)
        cursor = conn.cursor()
        
        # Check if column already exists
        cursor.execute("PRAGMA table_info(guri_records)")
        columns = [col[1] for col in cursor.fetchall()]
        
        if 'sensitivity' in columns:
            logger.info("Sensitivity column already exists in database")
            return True
        
        # Add sensitivity column with default value
        logger.info("Adding sensitivity column to guri_records table...")
        cursor.execute('''
            ALTER TABLE guri_records 
            ADD COLUMN sensitivity TEXT DEFAULT '[SEC1:(U)EXTERNAL/UNRATED]'
        ''')
        
        conn.commit()
        cursor.close()
        conn.close()
        
        logger.info("✓ Successfully added sensitivity column to SQLite database")
        return True
        
    except Exception as e:
        logger.error(f"Failed to migrate database: {e}")
        return False

def migrate_mysql_database(host, user, password, database, port=3306):
    """Add sensitivity column to MySQL database."""
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
        
        # Check if column already exists
        cursor.execute(f"""
            SELECT COUNT(*) 
            FROM information_schema.COLUMNS 
            WHERE TABLE_SCHEMA = '{database}' 
            AND TABLE_NAME = 'guri_records' 
            AND COLUMN_NAME = 'sensitivity'
        """)
        
        if cursor.fetchone()[0] > 0:
            logger.info("Sensitivity column already exists in database")
            return True
        
        # Add sensitivity column
        logger.info("Adding sensitivity column to guri_records table...")
        cursor.execute('''
            ALTER TABLE guri_records 
            ADD COLUMN sensitivity VARCHAR(100) DEFAULT '[SEC1:(U)EXTERNAL/UNRATED]'
        ''')
        
        conn.commit()
        cursor.close()
        conn.close()
        
        logger.info("✓ Successfully added sensitivity column to MySQL database")
        return True
        
    except ImportError:
        logger.error("MySQL connector not available. Install with: pip install mysql-connector-python")
        return False
    except Exception as e:
        logger.error(f"Failed to migrate MySQL database: {e}")
        return False

if __name__ == "__main__":
    print("=" * 60)
    print("GURI Database Migration: Add Sensitivity Column")
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

