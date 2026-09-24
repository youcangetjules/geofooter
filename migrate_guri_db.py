#!/usr/bin/env python3
"""
GURI Database Migration Script
Migrates existing GURI databases from random_block to document_type column.
"""

import sqlite3
import os
import sys
import logging
from datetime import datetime

logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)


def backup_database(db_path):
    """Create a backup of the database before migration."""
    if not os.path.exists(db_path):
        logger.error(f"Database file not found: {db_path}")
        return None
    
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    backup_path = f"{db_path}.backup_{timestamp}"
    
    try:
        import shutil
        shutil.copy2(db_path, backup_path)
        logger.info(f"✓ Backup created: {backup_path}")
        return backup_path
    except Exception as e:
        logger.error(f"Failed to create backup: {e}")
        return None


def check_column_exists(cursor, table_name, column_name):
    """Check if a column exists in a table."""
    cursor.execute(f"PRAGMA table_info({table_name})")
    columns = [row[1] for row in cursor.fetchall()]
    return column_name in columns


def migrate_sqlite_database(db_path):
    """Migrate SQLite database from random_block to document_type."""
    
    print("=" * 80)
    print("GURI Database Migration Tool")
    print("=" * 80)
    print()
    print(f"Database: {db_path}")
    print()
    
    # Check if database exists
    if not os.path.exists(db_path):
        logger.error(f"Database file not found: {db_path}")
        print("\n❌ Migration failed: Database file not found")
        return False
    
    # Create backup
    print("Step 1: Creating backup...")
    backup_path = backup_database(db_path)
    if not backup_path:
        print("\n❌ Migration aborted: Could not create backup")
        return False
    
    try:
        # Connect to database
        print("\nStep 2: Connecting to database...")
        conn = sqlite3.connect(db_path)
        cursor = conn.cursor()
        logger.info("✓ Connected to database")
        
        # Check current schema
        print("\nStep 3: Checking current schema...")
        has_random_block = check_column_exists(cursor, 'guri_records', 'random_block')
        has_document_type = check_column_exists(cursor, 'guri_records', 'document_type')
        
        logger.info(f"  - random_block column exists: {has_random_block}")
        logger.info(f"  - document_type column exists: {has_document_type}")
        
        if has_document_type and not has_random_block:
            print("\n✓ Database is already up to date!")
            print("  The document_type column exists.")
            conn.close()
            return True
        
        if has_document_type and has_random_block:
            print("\n⚠ Warning: Both columns exist!")
            print("  This database appears to be in a transitional state.")
            response = input("  Drop the old random_block column? (yes/no): ")
            if response.lower() == 'yes':
                print("\nStep 4: Dropping random_block column...")
                # SQLite doesn't support DROP COLUMN directly, need to recreate table
                cursor.execute('''
                    CREATE TABLE guri_records_new (
                        id INTEGER PRIMARY KEY AUTOINCREMENT,
                        guri TEXT UNIQUE NOT NULL,
                        sender TEXT NOT NULL,
                        recipients TEXT,
                        subject TEXT,
                        datetime TEXT,
                        avg_risk TEXT,
                        document_type TEXT DEFAULT '01',
                        created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
                    )
                ''')
                
                cursor.execute('''
                    INSERT INTO guri_records_new 
                    (id, guri, sender, recipients, subject, datetime, avg_risk, document_type, created_at)
                    SELECT id, guri, sender, recipients, subject, datetime, avg_risk, document_type, created_at
                    FROM guri_records
                ''')
                
                cursor.execute('DROP TABLE guri_records')
                cursor.execute('ALTER TABLE guri_records_new RENAME TO guri_records')
                
                # Recreate indexes
                cursor.execute('CREATE INDEX IF NOT EXISTS idx_guri ON guri_records(guri)')
                cursor.execute('CREATE INDEX IF NOT EXISTS idx_sender ON guri_records(sender)')
                
                conn.commit()
                logger.info("✓ Dropped random_block column")
            conn.close()
            return True
        
        # Migration needed
        print("\nStep 4: Migrating database schema...")
        print("  Adding document_type column...")
        
        # Add document_type column
        cursor.execute('''
            ALTER TABLE guri_records 
            ADD COLUMN document_type TEXT DEFAULT '01'
        ''')
        logger.info("✓ Added document_type column")
        
        # Migrate data from random_block to document_type if it contains valid codes
        print("\nStep 5: Migrating data...")
        cursor.execute('SELECT COUNT(*) FROM guri_records')
        total_records = cursor.fetchone()[0]
        
        if total_records > 0:
            # Update document_type from random_block where applicable
            cursor.execute('''
                UPDATE guri_records 
                SET document_type = random_block 
                WHERE random_block IN ('01', '02', '03', '04', '05', '06', '07', '08', '09', '10', '99')
            ''')
            migrated = cursor.rowcount
            logger.info(f"✓ Migrated {migrated} records with valid document type codes")
            
            # Set remaining to '01' (Email) as default
            cursor.execute('''
                UPDATE guri_records 
                SET document_type = '01' 
                WHERE document_type IS NULL OR document_type = ''
            ''')
            defaulted = cursor.rowcount
            if defaulted > 0:
                logger.info(f"✓ Set {defaulted} records to default document type '01' (Email)")
        
        conn.commit()
        
        # Drop random_block column
        print("\nStep 6: Removing old random_block column...")
        
        # SQLite doesn't support DROP COLUMN directly before version 3.35.0
        # We need to recreate the table
        cursor.execute('''
            CREATE TABLE guri_records_new (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                guri TEXT UNIQUE NOT NULL,
                sender TEXT NOT NULL,
                recipients TEXT,
                subject TEXT,
                datetime TEXT,
                avg_risk TEXT,
                document_type TEXT DEFAULT '01',
                created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
            )
        ''')
        
        cursor.execute('''
            INSERT INTO guri_records_new 
            (id, guri, sender, recipients, subject, datetime, avg_risk, document_type, created_at)
            SELECT id, guri, sender, recipients, subject, datetime, avg_risk, document_type, created_at
            FROM guri_records
        ''')
        
        cursor.execute('DROP TABLE guri_records')
        cursor.execute('ALTER TABLE guri_records_new RENAME TO guri_records')
        
        # Recreate indexes
        print("\nStep 7: Recreating indexes...")
        cursor.execute('CREATE INDEX IF NOT EXISTS idx_guri ON guri_records(guri)')
        cursor.execute('CREATE INDEX IF NOT EXISTS idx_sender ON guri_records(sender)')
        logger.info("✓ Indexes recreated")
        
        conn.commit()
        conn.close()
        
        print("\n" + "=" * 80)
        print("✓ Migration completed successfully!")
        print("=" * 80)
        print(f"\nBackup saved at: {backup_path}")
        print(f"Total records migrated: {total_records}")
        print("\nYou can now use the GURI GUI with the updated database.")
        print()
        
        return True
        
    except Exception as e:
        logger.error(f"Migration failed: {e}")
        print(f"\n❌ Migration failed: {e}")
        print(f"\nYour original database has been backed up to:")
        print(f"  {backup_path}")
        print("\nYou can restore it by renaming the backup file.")
        return False


def migrate_mysql_database(mysql_config):
    """Migrate MySQL database from random_block to document_type."""
    
    try:
        import mysql.connector
        from mysql.connector import Error as MySQLError
    except ImportError:
        print("❌ MySQL connector not installed. Run: pip install mysql-connector-python")
        return False
    
    print("=" * 80)
    print("GURI MySQL Database Migration Tool")
    print("=" * 80)
    print()
    
    try:
        print("Step 1: Connecting to MySQL...")
        conn = mysql.connector.connect(**mysql_config)
        cursor = conn.cursor()
        logger.info("✓ Connected to MySQL database")
        
        print("\nStep 2: Checking current schema...")
        
        # Check if columns exist
        cursor.execute("""
            SELECT COLUMN_NAME 
            FROM INFORMATION_SCHEMA.COLUMNS 
            WHERE TABLE_SCHEMA = %s AND TABLE_NAME = 'guri_records'
        """, (mysql_config['database'],))
        
        columns = [row[0] for row in cursor.fetchall()]
        has_random_block = 'random_block' in columns
        has_document_type = 'document_type' in columns
        
        logger.info(f"  - random_block column exists: {has_random_block}")
        logger.info(f"  - document_type column exists: {has_document_type}")
        
        if has_document_type and not has_random_block:
            print("\n✓ Database is already up to date!")
            cursor.close()
            conn.close()
            return True
        
        if not has_document_type:
            print("\nStep 3: Adding document_type column...")
            cursor.execute('''
                ALTER TABLE guri_records 
                ADD COLUMN document_type VARCHAR(10) DEFAULT '01'
            ''')
            logger.info("✓ Added document_type column")
        
        if has_random_block:
            print("\nStep 4: Migrating data...")
            cursor.execute('''
                UPDATE guri_records 
                SET document_type = random_block 
                WHERE random_block IN ('01', '02', '03', '04', '05', '06', '07', '08', '09', '10', '99')
            ''')
            migrated = cursor.rowcount
            logger.info(f"✓ Migrated {migrated} records")
            
            print("\nStep 5: Dropping old random_block column...")
            cursor.execute('ALTER TABLE guri_records DROP COLUMN random_block')
            logger.info("✓ Dropped random_block column")
        
        conn.commit()
        cursor.close()
        conn.close()
        
        print("\n" + "=" * 80)
        print("✓ MySQL migration completed successfully!")
        print("=" * 80)
        print()
        
        return True
        
    except Exception as e:
        logger.error(f"MySQL migration failed: {e}")
        print(f"\n❌ Migration failed: {e}")
        return False


def main():
    """Main migration function."""
    
    if len(sys.argv) > 1:
        db_path = sys.argv[1]
    else:
        db_path = "C:/GeoFooter/guri_records.db"
    
    print()
    print("This tool will migrate your GURI database from the old schema")
    print("(with random_block) to the new schema (with document_type).")
    print()
    
    if db_path.endswith('.db'):
        # SQLite migration
        response = input(f"Migrate SQLite database: {db_path}? (yes/no): ")
        if response.lower() == 'yes':
            success = migrate_sqlite_database(db_path)
            sys.exit(0 if success else 1)
        else:
            print("Migration cancelled.")
            sys.exit(0)
    else:
        print("For MySQL migration, edit this script with your MySQL credentials.")
        print()
        mysql_config = {
            "host": "localhost",
            "user": "root",
            "password": "your_password",
            "database": "guri_db",
            "port": 3306
        }
        print("Or call migrate_mysql_database(mysql_config) programmatically.")
        sys.exit(1)


if __name__ == "__main__":
    main()

