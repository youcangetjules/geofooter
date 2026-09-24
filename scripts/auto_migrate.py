#!/usr/bin/env python3
"""
Automatic GURI Database Migration
Runs without user interaction - updates database automatically.
"""

import sqlite3
import os
import shutil
from datetime import datetime

def migrate_database(db_path="C:/GeoFooter/guri_records.db"):
    """Automatically migrate the database."""
    
    print("=" * 80)
    print("GURI Database Auto-Migration")
    print("=" * 80)
    print()
    
    if not os.path.exists(db_path):
        print(f"[OK] No existing database found at {db_path}")
        print("     A new database will be created when you first use the GUI.")
        return True
    
    # Create backup
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    backup_path = f"{db_path}.backup_{timestamp}"
    
    try:
        shutil.copy2(db_path, backup_path)
        print(f"[OK] Backup created: {backup_path}")
    except Exception as e:
        print(f"[ERROR] Failed to create backup: {e}")
        return False
    
    try:
        conn = sqlite3.connect(db_path)
        cursor = conn.cursor()
        
        # Check if migration is needed
        cursor.execute("PRAGMA table_info(guri_records)")
        columns = [row[1] for row in cursor.fetchall()]
        
        has_random_block = 'random_block' in columns
        has_document_type = 'document_type' in columns
        
        if has_document_type and not has_random_block:
            print("[OK] Database is already up to date!")
            conn.close()
            return True
        
        print("\nMigrating database schema...")
        
        # Create new table with correct schema
        cursor.execute('''
            CREATE TABLE IF NOT EXISTS guri_records_new (
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
        
        # Copy data, migrating random_block to document_type
        has_created_at = 'created_at' in columns
        
        if has_random_block:
            if has_created_at:
                cursor.execute('''
                    INSERT INTO guri_records_new 
                    (id, guri, sender, recipients, subject, datetime, avg_risk, document_type, created_at)
                    SELECT id, guri, sender, recipients, subject, datetime, avg_risk, 
                           CASE 
                               WHEN random_block IN ('01', '02', '03', '04', '05', '06', '07', '08', '09', '10', '99') 
                               THEN random_block 
                               ELSE '01' 
                           END as document_type,
                           created_at
                    FROM guri_records
                ''')
            else:
                cursor.execute('''
                    INSERT INTO guri_records_new 
                    (id, guri, sender, recipients, subject, datetime, avg_risk, document_type)
                    SELECT id, guri, sender, recipients, subject, datetime, avg_risk, 
                           CASE 
                               WHEN random_block IN ('01', '02', '03', '04', '05', '06', '07', '08', '09', '10', '99') 
                               THEN random_block 
                               ELSE '01' 
                           END as document_type
                    FROM guri_records
                ''')
        else:
            if has_created_at:
                cursor.execute('''
                    INSERT INTO guri_records_new 
                    (id, guri, sender, recipients, subject, datetime, avg_risk, document_type, created_at)
                    SELECT id, guri, sender, recipients, subject, datetime, avg_risk, 
                           COALESCE(document_type, '01'), created_at
                    FROM guri_records
                ''')
            else:
                cursor.execute('''
                    INSERT INTO guri_records_new 
                    (id, guri, sender, recipients, subject, datetime, avg_risk, document_type)
                    SELECT id, guri, sender, recipients, subject, datetime, avg_risk, 
                           COALESCE(document_type, '01')
                    FROM guri_records
                ''')
        
        rows_migrated = cursor.rowcount
        
        # Drop old table and rename new one
        cursor.execute('DROP TABLE guri_records')
        cursor.execute('ALTER TABLE guri_records_new RENAME TO guri_records')
        
        # Create indexes
        cursor.execute('CREATE INDEX IF NOT EXISTS idx_guri ON guri_records(guri)')
        cursor.execute('CREATE INDEX IF NOT EXISTS idx_sender ON guri_records(sender)')
        
        conn.commit()
        conn.close()
        
        print(f"[OK] Migration completed successfully!")
        print(f"     Migrated {rows_migrated} records")
        print(f"     Backup: {backup_path}")
        print()
        
        return True
        
    except Exception as e:
        print(f"[ERROR] Migration failed: {e}")
        print(f"\nBackup saved at: {backup_path}")
        print("You can restore by renaming the backup file.")
        return False


if __name__ == "__main__":
    import sys
    db_path = sys.argv[1] if len(sys.argv) > 1 else "C:/GeoFooter/guri_records.db"
    success = migrate_database(db_path)
    sys.exit(0 if success else 1)

