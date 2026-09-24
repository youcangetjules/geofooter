#!/usr/bin/env python3
"""
Fix Duplicate GURI Issues

This script handles records that couldn't be updated due to UNIQUE constraint violations.
It generates completely new GURIs for these records.
"""

import sqlite3
import random
from datetime import datetime

def randhex(n):
    """Generate random hexadecimal string of length n."""
    return ''.join(random.choices('0123456789abcdef', k=n))

def fix_duplicates():
    """Fix records with duplicate GURI issues."""
    
    db_path = os.path.join(r"C:\GeoFooter", "datastore", "guri_records.db")
    
    print("=" * 80)
    print("FIXING DUPLICATE GURI RECORDS")
    print("=" * 80)
    print()
    
    conn = sqlite3.connect(db_path)
    conn.row_factory = sqlite3.Row
    cursor = conn.cursor()
    
    # Find records where Component 6 doesn't match document_type
    cursor.execute("SELECT * FROM guri_records")
    all_records = cursor.fetchall()
    
    mismatches = []
    for record in all_records:
        guri = record['guri']
        document_type = record['document_type']
        components = guri.split('x')
        
        if len(components) == 6:
            component6 = components[5]
            if component6 != document_type:
                mismatches.append(record)
    
    print(f"Found {len(mismatches)} records with mismatched Component 6")
    print()
    
    if len(mismatches) == 0:
        print("No mismatches found!")
        cursor.close()
        conn.close()
        return
    
    # Backup
    print("Creating backup...")
    import shutil
    backup_path = f"guri_records.db.backup_duplicates_{datetime.now().strftime('%Y%m%d_%H%M%S')}"
    shutil.copy2(db_path, backup_path)
    print(f"Backup created: {backup_path}")
    print()
    
    # Fix each mismatch by generating a completely new GURI
    print("Generating new GURIs for mismatched records...")
    print("-" * 80)
    
    fixed_count = 0
    
    for record in mismatches:
        record_id = record['id']
        old_guri = record['guri']
        document_type = record['document_type']
        
        # Generate new GURI with correct document_type in Component 6
        max_attempts = 100
        for attempt in range(max_attempts):
            new_guri = f"{randhex(5)}x{randhex(5)}x{randhex(8)}x{randhex(8)}x{randhex(3)}x{document_type}"
            
            # Check if this GURI already exists
            cursor.execute("SELECT guri FROM guri_records WHERE guri = ?", (new_guri,))
            if cursor.fetchone() is None:
                # GURI is unique, use it
                try:
                    cursor.execute(
                        "UPDATE guri_records SET guri = ? WHERE id = ?",
                        (new_guri, record_id)
                    )
                    fixed_count += 1
                    print(f"FIXED: Record {record_id}")
                    print(f"  Old GURI: {old_guri}")
                    print(f"  New GURI: {new_guri}")
                    print(f"  Doc Type: {document_type}")
                    print()
                    break
                except Exception as e:
                    print(f"ERROR: Record {record_id}: {e}")
                    break
        else:
            print(f"WARNING: Could not generate unique GURI for record {record_id} after {max_attempts} attempts")
    
    # Commit changes
    conn.commit()
    
    print("-" * 80)
    print()
    print("SUMMARY")
    print("=" * 80)
    print(f"Mismatched Records: {len(mismatches)}")
    print(f"Fixed:              {fixed_count}")
    print()
    
    if fixed_count > 0:
        print("Database updated successfully!")
        print(f"Backup saved to: {backup_path}")
    
    print()
    
    cursor.close()
    conn.close()
    
    return fixed_count

if __name__ == "__main__":
    print()
    response = input("Fix duplicate GURI records? (yes/no): ").strip().lower()
    
    if response in ['yes', 'y']:
        print()
        fixed = fix_duplicates()
        
        print()
        print("Done!")
    else:
        print()
        print("Cancelled.")

