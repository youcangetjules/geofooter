#!/usr/bin/env python3
"""
Migration Script: Fix Component 6 to Match Document Type

This script updates all existing GURIs in the database so that Component 6
(the last 2 characters) matches the actual document_type stored in the record.

GURI Format: {5hex}x{5hex}x{8hex}x{8hex}x{3hex}x{DOC_TYPE}
                                                      ^^^^^^^
                                                   Component 6

Before: 3ae2bx7f4d1xa1b3c5e7x9f8e7d6cx2a4x5f  (random)
After:  3ae2bx7f4d1xa1b3c5e7x9f8e7d6cx2a4x01  (matches document_type)
"""

import sqlite3
from datetime import datetime
import sys

def fix_component6():
    """Update all GURIs to have Component 6 match the document_type."""
    
    # Database path
    db_path = "guri_records.db"
    
    print("=" * 80)
    print("GURI Component 6 Document Type Fix")
    print("=" * 80)
    print()
    print(f"Database: {db_path}")
    print(f"Started: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    print()
    
    try:
        # Connect to database
        conn = sqlite3.connect(db_path)
        conn.row_factory = sqlite3.Row
        cursor = conn.cursor()
        
        # Get all records
        print("Fetching all records...")
        cursor.execute("SELECT * FROM guri_records ORDER BY id")
        records = cursor.fetchall()
        
        print(f"Found {len(records)} records")
        print()
        
        if len(records) == 0:
            print("No records to update.")
            return
        
        # Backup notification
        print("IMPORTANT: Creating backup...")
        backup_path = f"guri_records.db.backup_{datetime.now().strftime('%Y%m%d_%H%M%S')}"
        import shutil
        shutil.copy2(db_path, backup_path)
        print(f"Backup created: {backup_path}")
        print()
        
        # Process each record
        updated_count = 0
        skipped_count = 0
        error_count = 0
        
        print("Processing records...")
        print("-" * 80)
        
        for record in records:
            record_id = record['id']
            old_guri = record['guri']
            document_type = record['document_type']
            
            # Split GURI into components
            components = old_guri.split('x')
            
            if len(components) != 6:
                print(f"WARNING: Record {record_id}: Invalid GURI format (expected 6 components, found {len(components)})")
                error_count += 1
                continue
            
            # Check if Component 6 already matches document_type
            current_component6 = components[5]
            
            if current_component6 == document_type:
                # Already correct
                skipped_count += 1
                if skipped_count <= 5:  # Show first 5 skipped
                    print(f"OK: Record {record_id}: Already correct (Component 6 = {document_type})")
                continue
            
            # Update Component 6 to match document_type
            components[5] = document_type
            new_guri = 'x'.join(components)
            
            # Update database
            try:
                cursor.execute(
                    "UPDATE guri_records SET guri = ? WHERE id = ?",
                    (new_guri, record_id)
                )
                updated_count += 1
                
                # Show progress for first 20 updates
                if updated_count <= 20:
                    print(f"UPDATED: Record {record_id}")
                    print(f"    Old: {old_guri}")
                    print(f"    New: {new_guri}")
                    print(f"    Component 6: {current_component6} -> {document_type}")
                    print()
                elif updated_count == 21:
                    print("... (showing first 20 updates, continuing in background) ...")
                    print()
                
            except Exception as e:
                print(f"ERROR: Record {record_id}: Error updating - {e}")
                error_count += 1
        
        # Commit changes
        conn.commit()
        
        print("-" * 80)
        print()
        print("SUMMARY")
        print("=" * 80)
        print(f"Total Records:     {len(records)}")
        print(f"Updated:           {updated_count}")
        print(f"Already Correct:   {skipped_count}")
        print(f"Errors:            {error_count}")
        print()
        
        if updated_count > 0:
            print("Database updated successfully!")
            print(f"Backup saved to: {backup_path}")
        else:
            print("No updates needed - all Component 6 values already match document types.")
        
        print()
        print(f"Completed: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
        print("=" * 80)
        
        # Close connection
        cursor.close()
        conn.close()
        
        return updated_count, skipped_count, error_count
        
    except sqlite3.Error as e:
        print(f"Database error: {e}")
        sys.exit(1)
    except Exception as e:
        print(f"Unexpected error: {e}")
        sys.exit(1)


def verify_fixes():
    """Verify that all Component 6 values match document_type."""
    
    print()
    print("=" * 80)
    print("VERIFICATION")
    print("=" * 80)
    print()
    
    db_path = "guri_records.db"
    conn = sqlite3.connect(db_path)
    conn.row_factory = sqlite3.Row
    cursor = conn.cursor()
    
    cursor.execute("SELECT * FROM guri_records ORDER BY id")
    records = cursor.fetchall()
    
    mismatches = []
    
    for record in records:
        guri = record['guri']
        document_type = record['document_type']
        components = guri.split('x')
        
        if len(components) == 6:
            component6 = components[5]
            if component6 != document_type:
                mismatches.append({
                    'id': record['id'],
                    'guri': guri,
                    'component6': component6,
                    'document_type': document_type
                })
    
    if mismatches:
        print(f"WARNING: Found {len(mismatches)} mismatches:")
        for m in mismatches[:10]:  # Show first 10
            print(f"  Record {m['id']}: Component 6 = {m['component6']}, Document Type = {m['document_type']}")
    else:
        print("All Component 6 values match their document types!")
        print(f"Verified {len(records)} records")
    
    cursor.close()
    conn.close()
    
    print()
    print("=" * 80)


if __name__ == "__main__":
    print()
    print("This script will update all GURIs so Component 6 matches the document_type.")
    print()
    
    response = input("Do you want to proceed? (yes/no): ").strip().lower()
    
    if response in ['yes', 'y']:
        print()
        updated, skipped, errors = fix_component6()
        
        if updated > 0 or errors > 0:
            verify_fixes()
        
        print()
        print("Migration complete!")
    else:
        print()
        print("Migration cancelled.")
        print()

