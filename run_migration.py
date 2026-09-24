#!/usr/bin/env python3
"""
Quick non-interactive migration from SQLite to MySQL using existing config.
"""

import os
import sys
import json
import logging

# Setup logging to console
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    stream=sys.stdout
)
logger = logging.getLogger(__name__)

# Import migration function
from migrate_sqlite_to_mysql import migrate_database

def main():
    print("=" * 60)
    print("GURI Database Migration Tool (Non-Interactive)")
    print("=" * 60)
    print()
    
    # Paths
    base_path = "C:/GeoFooter"
    sqlite_path = os.path.join(base_path, "guri_records.db")
    config_path = os.path.join(base_path, "guri_mysql_config.json")
    
    # Check SQLite exists
    if not os.path.exists(sqlite_path):
        print(f"ERROR: SQLite database not found: {sqlite_path}")
        return 1
    
    # Load MySQL config
    if not os.path.exists(config_path):
        print(f"ERROR: MySQL config not found: {config_path}")
        return 1
    
    try:
        with open(config_path, 'r') as f:
            mysql_config = json.load(f)
        print(f"Loaded MySQL config from: {config_path}")
    except Exception as e:
        print(f"ERROR: Failed to load MySQL config: {e}")
        return 1
    
    print(f"SQLite Database: {sqlite_path}")
    print(f"MySQL Database: {mysql_config.get('host')}:{mysql_config.get('port', 3306)}/{mysql_config.get('database')}")
    print()
    
    # Run migration
    result = migrate_database(sqlite_path, mysql_config, skip_duplicates=True, logger=logger)
    
    if result['success']:
        print()
        print("=" * 60)
        print("MIGRATION COMPLETED SUCCESSFULLY!")
        print("=" * 60)
        print(f"GURI Records: {result['guri_records']['migrated']} migrated, {result['guri_records']['skipped']} skipped")
        print(f"Supplier Codes: {result['supplier_codes']['migrated']} migrated, {result['supplier_codes']['skipped']} skipped")
        return 0
    else:
        print()
        print(f"MIGRATION FAILED: {result.get('error', 'Unknown error')}")
        return 1


if __name__ == "__main__":
    sys.exit(main())
