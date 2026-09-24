#!/usr/bin/env python3
"""
Test MySQL connection from the Python environment used by Outlook.
Run this with: C:\Python313\python.exe test_mysql_connection.py
"""

import sys
import os

print("=" * 60)
print("MySQL Connection Test")
print("=" * 60)
print(f"Python Version: {sys.version}")
print(f"Python Executable: {sys.executable}")
print()

# Test 1: Check if mysql.connector is available
print("Test 1: Checking mysql-connector-python...")
try:
    import mysql.connector
    print(f"  ✓ mysql-connector-python is installed (version: {mysql.connector.__version__})")
except ImportError as e:
    print(f"  ✗ FAILED: mysql-connector-python is NOT installed")
    print(f"    Error: {e}")
    print(f"    Fix: Run 'C:\\Python313\\python.exe -m pip install mysql-connector-python'")
    sys.exit(1)

# Test 2: Load config file
print()
print("Test 2: Loading MySQL config file...")
config_path = "C:/GeoFooter/guri_mysql_config.json"
try:
    import json
    with open(config_path, 'r') as f:
        config = json.load(f)
    print(f"  ✓ Config loaded from {config_path}")
    print(f"    Host: {config.get('host')}")
    print(f"    Port: {config.get('port', 3306)}")
    print(f"    Database: {config.get('database')}")
    print(f"    User: {config.get('user')}")
except Exception as e:
    print(f"  ✗ FAILED: Could not load config file")
    print(f"    Error: {e}")
    sys.exit(1)

# Test 3: Connect to MySQL
print()
print("Test 3: Connecting to MySQL database...")
try:
    conn = mysql.connector.connect(**config)
    print(f"  ✓ Connected to MySQL successfully!")
    
    cursor = conn.cursor()
    cursor.execute("SELECT COUNT(*) FROM guri_records")
    count = cursor.fetchone()[0]
    print(f"  ✓ Found {count} records in guri_records table")
    
    cursor.close()
    conn.close()
    print(f"  ✓ Connection closed")
except mysql.connector.Error as e:
    print(f"  ✗ FAILED: Could not connect to MySQL")
    print(f"    Error: {e}")
    sys.exit(1)

# Test 4: Test GURIDatabase class
print()
print("Test 4: Testing GURIDatabase class...")
try:
    sys.path.insert(0, "C:/GeoFooter")
    from guri import GURIDatabase
    
    guri_db = GURIDatabase(db_type="mysql", mysql_config=config)
    print(f"  ✓ GURIDatabase initialized (type: {guri_db.db_type})")
    
    # Test get_all_records
    records = guri_db.get_all_records()
    print(f"  ✓ get_all_records() returned {len(records)} records")
    
except Exception as e:
    print(f"  ✗ FAILED: GURIDatabase test failed")
    print(f"    Error: {e}")
    import traceback
    traceback.print_exc()
    sys.exit(1)

print()
print("=" * 60)
print("ALL TESTS PASSED!")
print("=" * 60)
print()
print("The Python 3.13 environment can connect to MySQL.")
print("Outlook should be able to save records to MySQL when processing emails.")
print()
print("If records still aren't appearing in the GUI:")
print("1. Make sure you've RESTARTED Outlook")
print("2. Send yourself a test email and check geolocate_debug.log")
print("3. The log should show 'GURI database type: MySQL'")
