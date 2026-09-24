#!/usr/bin/env python3
"""Test script to verify GURI module works correctly."""

import sys
import os

print("=" * 80)
print("Testing GURI Module Import and Functionality")
print("=" * 80)
print()

# Test 1: Import guri module
print("Test 1: Importing guri module...")
try:
    from guri import GURIDatabase
    print("✓ SUCCESS: guri module imported")
except Exception as e:
    print(f"✗ FAILED: Could not import guri module: {e}")
    sys.exit(1)

# Test 2: Initialize database
print("\nTest 2: Initializing GURIDatabase...")
try:
    db = GURIDatabase()
    print(f"✓ SUCCESS: Database initialized at {db.db_path}")
except Exception as e:
    print(f"✗ FAILED: Could not initialize database: {e}")
    sys.exit(1)

# Test 3: Generate a test GURI
print("\nTest 3: Generating test GURI...")
try:
    guri = db.generate_guri(
        sender="test@example.com",
        recipients="recipient@example.com",
        subject="Test Email",
        dt_str="2025-11-02 12:00:00",
        avg_risk="LOW (5/100)",
        document_type="01"
    )
    print(f"✓ SUCCESS: Generated GURI: {guri}")
except Exception as e:
    print(f"✗ FAILED: Could not generate GURI: {e}")
    sys.exit(1)

# Test 4: Get record count
print("\nTest 4: Getting record count...")
try:
    count = db.get_record_count()
    print(f"✓ SUCCESS: Database has {count} records")
except Exception as e:
    print(f"✗ FAILED: Could not get record count: {e}")
    sys.exit(1)

# Test 5: Test geolocate_headers import
print("\nTest 5: Testing geolocate_headers imports...")
try:
    import geolocate_headers
    print("✓ SUCCESS: geolocate_headers module loaded")
except Exception as e:
    print(f"✗ FAILED: Could not import geolocate_headers: {e}")
    sys.exit(1)

print("\n" + "=" * 80)
print("ALL TESTS PASSED!")
print("=" * 80)
print("\nThe geo footer system is ready to use.")
print("Run: python geolocate_headers.py <header_file> [output_file]")
print()

