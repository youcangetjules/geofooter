import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
print("Python is working!")
print(f"Python version: {sys.version}")
print(f"Current directory: {sys.path[0]}")

# Test imports
try:
    from guri.core import GURIDatabase
    print("SUCCESS: guri module imported")
    
    db = GURIDatabase()
    print(f"SUCCESS: Database at {db.db_path}")
    
    guri = db.generate_guri("test@test.com", "recip@test.com", "Test", "2025-11-02 12:00:00", "LOW", "01")
    print(f"SUCCESS: Generated GURI: {guri}")
    
except Exception as e:
    print(f"ERROR: {e}")
    import traceback
    traceback.print_exc()

