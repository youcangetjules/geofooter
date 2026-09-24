# GURI Database System

**Globally Unique Record Identifier** - Generator, Database Manager, and GUI Viewer

## Overview

The GURI system provides a comprehensive solution for generating, storing, and managing unique identifiers for emails and documents. It includes:

- **Core Library** (`guri.py`) - Database operations and GURI generation
- **GUI Application** (`guri_gui.py`) - Visual interface for managing GURIs
- **Dual Database Support** - Works with both SQLite and MySQL

## Features

✅ Generate unique hexadecimal identifiers  
✅ Support for multiple document types (Email, Word, Excel, PDF, etc.)  
✅ SQLite and MySQL database backends  
✅ Comprehensive GUI with:
   - Record viewing and browsing
   - GURI creation with validation
   - Search functionality
   - Export to JSON  
✅ Auto-fill for non-email documents  
✅ Indexed database for fast queries  

## Installation

### 1. Basic Installation (SQLite only)

No additional packages required! SQLite is built into Python.

```bash
# Just run the GUI
python guri_gui.py
```

### 2. With MySQL Support

```bash
pip install -r requirements_guri.txt
```

Or install manually:

```bash
pip install mysql-connector-python
```

## Quick Start

### Launch the GUI

```bash
python guri_gui.py
```

The GUI will automatically connect to a local SQLite database.

### Document Type Codes

| Code | Document Type |
|------|---------------|
| 01   | Email |
| 02   | Word Document |
| 03   | Excel Spreadsheet |
| 04   | PowerPoint Presentation |
| 05   | PDF Document |
| 06   | Text File |
| 07   | Image File |
| 08   | Video File |
| 09   | Audio File |
| 10   | Archive/Zip |
| 99   | Other |

### Document Type Rules

**For Emails (Type 01):**
- All fields are required
- Sender: actual email address
- Recipients: actual recipient addresses
- Risk Level: actual risk assessment (e.g., "LOW (10/100)")

**For Non-Email Documents (Types 02-99):**
- Sender: automatically set to `julian.garrett@aliniant.com`
- Recipients: automatically set to `FFFFF`
- Document Location: **REQUIRED** - full file path (e.g., "C:\Documents\Report.docx")
- Subject: document name/description

## Using the GUI

### 1. View Records Tab

- Browse all GURI records
- Navigate with Previous/Next buttons
- Double-click any record to view details
- Refresh to see latest data

### 2. Create GURI Tab

**For Emails:**
1. Select "01 - Email" document type
2. Fill in all fields (Sender, Recipients, Subject, DateTime, Risk)
3. Click "Generate GURI"

**For Other Documents:**
1. Select appropriate document type (02-10, 99)
2. Click "Auto-fill Non-Email" to populate sender and recipients
3. Enter document name in Subject field
4. **Enter file path in Document Location field** (e.g., "C:\Documents\Report.docx")
5. Click "Generate GURI"

### 3. Search Tab

- Search for records by sender email
- Supports partial matching
- Double-click results to view details

### 4. Decompositor Tab

- Break down any GURI into its 6 components
- Validates format and length
- Shows hex validation for each component
- Paste from clipboard support
- Press Enter to decompose quickly

## Connecting to MySQL

### GUI Method

1. Click "File" → "Connect to MySQL..."
2. Enter connection details:
   - Host: your MySQL server address
   - Port: 3306 (default)
   - User: your MySQL username
   - Password: your MySQL password
   - Database: guri_db (or your database name)
3. Click "Connect"

### Configuration File Method

1. Copy `guri_mysql_config.example.json` to `guri_mysql_config.json`
2. Edit with your MySQL credentials:

```json
{
    "host": "localhost",
    "user": "root",
    "password": "your_password",
    "database": "guri_db",
    "port": 3306,
    "charset": "utf8mb4"
}
```

## Using the Python Library

### Basic Example

```python
from guri import GURIDatabase

# Connect to SQLite (default)
db = GURIDatabase()

# Generate a GURI for an email
guri = db.generate_guri(
    sender="user@example.com",
    recipients="recipient@example.com",
    subject="Important Email",
    dt_str="2025-11-01 12:00:00",
    avg_risk="LOW (10/100)",
    document_type="01"  # Email
)

print(f"Generated GURI: {guri}")
```

### MySQL Example

```python
from guri import GURIDatabase

# MySQL configuration
mysql_config = {
    "host": "localhost",
    "user": "root",
    "password": "your_password",
    "database": "guri_db",
    "port": 3306
}

# Connect to MySQL
db = GURIDatabase(db_type="mysql", mysql_config=mysql_config)

# Generate GURI for a Word document
guri = db.generate_guri(
    sender="julian.garrett@aliniant.com",
    recipients="FFFFF",
    subject="Project Proposal.docx",
    dt_str="2025-11-01 14:30:00",
    avg_risk="C:\\Documents\\Project Proposal.docx",  # Document location
    document_type="02"  # Word Document
)
```

### Search Records

```python
# Search by sender
results = db.search_by_sender("julian.garrett")

for record in results:
    print(f"GURI: {record['guri']}")
    print(f"Subject: {record['subject']}")
    print(f"Doc Type: {record['document_type']}")
    print("-" * 40)
```

### Get All Records

```python
# Get first 100 records
records = db.get_all_records(limit=100, offset=0)

# Get record count
total = db.get_record_count()
print(f"Total GURIs in database: {total}")
```

## GURI Format

Format: `{5hex}x{5hex}x{8hex}x{8hex}x{3hex}x{2hex}`

Example: `a1b2cx3d4e5xf6g7h8i9xj0k1l2m3xn4oxp5`

Components:
- 5 random hexadecimal characters
- separator 'x'
- 5 more hex characters
- separator 'x'
- 8 hex characters
- separator 'x'
- 8 hex characters
- separator 'x'
- 3 hex characters
- separator 'x'
- 2 hex characters

**Total length: 37 characters**

## Database Schema

```sql
CREATE TABLE guri_records (
    id              INT AUTO_INCREMENT PRIMARY KEY,
    guri            VARCHAR(50) UNIQUE NOT NULL,
    sender          VARCHAR(255) NOT NULL,
    recipients      TEXT,
    subject         TEXT,
    datetime        VARCHAR(50),
    avg_risk        VARCHAR(50),
    document_type   VARCHAR(10) DEFAULT '01',
    created_at      TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
    INDEX idx_guri (guri),
    INDEX idx_sender (sender),
    INDEX idx_created_at (created_at)
);
```

## Command Line Help

```bash
# Display help
python guri.py --help

# Run demonstration
python guri.py
```

## Exporting Data

### From GUI

1. Go to "File" → "Export Records..."
2. Choose save location
3. Records saved as JSON

### Programmatically

```python
import json

records = db.get_all_records(limit=10000)

with open('guri_export.json', 'w') as f:
    json.dump(records, f, indent=2, default=str)
```

## Troubleshooting

### MySQL Connection Issues

1. Ensure MySQL server is running
2. Check credentials are correct
3. Verify database exists:

```sql
CREATE DATABASE IF NOT EXISTS guri_db 
CHARACTER SET utf8mb4 
COLLATE utf8mb4_unicode_ci;
```

4. Grant permissions:

```sql
GRANT ALL PRIVILEGES ON guri_db.* TO 'your_user'@'localhost';
FLUSH PRIVILEGES;
```

### SQLite File Location

Default: `C:/GeoFooter/guri_records.db`

To use a different location:

```python
db = GURIDatabase(db_type="sqlite", db_path="path/to/your/database.db")
```

### Database Migration

To migrate from SQLite to MySQL or vice versa:

```python
# Export from SQLite
sqlite_db = GURIDatabase(db_type="sqlite")
records = sqlite_db.get_all_records(limit=100000)

# Import to MySQL
mysql_db = GURIDatabase(db_type="mysql", mysql_config=mysql_config)

for record in records:
    mysql_db.insert_guri(
        guri=record['guri'],
        sender=record['sender'],
        recipients=record['recipients'],
        subject=record['subject'],
        dt_str=record['datetime'],
        avg_risk=record['avg_risk'],
        document_type=record['document_type']
    )
```

## File Structure

```
GeoFooter/
├── guri.py                          # Core GURI library
├── guri_gui.py                      # GUI application
├── guri_mysql_config.example.json  # MySQL config template
├── requirements_guri.txt            # Python dependencies
├── GURI_README.md                   # This file
└── guri_records.db                  # SQLite database (auto-created)
```

## Version

**Version:** 1.0.0  
**Author:** Extracted from geolocate_headers.py  
**License:** Use as needed  
**Copyright:** © 2025 Aliniant Labs

## Support

For issues or questions, refer to the main documentation or contact support.

---

**Happy GURI Managing! 🎯**

