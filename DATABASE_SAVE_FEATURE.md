# GURI Database Auto-Save Feature

## Overview

When you click **"Generate GURI"** in the GUI, the GURI is **automatically saved to the database**. This document explains how this works and what happens when you create a new GURI.

---

## How It Works

### 1. Generate GURI Button Click

When you click the "Generate GURI" button:

```
1. ✅ Validates all required fields
2. ✅ Calls db.generate_guri() which:
   - Checks if a GURI with same parameters already exists
   - If exists: returns the existing GURI
   - If new: generates random hexadecimal GURI
   - Inserts into database via db.insert_guri()
3. ✅ Resets view to page 1 (most recent records)
4. ✅ Refreshes the record list
5. ✅ Highlights the newest record (at top)
6. ✅ Shows success dialog
7. ✅ Offers to switch to View Records tab
```

### 2. Database Insert

The `generate_guri()` method automatically calls:

```python
self.insert_guri(guri, sender, recipients, subject, dt_str, avg_risk, document_type)
```

This inserts a record with:
- **guri** - The generated unique identifier
- **sender** - Email or julian.garrett@aliniant.com
- **recipients** - Recipient emails or FFFFF
- **subject** - Email subject or document name
- **datetime** - Timestamp
- **avg_risk** - Risk level or document location
- **document_type** - Type code (01-10, 99)
- **created_at** - Auto-generated timestamp

### 3. Duplicate Prevention

If you try to generate a GURI with the **exact same parameters** (sender, recipients, subject, datetime, avg_risk), it will:
- ✅ Return the **existing GURI** instead of creating a duplicate
- ✅ This prevents duplicate records in the database

---

## Visual Feedback

### Success Message

After clicking "Generate GURI", you'll see:

```
┌─────────────────────────────────────────────┐
│ ✓ GURI Generated and Saved to Database!    │
│                                             │
│ GURI: a1b2cx3d4e5xf6g7h8i9xj0k1l2m3xn4oxp5 │
│ ==================================================│
│                                             │
│ Document Type: Email                        │
│ Sender: user@example.com                    │
│ Recipients: recipient@example.com           │
│ Subject: Test Email                         │
│ DateTime: 2025-11-01 14:30:00              │
│ Risk Level: LOW (10/100)                    │
│                                             │
│ ==================================================│
│ Status: ✓ Saved to database                │
│ View in 'View Records' tab (First Page)    │
└─────────────────────────────────────────────┘
```

### Confirmation Dialog

A popup will appear:

```
┌──────────────────────────────────────────┐
│ Success                                  │
├──────────────────────────────────────────┤
│ GURI Generated and Saved to Database!   │
│                                          │
│ GURI: a1b2cx3d4e5xf6g7h8i9xj0k1l2m3xn... │
│                                          │
│ The record has been added to the        │
│ database. You can view it in the        │
│ 'View Records' tab.                      │
│                                          │
│           [ OK ]                         │
└──────────────────────────────────────────┘
```

### View Record Option

Then you'll be asked:

```
┌──────────────────────────────────────────┐
│ View Record?                             │
├──────────────────────────────────────────┤
│ Would you like to switch to the View    │
│ Records tab to see your new GURI?        │
│                                          │
│         [ Yes ]    [ No ]                │
└──────────────────────────────────────────┘
```

If you click **Yes**:
- Switches to "View Records" tab
- Shows first page (newest records)
- Your new GURI is **highlighted at the top**

---

## Viewing Saved Records

### Automatic Refresh

After creating a GURI:
1. Current page resets to **Page 1** (first page)
2. Record list refreshes automatically
3. **Newest record appears at the top** (highlighted)
4. Total record count updates

### Manual Navigation

Use the toolbar buttons:

| Button | Action |
|--------|--------|
| **🔄 Refresh** | Reload current page |
| **⬅ Previous** | Go to previous page |
| **Next ➡** | Go to next page |
| **📄 First Page** | Jump to page 1 (newest records) |

### Record Sorting

Records are **always sorted by creation date** (newest first):
- Page 1 = Most recent records
- Last page = Oldest records
- Your new GURI always appears on Page 1

---

## Status Bar Updates

The status bar at the bottom shows:

**After generating GURI:**
```
GURI generated and saved to database: a1b2cx3d4e5xf6g7h8i9xj0k1l2m3xn4oxp5
```

**When viewing Page 1:**
```
Loaded 50 records - Most recent at top
```

**When viewing other pages:**
```
Loaded 50 records
```

---

## Database Verification

### Check Database Directly

**SQLite:**
```bash
sqlite3 C:/GeoFooter/guri_records.db
SELECT * FROM guri_records ORDER BY created_at DESC LIMIT 5;
.quit
```

**MySQL:**
```sql
SELECT * FROM guri_records ORDER BY created_at DESC LIMIT 5;
```

### Check in GUI

1. Go to **"View Records"** tab
2. Click **"📄 First Page"** button
3. Your most recent GURI should be **highlighted at the top**
4. Double-click to view full details

### Export and Verify

1. **File → Export Records...**
2. Save as JSON
3. Open the JSON file
4. Search for your GURI

---

## Duplicate Handling

### Same Parameters = Same GURI

If you enter:
- Same sender
- Same recipients
- Same subject
- Same datetime
- Same risk/location
- Same document type

The system will:
- ✅ Return the **existing GURI** (not create a new one)
- ✅ Show the existing GURI in the result
- ✅ **Not create a duplicate** database entry

This is intentional to maintain data integrity!

### Different Parameters = New GURI

Even if one parameter changes, you get a **new unique GURI**:

```
First GURI:
  Subject: "Report.docx"
  → GURI: a1b2cx3d4e5xf6g7h8i9xj0k1l2m3xn4oxp5

Second GURI:
  Subject: "Report_v2.docx"  ← Changed!
  → GURI: z9y8wx7v6ux5t4s3r2q1xp0o9n8m7xl6xk5  ← Different!
```

---

## Record Count

The **Total Records** label updates automatically:

```
Before: Total Records: 1556
After:  Total Records: 1557  ← Incremented!
```

This confirms your GURI was saved.

---

## Troubleshooting

### "GURI not appearing in list"

**Solution:**
1. Click **"📄 First Page"** button
2. Click **"🔄 Refresh"** button
3. Check page number - should be **Page: 1**

### "Same GURI generated twice"

**Explanation:**
- This is correct behavior!
- Same parameters = same GURI
- Change at least one parameter to get a new GURI

### "Can't find my GURI"

**Solution:**
1. Use **Search** tab
2. Enter your sender email
3. Or go to **First Page** to see newest records

### "Record count didn't increase"

**Possible Causes:**
1. Generated duplicate (same parameters)
2. Database connection issue
3. Database error (check logs)

**Solution:**
- Check the result display - does it show a new GURI?
- Look at the success message
- Try searching for the GURI
- Check database logs: `C:/GeoFooter/geolocate_debug.log`

---

## Technical Details

### Database Columns

When saved, each GURI record contains:

```sql
CREATE TABLE guri_records (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    guri TEXT UNIQUE NOT NULL,           -- The generated GURI
    sender TEXT NOT NULL,                -- Email or julian.garrett@aliniant.com
    recipients TEXT,                     -- Recipients or FFFFF
    subject TEXT,                        -- Subject or document name
    datetime TEXT,                       -- Timestamp
    avg_risk TEXT,                       -- Risk level or document location
    document_type TEXT DEFAULT '01',     -- Type code
    created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP  -- Auto-generated
)
```

### Indexes

For fast retrieval:
- `idx_guri` - Index on guri column
- `idx_sender` - Index on sender column

### Transaction Safety

The database insert uses:
- **SQLite:** `INSERT OR IGNORE` (prevents duplicates)
- **MySQL:** `INSERT IGNORE` (prevents duplicates)

This ensures data integrity even if there's an error.

---

## Best Practices

### After Creating a GURI

1. ✅ **Note the GURI** - Copy it for your records
2. ✅ **Verify in database** - Click "Yes" to view it
3. ✅ **Export regularly** - Backup your database
4. ✅ **Check record count** - Ensure it incremented

### Preventing Duplicates

1. ✅ Use unique subjects for documents
2. ✅ Include timestamps in filenames
3. ✅ Check if GURI exists before creating
4. ✅ Use search to find existing GURIs

### Database Maintenance

1. ✅ Regular backups (automatic on migration)
2. ✅ Export to JSON periodically
3. ✅ Monitor record count
4. ✅ Clean up test records if needed

---

## Summary

✅ **Click "Generate GURI"** → GURI is **automatically saved** to database  
✅ **No manual save needed** - it's instant and automatic  
✅ **View immediately** - Record appears at top of list  
✅ **Highlighted by default** - Easy to spot your new GURI  
✅ **Duplicate prevention** - Same parameters = same GURI  
✅ **Instant feedback** - Success messages and status updates  

**You don't need to do anything extra - just click "Generate GURI" and it's saved!**

---

**Last Updated:** November 1, 2025  
**Version:** 1.1.0  
**Feature Status:** ✅ Active

