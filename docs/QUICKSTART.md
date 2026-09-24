# GURI Database - Quick Start Guide

## 🚀 Getting Started in 60 Seconds

### Step 1: Install MySQL Support (Optional)

If you want to use MySQL instead of SQLite:

```bash
pip install mysql-connector-python
```

### Step 2: Launch the GUI

**Option A: Double-click the launcher**
```
Launch_GURI_GUI.bat
```

**Option B: Command line**
```bash
python guri_gui.py
```

### Step 3: Start Using!

The GUI will automatically connect to SQLite. You're ready to go!

---

## 📋 Common Tasks

### Create a GURI for an Email

1. Go to **"Create GURI"** tab
2. Make sure **"01 - Email"** is selected
3. Fill in:
   - **Sender**: actual sender email
   - **Recipients**: recipient email(s)
   - **Subject**: email subject
   - **Date/Time**: click "Now" or enter manually
   - **Risk**: e.g., "LOW (10/100)"
4. Click **"Generate GURI"**

### Create a GURI for a Document (Word, Excel, etc.)

1. Go to **"Create GURI"** tab
2. Select document type:
   - **02** for Word
   - **03** for Excel
   - **05** for PDF
   - etc.
3. Click **"Auto-fill Non-Email"** button (auto-fills sender & recipients)
4. Enter document name in **Subject** field
5. Enter file path in **Document Location** field (e.g., "C:\Documents\Report.docx")
6. Click **"Generate GURI"**

### Search for Records

1. Go to **"Search"** tab
2. Enter sender email (partial match works)
3. Click **"Search"**
4. Double-click any result to view details

### Decompose a GURI

1. Go to **"Decompositor"** tab
2. Paste or type a GURI
3. Press **Enter** or click **"Decompose"**
4. View the breakdown of all 6 components

### Export Data

1. Go to **File → Export Records...**
2. Choose save location
3. Records exported as JSON

---

## 🔌 Connect to MySQL

### GUI Method

1. Click **File → Connect to MySQL...**
2. Enter:
   - **Host**: `localhost` (or your MySQL server)
   - **Port**: `3306`
   - **User**: `root` (or your MySQL user)
   - **Password**: your MySQL password
   - **Database**: `guri_db`
3. Click **Connect**

### First Time MySQL Setup

Before connecting, create the database:

```sql
CREATE DATABASE guri_db CHARACTER SET utf8mb4 COLLATE utf8mb4_unicode_ci;
```

---

## 📁 Document Type Quick Reference

| Code | Type |
|------|------|
| **01** | Email (all fields required) |
| **02** | Word Document |
| **03** | Excel Spreadsheet |
| **04** | PowerPoint |
| **05** | PDF |
| **06** | Text File |
| **07** | Image |
| **08** | Video |
| **09** | Audio |
| **10** | Archive/Zip |
| **99** | Other |

---

## 🔍 GURI Format

Example: `a1b2cx3d4e5xf6g7h8i9xj0k1l2m3xn4oxp5`

- **37 characters total**
- **Hexadecimal** (0-9, a-f)
- **Format**: `{5}x{5}x{8}x{8}x{3}x{2}`

---

## 💡 Tips

### For Emails
✅ All fields are **required**  
✅ Use actual email addresses  
✅ Include proper risk assessment  

### For Documents
✅ Sender is **always** `julian.garrett@aliniant.com`  
✅ Recipients set to **FFFFF**  
✅ **Document Location** must be entered (e.g., "C:\Documents\Report.docx")  
✅ Subject should be the **document name**  

### General
✅ **Double-click** any record to view full details  
✅ Use **Previous/Next** buttons to browse records  
✅ Click **Refresh** after creating new GURIs  
✅ **Search** supports partial email matches  

---

## ❓ Troubleshooting

### GUI won't start?
- Make sure Python 3.7+ is installed
- Run: `python --version` to check

### Can't connect to MySQL?
- Make sure MySQL server is running
- Check username/password are correct
- Verify database exists
- Try SQLite first to test the GUI

### Records not showing?
- Click the **Refresh** button
- Check you're connected (top of window should be green)

---

## 📚 More Information

For complete documentation, see **GURI_README.md**

For command-line help:
```bash
python guri.py --help
```

---

## 🎯 Next Steps

1. ✅ Launch the GUI
2. ✅ Create your first GURI
3. ✅ Explore the search feature
4. ✅ Try exporting data
5. ✅ Connect to MySQL (optional)

---

**Happy GURI Creating! 🎉**

*Version 1.0.0 | © 2025 Aliniant Labs*

