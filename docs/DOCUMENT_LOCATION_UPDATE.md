# Document Location Field Update

## Summary of Changes

The GURI GUI has been updated to properly handle the difference between emails and documents. 

### Key Changes

**Document Type 01 (Email):**
- Field labeled: **"Risk Level"**
- User enters risk assessment (e.g., "LOW (10/100)")
- All fields required

**Document Types 02-99 (Non-Email Documents):**
- Field labeled: **"Document Location"**  
- User **MUST enter** the file path/location (e.g., "C:\Documents\Report.docx")
- No longer auto-fills to "FFFFF"
- Sender auto-fills to: `julian.garrett@aliniant.com`
- Recipients auto-fill to: `FFFFF`

---

## What Changed in the GUI

### 1. Dynamic Field Label

The label for the risk/location field now changes based on document type:

```
Email (Type 01):        Risk Level: [           ]
Word Doc (Type 02):     Document Location: [           ]
Excel (Type 03):        Document Location: [           ]
PDF (Type 05):          Document Location: [           ]
etc.
```

### 2. Auto-Fill Behavior

**BEFORE:**
- Clicking "Auto-fill Non-Email" set:
  - Sender → `julian.garrett@aliniant.com`
  - Recipients → `FFFFF`
  - Risk → `FFFFF` ❌

**AFTER:**
- Clicking "Auto-fill Non-Email" now sets:
  - Sender → `julian.garrett@aliniant.com`
  - Recipients → `FFFFF`
  - Document Location → *(cleared, user must enter)* ✅

### 3. Validation Messages

Validation now shows the correct field name:

**For Emails:**
```
The following fields are required:
• Risk Level
```

**For Documents:**
```
The following fields are required:
• Document Location
```

### 4. Display in Tables

The column header in the record view has been updated:

**BEFORE:** `Risk Level`  
**AFTER:** `Risk/Location`

This makes it clear the field serves dual purposes.

### 5. Record Details

When viewing record details (double-click), the field name changes:

**For Email:**
```
Risk Level: LOW (10/100)
```

**For Document:**
```
Document Location: C:\Documents\Report.docx
```

---

## How to Use

### Creating Email GURI (Type 01)

1. Select **"01 - Email"**
2. Fill in:
   - Sender: actual email
   - Recipients: actual recipients
   - Subject: email subject
   - DateTime: email timestamp
   - **Risk Level**: e.g., "LOW (10/100)"
3. Click "Generate GURI"

### Creating Document GURI (Types 02-99)

1. Select document type (e.g., **"02 - Word Document"**)
2. Click **"Auto-fill Non-Email"** button
   - ✅ Sender fills automatically
   - ✅ Recipients fills automatically
   - ⚠️ Document Location is CLEARED (you must fill it)
3. Enter:
   - Subject: document name (e.g., "Q4 Report")
   - **Document Location**: full path (e.g., "C:\Documents\Q4_Report.docx")
4. DateTime is auto-filled to current time (or change as needed)
5. Click "Generate GURI"

---

## Examples

### Email Example

```
Document Type: 01 - Email
Sender: john.doe@company.com
Recipients: jane.smith@company.com
Subject: Meeting Notes
DateTime: 2025-11-01 14:30:00
Risk Level: LOW (15/100)
```

### Word Document Example

```
Document Type: 02 - Word Document
Sender: julian.garrett@aliniant.com
Recipients: FFFFF
Subject: Project Proposal
DateTime: 2025-11-01 14:30:00
Document Location: C:\Projects\Proposals\Project_Proposal_2025.docx
```

### Excel Spreadsheet Example

```
Document Type: 03 - Excel Spreadsheet
Sender: julian.garrett@aliniant.com
Recipients: FFFFF
Subject: Q4 Financial Report
DateTime: 2025-11-01 14:30:00
Document Location: \\NetworkDrive\Finance\Q4_2025_Report.xlsx
```

### PDF Document Example

```
Document Type: 05 - PDF Document
Sender: julian.garrett@aliniant.com
Recipients: FFFFF
Subject: Contract Agreement
DateTime: 2025-11-01 14:30:00
Document Location: C:\Contracts\Client_A\Agreement_2025.pdf
```

---

## Database Schema

The database column name remains `avg_risk`, but the meaning changes:

| Document Type | Column Name | Actual Meaning |
|---------------|-------------|----------------|
| 01 (Email) | `avg_risk` | Risk level assessment |
| 02-99 (Documents) | `avg_risk` | Document file location/path |

This allows backward compatibility while supporting the new functionality.

---

## Migration Note

**Existing Records:** All previously migrated records were set to document type "01" (Email) by default. If any of those were actually documents, you can:

1. View the record in the GUI
2. Note the GURI and details
3. Create a new GURI with the correct document type and location
4. Update your references to use the new GURI

---

## Benefits

✅ **Clear Separation:** Email risk vs document location clearly distinguished  
✅ **User-Friendly:** Field label changes automatically based on selection  
✅ **Validation:** Proper error messages showing the correct field name  
✅ **Flexible:** Supports network paths, local paths, relative paths  
✅ **Trackable:** Know exactly where each document is stored  

---

## Document Location Best Practices

### Recommended Formats

✅ **Full Path:** `C:\Documents\Reports\Q4_2025.docx`  
✅ **Network Path:** `\\server\share\folder\file.xlsx`  
✅ **UNC Path:** `\\FILESERVER\Projects\Design\mockup.pdf`  
✅ **Relative Path:** `.\Documents\notes.txt` (if consistent)  
✅ **SharePoint:** `https://company.sharepoint.com/sites/docs/file.docx`  
✅ **OneDrive:** `C:\Users\julian\OneDrive\Documents\file.xlsx`  

### Avoid

❌ Just filename: `report.docx`  
❌ Vague: "On my desktop"  
❌ Incomplete: "C:\Documents"  

### Tips

- **Be specific:** Include full path whenever possible
- **Use consistent format:** Decide on a standard (UNC, local, etc.)
- **Include extension:** Always include the file extension
- **Test access:** Ensure the path is accessible from where you need it
- **Version control:** Consider including version in filename if path changes

---

## Troubleshooting

### "Document Location is required" error

**Cause:** You selected a non-email document type but didn't enter the location.

**Solution:** Enter the full file path in the Document Location field.

### Auto-fill cleared my location

**Cause:** This is intentional - auto-fill only fills sender and recipients for documents.

**Solution:** Enter the document location manually after clicking auto-fill.

### Wrong field label showing

**Cause:** The document type selection hasn't triggered the label update.

**Solution:** 
1. Click a different document type
2. Click back to your desired type
3. Or restart the GUI

---

## Technical Details

### Files Modified

- `guri_gui.py` - Main GUI logic updated
- `QUICKSTART.md` - Usage instructions updated
- `GURI_README.md` - Full documentation updated
- `DOCUMENT_LOCATION_UPDATE.md` - This file (new)

### Code Changes

1. Added `self.risk_label` as a configurable label widget
2. Updated `_on_doc_type_change()` to set label text dynamically
3. Modified `_autofill_non_email()` to NOT fill risk/location field
4. Enhanced validation to show correct field name
5. Updated result display to show correct field name
6. Modified record details to show correct field name
7. Changed table column header to "Risk/Location"

---

## Version History

**Version 1.1.0** (Current)
- ✅ Dynamic field labeling based on document type
- ✅ Document location required for non-email types
- ✅ Updated validation and display messages
- ✅ Documentation updated

**Version 1.0.0** (Previous)
- Initial release with document type support
- Risk field auto-filled to "FFFFF" for documents

---

**Update Date:** November 1, 2025  
**Status:** ✅ Active  
**Compatibility:** Backward compatible with existing database

*For more information, see GURI_README.md and QUICKSTART.md*

