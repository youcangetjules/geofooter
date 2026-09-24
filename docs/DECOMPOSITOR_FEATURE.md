# GURI Decompositor Feature

## Overview

The **Decompositor** tab allows you to break down a GURI into its individual components by splitting on the "x" delimiter. This is useful for analyzing GURI structure, validating format, and understanding each component.

**NEW:** The Decompositor now includes a built-in database viewer with search functionality. Browse your database records on the right side while decomposing GURIs on the left side. Double-click any record to automatically load its GURI into the decompositor for analysis.

---

## GURI Format

A valid GURI has the following structure:

```
{5hex}x{5hex}x{8hex}x{8hex}x{3hex}x{2hex}
```

**Example:**
```
a1b2cxd3e4fxf6g7h8i9xj0k1l2m3xn4oxp5
```

**Components:**
1. **Component 1**: 5 hexadecimal characters
2. **Component 2**: 5 hexadecimal characters  
3. **Component 3**: 8 hexadecimal characters
4. **Component 4**: 8 hexadecimal characters
5. **Component 5**: 3 hexadecimal characters
6. **Component 6**: 2 hexadecimal characters

**Total Length:** 37 characters (31 hex + 5 delimiters)

---

## How to Use

### 1. Enter a GURI

Four ways to input:

**Method A: Type it in**
```
1. Go to "Decompositor" tab
2. Type or paste GURI in the input field
3. Click "Decompose"
```

**Method B: Paste from clipboard**
```
1. Copy a GURI (Ctrl+C)
2. Go to "Decompositor" tab
3. Click "Paste from Clipboard" button
4. Click "Decompose"
```

**Method C: Press Enter**
```
1. Type GURI in input field
2. Press Enter key
3. Decomposition appears automatically
```

**Method D: Load from Database (NEW)**
```
1. Go to "Decompositor" tab
2. Browse database records in the right panel
3. Double-click any record to load its GURI
4. Decomposition appears automatically
```

### 2. View Results

The decomposition shows two sections:

**A. GURI Summary (top panel):**
```
Original GURI: a1b2cxd3e4fxf6g7h8i9xj0k1l2m3xn4oxp5
Total Length: 37 characters
Components Found: 6
✓ VALID GURI - All checks passed
```

**B. GURI Components Table (bottom panel):**

| Component | Value | Expected Length | Actual Length | Length Valid | Hex Valid | Status |
|-----------|-------|-----------------|---------------|--------------|-----------|--------|
| Component 1 (5 hex) | a1b2c | 5 | 5 | ✓ | ✓ | Valid |
| Component 2 (5 hex) | d3e4f | 5 | 5 | ✓ | ✓ | Valid |
| Component 3 (8 hex) | f6g7h8i9 | 8 | 8 | ✓ | ✓ | Valid |
| Component 4 (8 hex) | j0k1l2m3 | 8 | 8 | ✓ | ✓ | Valid |
| Component 5 (3 hex) | n4o | 3 | 3 | ✓ | ✓ | Valid |
| Component 6 (2 hex) | p5 | 2 | 2 | ✓ | ✓ | Valid |

**Color Coding:**
- 🟢 **Green rows** = Valid components
- 🟡 **Yellow rows** = Wrong length
- 🔴 **Red rows** = Invalid hex characters

### 3. Clear Results

Click **"Clear"** button to:
- Clear the input field
- Clear the results display
- Start fresh

---

## Database Viewer (NEW)

The Decompositor tab now features a **split-pane layout** with the decompositor on the left and a database viewer on the right.

### Layout

```
┌──────────────────────────────────────┬──────────────────────────────────────┐
│  GURI Decompositor (Left)           │  Database Records (Right)            │
├──────────────────────────────────────┼──────────────────────────────────────┤
│  Enter GURI                          │  Search Database                     │
│  [GURI input field]                 │  [Search field]                      │
│                                      │                                      │
│  GURI Summary                        │  🔄 Refresh ⬅ Prev Next ➡          │
│  [Validation details]               │  Page: 1 | Total: 1234               │
│                                      │                                      │
│  GURI Components                     │  GURI | C1 | C2 | C3 | C4 | C5 | C6 │
│  [Component table]                  │  [Shows decomposed components]       │
│                                      │                                      │
└──────────────────────────────────────┴──────────────────────────────────────┘
```

### Database Features

#### Browsing Records

The right panel displays your database records in a searchable table with **decomposed GURI components**:
- **GURI**: The complete unique identifier
- **Comp 1**: Component 1 (5 hex characters)
- **Comp 2**: Component 2 (5 hex characters)
- **Comp 3**: Component 3 (8 hex characters)
- **Comp 4**: Component 4 (8 hex characters)
- **Comp 5**: Component 5 (3 hex characters)
- **Comp 6**: Component 6 (2 hex characters)
- **Subject**: Email subject or document name
- **Doc Type**: Document type (Email, Word, PDF, etc.)

**This allows you to see the decomposed GURI values at a glance without needing to decompose each one individually!**

**Example Table View:**

| GURI | Comp 1 | Comp 2 | Comp 3 | Comp 4 | Comp 5 | Comp 6 | Subject | Doc Type |
|------|--------|--------|--------|--------|--------|--------|---------|----------|
| a1b2cx3d4e5xf6g7h8i9x0a1b2c3dx4e5x6f | a1b2c | 3d4e5 | f6g7h8i9 | 0a1b2c3d | 4e5 | 6f | Important Email | Email |
| 1a2b3x4c5d6xe7f8g9h0xi1j2k3l4xm5nx6o | 1a2b3 | 4c5d6 | e7f8g9h0 | i1j2k3l4 | m5n | 6o | Report.docx | Word Document |

You can quickly scan the components to identify patterns, verify structure, or find specific component values.

#### Navigation

- **🔄 Refresh**: Reload current page
- **⬅ Prev**: Go to previous page
- **Next ➡**: Go to next page
- **Page indicator**: Shows current page number
- **Total count**: Shows total records in database

#### Searching

1. Type search term in "Search" field
2. Click "Search" button
3. View matching records
4. Click "Clear" to return to full database view

**Note:** Search currently searches by sender email address.

#### Loading GURI for Decomposition

**Fastest way to analyze a record:**

1. Locate the record in the database table
2. **Double-click** the record
3. GURI automatically loads into decompositor
4. Decomposition appears instantly

This eliminates the need to manually copy/paste GURIs!

### Use Cases for Database Viewer

#### 1. Quick Component Analysis

```
1. Browse database records
2. Visually scan the 6 component columns
3. Identify patterns or anomalies at a glance
4. No need to decompose individual GURIs
```

#### 2. Verify Recently Created GURIs

```
1. Create a new GURI in "Create GURI" tab
2. Switch to "Decompositor" tab
3. See the decomposed components in the table
4. Double-click to perform full validation
```

#### 3. Batch Component Inspection

```
1. Search for specific sender
2. View all decomposed components in results
3. Identify component patterns or errors
4. Double-click suspicious records for detailed analysis
```

#### 4. Quality Assurance

```
1. Browse records by page
2. Scan component columns for invalid values
3. Check for consistent component lengths
4. Double-click to validate any questionable GURIs
```

#### 5. Testing and Debugging

```
1. Generate test GURIs
2. Search for your test records
3. Visually inspect components in table
4. Verify each component matches expected format
```

#### 6. Component Pattern Recognition

```
1. Browse database records
2. Look for patterns in specific components
3. Identify common component values
4. Analyze GURI generation consistency
```

---

## Validation Features

### Length Validation

Each component is checked for correct length in the table:

✅ **Valid Length:**

| Component | Value | Expected Length | Actual Length | Length Valid | Status |
|-----------|-------|-----------------|---------------|--------------|--------|
| Component 1 (5 hex) | a1b2c | 5 | 5 | ✓ | Valid |

❌ **Invalid Length:**

| Component | Value | Expected Length | Actual Length | Length Valid | Status |
|-----------|-------|-----------------|---------------|--------------|--------|
| Component 1 (5 hex) | a1b | 5 | 3 | ❌ | Wrong Length |

### Hex Character Validation

Each component is checked for valid hexadecimal characters (0-9, a-f, A-F):

✅ **Valid Hex:**

| Component | Value | Hex Valid | Status |
|-----------|-------|-----------|--------|
| Component 1 (5 hex) | a1b2c | ✓ | Valid |

❌ **Invalid Hex:**

| Component | Value | Hex Valid | Status |
|-----------|-------|-----------|--------|
| Component 1 (5 hex) | a1bZc | ❌ | Invalid Hex |

### Component Count Validation

The GURI must have exactly **6 components** (split by 'x'):

❌ **Wrong Count - Summary shows:**
```
Original GURI: a1b2cxd3e4fxf6g7h8i9
Total Length: 23 characters
Components Found: 4
❌ INVALID: Expected 6 components, found 4
```

The table will still show whatever components were found, but marked as invalid.

### Overall Status

The summary panel shows the final validation result:

✅ **All Valid:**
```
✓ VALID GURI - All checks passed
```

❌ **Length Mismatch:**
```
❌ INVALID - Component length mismatch
```

❌ **Invalid Hex:**
```
❌ INVALID - Non-hexadecimal characters detected
```

---

## Use Cases

### 1. Validate GURI Format

**Before saving to external system:**
```
1. Paste GURI into decompositor
2. Click "Decompose"
3. Check for ✓ VALID GURI message
4. If valid, proceed with saving
```

### 2. Debug Generation Issues

**If GURIs seem malformed:**
```
1. Generate a GURI
2. Copy it to decompositor
3. Check each component length
4. Identify which component is wrong
5. Debug the generation code
```

### 3. Analyze GURI Structure

**Understanding the format:**
```
1. Take sample GURIs
2. Decompose several
3. See consistent pattern
4. Understand component sizes
```

### 4. Quality Assurance

**Before migrating data:**
```
1. Export GURIs from old system
2. Decompose each one
3. Validate all components
4. Ensure no corruption occurred
```

### 5. Educational/Training

**Teaching others about GURIs:**
```
1. Show example GURI
2. Decompose it live
3. Explain each component
4. Demonstrate validation
```

---

## Example Decompositions

### Example 1: Valid GURI

**Input:**
```
a1b2cxd3e4fxf6g7h8i9xj0k1l2m3xn4oxp5
```

**Output:**
```
✓ VALID GURI - All components have correct lengths
Total Length: 37 characters
6 components detected
```

### Example 2: Invalid - Too Few Components

**Input:**
```
a1b2cxd3e4fxf6g7h8i9
```

**Output:**
```
❌ INVALID GURI FORMAT
Expected 6 components (separated by 'x'), found 3
```

### Example 3: Invalid - Wrong Length

**Input:**
```
a1xd3e4fxf6g7h8i9xj0k1l2m3xn4oxp5
```

**Output:**
```
Component 1 (5 hex):
  Value: a1
  Length: 2 / 5 ❌

⚠ WARNING - Some components have incorrect lengths
```

### Example 4: Invalid - Non-Hex Characters

**Input:**
```
a1b2cxd3e4fxf6g7h8i9xj0k1l2m3xn4oxpZ
```

**Output:**
```
Component 6 (2 hex):
  Value: pZ
  Length: 2 / 2 ✓
  Format: ❌ Invalid hex characters
```

---

## Buttons and Controls

| Button | Function |
|--------|----------|
| **Decompose** | Breaks down the GURI and displays components |
| **Clear** | Clears input and results |
| **Paste from Clipboard** | Pastes GURI from system clipboard |

### Keyboard Shortcuts

- **Enter** - Decompose (when in input field)
- **Ctrl+V** - Paste (standard)
- **Ctrl+A** - Select all in input field

---

## Status Bar Messages

The status bar shows operation results:

| Message | Meaning |
|---------|---------|
| `Valid GURI decomposed: 6 components` | Success - all valid |
| `Invalid GURI: length mismatch in components` | Warning - wrong lengths |
| `Invalid GURI format` | Error - wrong structure |
| `GURI pasted from clipboard` | Clipboard paste successful |
| `Decompositor cleared` | Fields cleared |
| `Error decomposing GURI` | Exception occurred |

---

## Tips and Best Practices

### Quick Validation

1. ✅ Copy GURI after generation
2. ✅ Paste into decompositor
3. ✅ Verify all green checkmarks
4. ✅ Look for "VALID GURI" message

### Batch Testing

For multiple GURIs:
1. Create a list of GURIs
2. Copy each one
3. Click "Paste from Clipboard"
4. Press Enter
5. Check results
6. Repeat

### Integration with Other Tabs

**From View Records:**
1. Find a GURI in View Records
2. Double-click to view details
3. Copy the GURI
4. Switch to Decompositor tab
5. Paste and decompose

**From Create GURI:**
1. Generate new GURI
2. Copy from result display
3. Switch to Decompositor tab
4. Validate the structure

---

## Color Coding

The decompositor uses **row background colors** for quick visual feedback:

| Color | Meaning | When It Appears |
|-------|---------|-----------------|
| **Light Green** 🟢 | Valid | All checks pass (correct length + valid hex) |
| **Light Red** 🔴 | Invalid Hex | Non-hexadecimal characters detected |
| **Light Orange** 🟡 | Wrong Length | Incorrect component length |

**Column Headers:**

| Column | Description |
|--------|-------------|
| **Component** | Component name and expected hex length |
| **Value** | The actual hex value from the GURI |
| **Expected Length** | The correct length for this component |
| **Actual Length** | The actual length found |
| **Length Valid** | ✓ or ❌ |
| **Hex Valid** | ✓ or ❌ |
| **Status** | Valid, Wrong Length, or Invalid Hex |

---

## Troubleshooting

### "Clipboard Error" Message

**Cause:** Nothing in clipboard or clipboard access denied

**Solution:**
- Copy the GURI first (Ctrl+C)
- Or type it manually

### "Input Required" Warning

**Cause:** Empty input field

**Solution:**
- Enter or paste a GURI
- Or use "Paste from Clipboard" button

### Wrong Number of Components

**Cause:** Missing or extra 'x' delimiters

**Solution:**
- Check GURI carefully
- Should have exactly 5 'x' characters
- Format: xxxxx**x**xxxxx**x**xxxxxxxx**x**xxxxxxxx**x**xxx**x**xx

### Invalid Hex Characters

**Cause:** Non-hexadecimal characters (g-z, special chars)

**Solution:**
- Valid: 0-9, a-f, A-F
- Check for typos
- Regenerate GURI if corrupted

---

## Technical Details

### Delimiter

- **Character:** Lowercase 'x'
- **Count:** Exactly 5 occurrences
- **Position:** Between each component

### Splitting Algorithm

```python
components = guri.split('x')
```

### Validation Logic

1. Split on 'x'
2. Check count == 6
3. For each component:
   - Check length matches expected
   - Check all characters are hexadecimal
4. Reassemble and compare

### Expected Lengths

```python
expected_lengths = [5, 5, 8, 8, 3, 2]
```

---

## FAQ

**Q: Why decompose a GURI?**  
A: To validate structure, debug issues, and understand the format.

**Q: Can I edit components and reassemble?**  
A: Not currently - this is read-only analysis. Use Create GURI tab to generate new ones.

**Q: What if I have a GURI with uppercase hex?**  
A: Works fine - validation accepts both uppercase and lowercase (a-f, A-F).

**Q: Can I decompose multiple GURIs at once?**  
A: No - one at a time. Repeat the process for each GURI.

**Q: Will invalid GURIs show in red?**  
A: Yes - invalid components show ❌ with red text.

**Q: Is the original GURI modified?**  
A: No - decomposition is non-destructive. Original is preserved.

---

## Examples for Testing

### Valid GURIs for Testing

```
a1b2cxd3e4fxf6g7h8i9xj0k1l2m3xn4oxp5
0123456789xabcdefx01234567x89abcdefx012x34
fffffxfffffxffffffffxffffffffxfffxff
00000x11111x22222222x33333333x444x55
```

### Invalid GURIs for Testing

**Too few components:**
```
a1b2cxd3e4fxf6g7h8i9
```

**Too many components:**
```
a1b2cxd3e4fxf6g7h8i9xj0k1l2m3xn4oxp5xextra
```

**Wrong length:**
```
a1xd3e4fxf6g7h8i9xj0k1l2m3xn4oxp5
```

**Non-hex characters:**
```
a1b2cxd3e4fxf6g7h8i9xj0k1l2m3xn4oxpZ
```

---

## Summary

### Core Features

✅ **Split by 'x'** delimiter  
✅ **6 components** expected  
✅ **Table format** with column headers  
✅ **Validates** length and hex format  
✅ **Color-coded rows** (green=valid, red=invalid hex, yellow=wrong length)  
✅ **Summary panel** shows overall validation  
✅ **Clipboard** support  
✅ **Enter key** shortcut  

### NEW: Database Integration

✅ **Built-in database viewer** on the right side  
✅ **Decomposed components displayed** in table columns (Comp 1-6)  
✅ **Visual component inspection** without manual decomposition  
✅ **Searchable records** by sender  
✅ **Pagination** with Prev/Next navigation  
✅ **Double-click to load** GURI into decompositor  
✅ **Automatic full decomposition** when loading from database  
✅ **Split-pane layout** for efficient workflow  

**Use the Decompositor to validate and analyze any GURI in an easy-to-read table format, with instant access to your entire database and visual component breakdown!**

---

**Feature Added:** November 1, 2025  
**Version:** 2.0 (Database Viewer Added)  
**Previous Version:** 1.0  
**Status:** ✅ Active  
**Tab:** Decompositor (Tab 4)

