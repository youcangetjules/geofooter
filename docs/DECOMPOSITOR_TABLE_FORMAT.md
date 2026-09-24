# Decompositor Table Format

## Overview

The GURI Decompositor now displays components in a **table format** with column headers and rows, making it easy to read and analyze each component at a glance.

---

## Layout

### Upper Panel: GURI Summary

Shows overall information:

```
Original GURI: a1b2cxd3e4fxf6g7h8i9xj0k1l2m3xn4oxp5
Total Length: 37 characters
Components Found: 6
✓ VALID GURI - All checks passed
```

### Lower Panel: Components Table

Displays all 6 components in a sortable table:

| Component | Value | Meaning | Expected Length | Actual Length | Length Valid | Hex Valid | Status |
|-----------|-------|---------|-----------------|---------------|--------------|-----------|--------|
| Component 1 (5 hex) | a1b2c | Primary Identifier | 5 | 5 | ✓ | ✓ | Valid |
| Component 2 (5 hex) | d3e4f | Secondary Identifier | 5 | 5 | ✓ | ✓ | Valid |
| Component 3 (8 hex) | f6g7h8i9 | Tertiary Identifier (Long) | 8 | 8 | ✓ | ✓ | Valid |
| Component 4 (8 hex) | j0k1l2m3 | Quaternary Identifier (Long) | 8 | 8 | ✓ | ✓ | Valid |
| Component 5 (3 hex) | n4o | Short Identifier | 3 | 3 | ✓ | ✓ | Valid |
| Component 6 (2 hex) | p5 | Check Suffix | 2 | 2 | ✓ | ✓ | Valid |

---

## Column Definitions

### 1. Component
**Description:** Component identifier with expected hex count  
**Example:** `Component 1 (5 hex)`, `Component 2 (5 hex)`, etc.  
**Purpose:** Identifies which component this row represents

### 2. Value
**Description:** The actual hexadecimal value from the GURI  
**Example:** `a1b2c`, `d3e4f`, `f6g7h8i9`  
**Purpose:** Shows the extracted component value

### 3. Meaning (NEW)
**Description:** Explains the purpose and role of this component  
**Examples:**
- `Primary Identifier - Random hex for uniqueness`
- `Secondary Identifier - Random hex for collision avoidance`
- `Tertiary Identifier (Long) - Extended random hex`
- `Quaternary Identifier (Long) - Extended random hex`
- `Short Identifier - 3-character random hex`
- `Check Suffix - 2-character random hex`

**Purpose:** Helps understand what each component represents in the GURI structure  
**Note:** GURIs use randomly generated hex values. The meanings describe the structural role of each component, not decoded data.

### 4. Expected Length
**Description:** The correct length this component should be  
**Example:** `5`, `8`, `3`, `2`  
**Purpose:** Reference for validation

### 5. Actual Length
**Description:** The actual length of the extracted value  
**Example:** `5`, `8`, `3`, `2`  
**Purpose:** Compare against expected length

### 6. Length Valid
**Description:** Validation indicator for length  
**Values:** `✓` (valid) or `❌` (invalid)  
**Purpose:** Quick visual check for length correctness

### 7. Hex Valid
**Description:** Validation indicator for hexadecimal characters  
**Values:** `✓` (valid) or `❌` (invalid)  
**Purpose:** Quick visual check for valid hex (0-9, a-f, A-F)

### 8. Status
**Description:** Overall component status  
**Values:**
- `Valid` - All checks pass
- `Wrong Length` - Length mismatch
- `Invalid Hex` - Non-hexadecimal characters
**Purpose:** Summary status at a glance

---

## Color Coding

### Row Background Colors

| Color | RGB | Meaning | Trigger |
|-------|-----|---------|---------|
| **Light Green** | #e8f5e9 | Valid | Length ✓ AND Hex ✓ |
| **Light Red** | #ffebee | Invalid Hex | Hex ❌ (regardless of length) |
| **Light Orange** | #fff3e0 | Wrong Length | Length ❌ (but hex valid) |

### Visual Example

```
┌─────────────────────────────────────────────────────────────────────────┐
│ 🟢 Component 1 (5 hex) │ a1b2c │ 5 │ 5 │ ✓ │ ✓ │ Valid           │
├─────────────────────────────────────────────────────────────────────────┤
│ 🟢 Component 2 (5 hex) │ d3e4f │ 5 │ 5 │ ✓ │ ✓ │ Valid           │
├─────────────────────────────────────────────────────────────────────────┤
│ 🟡 Component 3 (8 hex) │ f6g7   │ 8 │ 4 │ ❌│ ✓ │ Wrong Length    │
├─────────────────────────────────────────────────────────────────────────┤
│ 🔴 Component 4 (8 hex) │ j0k1l2mZ│ 8 │ 8 │ ✓ │ ❌│ Invalid Hex    │
└─────────────────────────────────────────────────────────────────────────┘
```

---

## Example Scenarios

### Scenario 1: Perfect GURI

**Input:**
```
a1b2cxd3e4fxf6g7h8i9xj0k1l2m3xn4oxp5
```

**Result:**
- All 6 rows display with **green background**
- All checkmarks show ✓
- Status: "Valid" for all rows
- Summary: "✓ VALID GURI - All checks passed"

### Scenario 2: Wrong Length

**Input:**
```
a1xd3e4fxf6g7h8i9xj0k1l2m3xn4oxp5
```

**Result:**

| Component | Value | Expected | Actual | Length Valid | Hex Valid | Status | Color |
|-----------|-------|----------|--------|--------------|-----------|--------|-------|
| Component 1 | a1 | 5 | 2 | ❌ | ✓ | Wrong Length | 🟡 Orange |
| Component 2 | d3e4f | 5 | 5 | ✓ | ✓ | Valid | 🟢 Green |
| ... | ... | ... | ... | ... | ... | ... | ... |

Summary: "❌ INVALID - Component length mismatch"

### Scenario 3: Invalid Hex

**Input:**
```
a1b2cxd3e4fxf6g7h8i9xj0k1l2m3xn4oxpZ
```

**Result:**

| Component | Value | Expected | Actual | Length Valid | Hex Valid | Status | Color |
|-----------|-------|----------|--------|--------------|-----------|--------|-------|
| Component 1-5 | ... | ... | ... | ✓ | ✓ | Valid | 🟢 Green |
| Component 6 | pZ | 2 | 2 | ✓ | ❌ | Invalid Hex | 🔴 Red |

Summary: "❌ INVALID - Non-hexadecimal characters detected"

### Scenario 4: Wrong Component Count

**Input:**
```
a1b2cxd3e4fxf6g7h8i9
```

**Result:**
- Only 4 rows shown (instead of 6)
- Expected Length shows "?"
- All marked as invalid
- Summary: "❌ INVALID: Expected 6 components, found 4"

---

## Benefits of Table Format

### Quick Scanning
✅ See all components at once  
✅ Compare values side-by-side  
✅ Spot issues immediately with color coding

### Professional Presentation
✅ Clean, organized layout  
✅ Consistent with other tabs (View Records, Search)  
✅ Easy to screenshot for documentation

### Data Analysis
✅ See patterns across components  
✅ Identify which component has issues  
✅ Validate multiple aspects simultaneously

### User-Friendly
✅ Column headers explain each field  
✅ Checkmarks (✓/❌) universally understood  
✅ Status column provides plain English summary

---

## Comparison: Before vs After

### Before (Text Format)
```
Component 1 (5 hex):
  Value: a1b2c
  Length: 5 / 5 ✓
  Format: ✓ Valid hex

Component 2 (5 hex):
  Value: d3e4f
  Length: 5 / 5 ✓
  Format: ✓ Valid hex

[continues for all 6...]
```

**Issues:**
- Hard to compare components
- Takes up vertical space
- Difficult to scan quickly
- No color coding

### After (Table Format)

| Component | Value | Expected | Actual | Length ✓ | Hex ✓ | Status |
|-----------|-------|----------|--------|----------|-------|--------|
| Component 1 (5 hex) | a1b2c | 5 | 5 | ✓ | ✓ | Valid |
| Component 2 (5 hex) | d3e4f | 5 | 5 | ✓ | ✓ | Valid |
| Component 3 (8 hex) | f6g7h8i9 | 8 | 8 | ✓ | ✓ | Valid |
| ... | ... | ... | ... | ... | ... | ... |

**Benefits:**
✅ All components visible at once  
✅ Color-coded rows for instant feedback  
✅ Easy to scan and compare  
✅ Professional appearance  

---

## Usage Tips

### 1. Quick Validation
- **Green rows?** → GURI is valid ✓
- **Any red/yellow?** → Check Status column for details
- **Check summary** → Overall validation result

### 2. Finding Problems
- **Red row?** → Invalid hex character (check Value column)
- **Yellow row?** → Wrong length (compare Expected vs Actual)
- **Less than 6 rows?** → Wrong number of components

### 3. Documentation
- **Screenshot** the table for reports
- **Share** results with team members
- **Reference** when debugging GURI generation

---

## Technical Details

### Widget Type
- **Treeview** (ttk.Treeview)
- Same as View Records and Search tabs
- Supports scrolling for consistency

### Column Widths
- Component: 150px
- Value: 150px
- Expected Length: 120px (centered)
- Actual Length: 100px (centered)
- Length Valid: 100px (centered)
- Hex Valid: 100px (centered)
- Status: 100px (centered)

### Tags (for row coloring)
```python
'valid'   → background='#e8f5e9'  # Light green
'invalid' → background='#ffebee'  # Light red
'warning' → background='#fff3e0'  # Light orange
```

---

## Summary

The table format provides:

✅ **Clear structure** - Column headers define each field  
✅ **Visual feedback** - Color-coded rows show status instantly  
✅ **Easy comparison** - All components visible simultaneously  
✅ **Professional look** - Consistent with rest of GUI  
✅ **Quick validation** - Status column and colors show issues  

**Perfect for validating GURIs quickly and accurately!**

---

**Updated:** November 1, 2025  
**Format Version:** 2.0 (Table-based)  
**Previous Version:** 1.0 (Text-based)

