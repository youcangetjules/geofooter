# Column Sorting Feature

## Overview

The GURI GUI now supports **sortable columns**. Click any column header to sort the records by that column.

---

## How to Use

### Sort Any Column

Simply **click on any column header** to sort by that column:

```
┌────────────────────────────────────────────────────────┐
│ ID ▲ │ GURI │ Sender │ Recipients │ Subject │ ...    │
├────────────────────────────────────────────────────────┤
│ 1    │ abc  │ alice  │ bob        │ Report  │ ...    │
│ 2    │ def  │ bob    │ charlie    │ Meeting │ ...    │
│ 3    │ ghi  │ charlie│ david      │ Notes   │ ...    │
└────────────────────────────────────────────────────────┘
```

Click "Sender" header:

```
┌────────────────────────────────────────────────────────┐
│ ID │ GURI │ Sender ▲ │ Recipients │ Subject │ ...    │
├────────────────────────────────────────────────────────┤
│ 1  │ abc  │ alice    │ bob        │ Report  │ ...    │
│ 2  │ def  │ bob      │ charlie    │ Meeting │ ...    │
│ 3  │ ghi  │ charlie  │ david      │ Notes   │ ...    │
└────────────────────────────────────────────────────────┘
```

### Toggle Sort Direction

Click the **same column header again** to reverse the sort order:

**First Click (Ascending ▲):**
```
ID ▲ │ GURI │ Sender │ ...
─────┼──────┼────────┼────
1    │ abc  │ alice  │ ...
2    │ def  │ bob    │ ...
3    │ ghi  │ charlie│ ...
```

**Second Click (Descending ▼):**
```
ID ▼ │ GURI │ Sender │ ...
─────┼──────┼────────┼────
3    │ ghi  │ charlie│ ...
2    │ def  │ bob    │ ...
1    │ abc  │ alice  │ ...
```

### Sort Indicators

- **▲ (Up Arrow)** = Ascending sort (A→Z, 1→9, oldest→newest)
- **▼ (Down Arrow)** = Descending sort (Z→A, 9→1, newest→oldest)
- **No Arrow** = Column not sorted

---

## Available Columns

All columns are sortable in both View Records and Search Results tabs:

| Column | Sort Type | Examples |
|--------|-----------|----------|
| **ID** | Numeric | 1, 2, 3... or 100, 99, 98... |
| **GURI** | Alphabetic | abc...xyz or xyz...abc |
| **Sender** | Alphabetic | alice, bob, charlie |
| **Recipients** | Alphabetic | alice, bob, charlie |
| **Subject** | Alphabetic | Meeting, Notes, Report |
| **Date/Time** | Alphabetic | 2025-01-01, 2025-01-02... |
| **Risk/Location** | Alphabetic | HIGH, LOW, MEDIUM |
| **Doc Type** | Alphabetic | Email, Excel, Word |
| **Created At** | Alphabetic | 2025-01-01 10:00, 2025-01-01 11:00... |

---

## Use Cases

### Sort by Sender

Want to see all emails from a specific sender together?

1. Click **"Sender"** header
2. All records grouped by sender alphabetically
3. Easy to find all records from one person

### Sort by Date/Time

Want to see chronological order?

1. Click **"Date/Time"** header
2. Ascending: Oldest first
3. Descending: Newest first

### Sort by Document Type

Want to group all Word documents together?

1. Click **"Doc Type"** header
2. All same types grouped together
3. Easy to see document distribution

### Sort by GURI

Want to find a specific GURI?

1. Click **"GURI"** header
2. GURIs in alphabetical order
3. Quick navigation to specific ID

### Sort by Risk Level

Want to see high-risk items first?

1. Click **"Risk/Location"** header
2. Descending: HIGH → MEDIUM → LOW
3. Focus on critical items

---

## How It Works

### View Records Tab

1. Load records from database
2. Click any column header
3. Records **on current page** are sorted
4. Sort is **maintained** when you navigate pages
5. Refresh keeps the sort order

### Search Results Tab

1. Perform search
2. Results appear
3. Click any column header
4. Search results are sorted
5. Independent from View Records sorting

### Sort Preservation

- Each tab maintains its **own sort state**
- View Records sort is independent from Search Results sort
- Sort state is **preserved** when:
  - Navigating between pages
  - Refreshing data
  - Switching tabs and coming back

### Sort Reset

Sort is reset when:
- Loading new data (new search, new page)
- Closing and reopening the GUI
- Connecting to a different database

---

## Examples

### Example 1: Finding All Emails from a Domain

**Goal:** Find all emails from "@company.com"

1. Go to **Search** tab
2. Search for "company.com"
3. Click **"Sender"** header
4. All senders from that domain grouped alphabetically
5. Easy to browse through

### Example 2: Most Recent Documents First

**Goal:** See newest documents at top

1. Go to **View Records** tab
2. Click **"Created At"** header twice (for descending ▼)
3. Newest records appear at top
4. Most recent work visible immediately

### Example 3: Organize by Document Type

**Goal:** See all Excel files together

1. Click **"Doc Type"** header
2. All document types grouped together
3. Scroll to "Excel Spreadsheet" section
4. All Excel files in one place

### Example 4: Find GURI by Pattern

**Goal:** Find GURIs starting with "a"

1. Click **"GURI"** header (ascending ▲)
2. All GURIs starting with 'a' appear at top
3. Easy to find specific patterns

---

## Tips and Tricks

### Quick Sorting

- **Double-click** doesn't work - single click only
- **Click header** anywhere - entire header is clickable
- **Visual feedback** - arrow appears immediately

### Multi-Column Workflow

1. Sort by **Doc Type** first (group by type)
2. Then sort by **Subject** within each type
3. Or sort by **Date/Time** to see chronological within type

*Note: Each sort replaces the previous one, not cumulative*

### Performance

- Sorting is **instant** for up to 1000 records
- Larger datasets may take a moment
- Only current page is sorted (not entire database)

### Alphabetic vs Numeric

- **ID column**: Sorts numerically (1, 2, 10 not 1, 10, 2)
- **All other columns**: Sort alphabetically
- **Case-insensitive**: "Alice" = "alice"

---

## Keyboard Shortcuts

Currently, sorting is **mouse-only**. Future versions may include:
- Arrow keys to change sort
- Keyboard shortcuts for common sorts
- Tab key to move between headers

---

## Troubleshooting

### "Sort not working"

**Check:**
- Are you clicking the **header** (not the data)?
- Is there data in the table?
- Try clicking a different column first

### "Wrong sort order"

**Possible Causes:**
- Column contains mixed data types
- Alphabetic sort on numbers
- Empty values may appear at top or bottom

**Solution:**
- Click header again to reverse
- Try a different column
- Check data quality

### "Arrow not showing"

**Check:**
- Did the data actually sort?
- Look carefully - arrow is small ▲▼
- Try clicking header again

### "Sort resets after refresh"

**Explanation:**
- This is expected behavior
- Refresh loads data in default order
- Click column header again to re-sort

---

## Technical Details

### Sort Algorithm

- Uses Python's built-in `sort()` method
- Stable sort (preserves relative order of equal items)
- Case-insensitive for alphabetic columns
- Numeric sort for ID column

### Sort State

Each tree (View Records, Search Results) tracks:
- `sort_column` - Which column is sorted
- `sort_reverse` - Ascending (False) or Descending (True)

### Visual Indicators

- **▲** = Unicode U+25B2 (Up Triangle)
- **▼** = Unicode U+25BC (Down Triangle)

### Performance

- **Best case**: O(n log n) - Python's Timsort
- **Typical**: Instant for < 1000 records
- **Large datasets**: May take 1-2 seconds for 10,000+ records

---

## Future Enhancements

Potential future features:

- [ ] Multi-column sort (primary, secondary, tertiary)
- [ ] Remember sort preference between sessions
- [ ] Custom sort orders (e.g., HIGH > MEDIUM > LOW)
- [ ] Right-click header for sort options
- [ ] Keyboard shortcuts for sorting
- [ ] Sort by multiple columns simultaneously
- [ ] Filter + Sort combination
- [ ] Save sort preferences per user

---

## Comparison with Other Features

### Sort vs Filter

- **Sort**: Rearranges existing records
- **Filter**: Shows/hides records
- Can be used together (filter first, then sort)

### Sort vs Search

- **Sort**: Orders all visible records
- **Search**: Finds specific records
- Search results can also be sorted

### Sort vs Pagination

- **Sort**: Works on current page only
- **Pagination**: Navigates between pages
- Sort is maintained across page changes

---

## Best Practices

### For Finding Records

1. Use **Search** to narrow down
2. Then **Sort** to organize results
3. Then browse through sorted list

### For Analysis

1. **Sort by Doc Type** to see distribution
2. **Sort by Date** to see timeline
3. **Sort by Sender** to see user activity

### For Data Quality

1. **Sort by GURI** to spot duplicates
2. **Sort by Subject** to find similar items
3. **Sort by Risk/Location** to verify data

---

## Summary

✅ **Click any column header** to sort  
✅ **Click again** to reverse sort order  
✅ **▲▼ indicators** show sort direction  
✅ **Independent** sorting for each tab  
✅ **All columns** are sortable  
✅ **Instant** sorting for typical datasets  

**Make finding records easier with column sorting!**

---

**Feature Version:** 1.0  
**Added:** November 1, 2025  
**Status:** ✅ Active  
**Compatibility:** View Records & Search Results tabs

