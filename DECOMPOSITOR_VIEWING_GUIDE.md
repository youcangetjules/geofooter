# How to View the GURI Component Breakdown in Decompositor

## Quick Steps

1. **Launch the GUI**: Run `python guri_gui.py` or double-click `Launch_GURI_GUI.bat`

2. **Navigate to Decompositor Tab**: Click on the **"Decompositor"** tab (4th tab at the top)

3. **Look at the RIGHT side**: The screen is split into two panels:
   - **LEFT**: GURI Decompositor (input and breakdown)
   - **RIGHT**: Database Records (THIS IS WHERE YOU SEE THE BREAKDOWN)

4. **Find the table with component columns**:
   ```
   You should see column headers like:
   GURI (Full) | C1 (5hex) | C2 (5hex) | C3 (8hex) | C4 (8hex) | C5 (3hex) | C6 (2hex) | Subject | Doc Type
   ```

5. **Scroll horizontally if needed**: The table is wide with 9 columns. Use the horizontal scrollbar at the bottom to see all columns.

---

## What You Should See

### Table Layout

```
┌─────────────────────────────────────────────────────────────────────────────────┐
│ ★ Table shows GURI broken into 6 components (C1-C6). Scroll right to see all → │
├─────────────────────────────────────────────────────────────────────────────────┤
│ GURI (Full)                           │ C1    │ C2    │ C3       │ C4       │...│
├───────────────────────────────────────┼───────┼───────┼──────────┼──────────┼───┤
│ a1b2cx3d4e5xf6g7h8i9x0a1b2c3dx4e5x6f │ a1b2c │ 3d4e5 │ f6g7h8i9 │ 0a1b2c3d │...│
│ 1a2b3x4c5d6xe7f8g9h0xi1j2k3l4xm5nx6o │ 1a2b3 │ 4c5d6 │ e7f8g9h0 │ i1j2k3l4 │...│
│ ffabcx12345x00112233x44556677x888x99 │ ffabc │ 12345 │ 00112233 │ 44556677 │...│
└───────────────────────────────────────┴───────┴───────┴──────────┴──────────┴───┘
```

*(Scroll right to see C5, C6, Subject, and Doc Type columns)*

### Example Row

If you have a GURI like: `a1b2cx3d4e5xf6g7h8i9x0a1b2c3dx4e5x6f`

You'll see it broken down in the table as:

| Column | Value |
|--------|-------|
| GURI (Full) | a1b2cx3d4e5xf6g7h8i9x0a1b2c3dx4e5x6f |
| C1 (5hex) | **a1b2c** |
| C2 (5hex) | **3d4e5** |
| C3 (8hex) | **f6g7h8i9** |
| C4 (8hex) | **0a1b2c3d** |
| C5 (3hex) | **4e5** |
| C6 (2hex) | **6f** |
| Subject | Important Email |
| Doc Type | Email |

---

## Troubleshooting

### "I don't see any columns with C1, C2, C3, etc."

**Solution**: You might be looking at the wrong panel or tab.
- Make sure you're on the **Decompositor tab** (4th tab)
- Look at the **RIGHT side panel** (not the left side)
- Look for the panel titled "Database Records"

### "I only see GURI and Subject columns"

**Solution**: You need to scroll horizontally!
- Look for a horizontal scrollbar at the bottom of the table
- Drag it to the right to see C1, C2, C3, C4, C5, C6 columns
- The table is wider than the screen

### "The table is empty"

**Solution**: 
- Check the top of the window - it should say "Connected: SQLite (C:\GeoFooter\guri_records.db)"
- If not connected, click **File > Connect to SQLite** and select `guri_records.db`
- Click the **Refresh** button (🔄 symbol) in the database toolbar
- Make sure you have records in the database (create some in the "Create GURI" tab first)

### "I can see the components but they look wrong"

**Solution**: This is normal if the GURI is malformed.
- Invalid GURIs might not have exactly 6 components
- Components might have wrong lengths
- Empty components will show as blank
- Double-click the row to load it into the decompositor for detailed validation

---

## Interactive Features

### Double-Click to Decompose

1. Find any record in the database table
2. **Double-click** on that row
3. The GURI automatically loads into the LEFT side decompositor
4. Full validation breakdown appears instantly

### Search and Filter

1. Type a search term in the "Search Database" field (searches by sender)
2. Click **Search**
3. Results show with components broken down
4. Click **Clear** to return to full database view

### Navigation

- **Refresh**: Reload current page
- **⬅ Prev**: Previous page of records
- **Next ➡**: Next page of records
- **Page indicator**: Shows current page number
- **Total count**: Shows total records in database

---

## Benefits of Component Breakdown View

✅ **Quick scanning**: See all 6 components at a glance  
✅ **Pattern detection**: Spot patterns across multiple GURIs  
✅ **Error identification**: Quickly find malformed components  
✅ **No manual decomposition needed**: Components automatically extracted  
✅ **Efficient workflow**: Browse and analyze simultaneously  

---

## Still Can't See It?

1. **Close and restart the GUI** - Sometimes a fresh start helps
2. **Check window size** - Make sure the window is maximized or wide enough
3. **Look for the info label** - Above the table, you should see:
   "★ Table shows GURI broken into 6 components (C1-C6). Scroll right to see all columns →"
4. **Try the horizontal scrollbar** - It's at the very bottom of the table

---

**Last Updated**: November 2, 2025  
**Version**: 2.0

