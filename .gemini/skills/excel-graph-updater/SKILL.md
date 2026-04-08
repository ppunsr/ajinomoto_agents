---
name: excel-graph-updater
description: Safely duplicates and updates the charts in an Excel file (like Aji_game copy.xlsx) to point to data for a specific user-requested month (e.g., Feb, Mar), preserving all unsupported chart types.
---
# Excel Graph Updater Skill

This skill provides a script to dynamically update line charts and funnel charts within `Aji_game copy.xlsx` so that their data series point to the row ranges corresponding to a specified month. 

## Key Improvements
1. **Preserves Unsupported Charts**: Instead of saving via `openpyxl` (which destroys Funnel Charts and other extended objects), this script unzips the `.xlsx` file, modifies the raw XML formulas of the charts, and re-zips it. This ensures 100% preservation of all visual elements!
2. **Creates a Duplicate**: The script creates a new file (e.g., `Aji_game copy_Feb.xlsx`) rather than modifying the original in-place.
3. **Preserves Series Mapping**: It retains the correct column mappings for every series.

## Target Sheets
The script searches and updates charts on the following sheets:
- `User_funnel`
- `User Engagement`
- `gameplay_report(score) `
- `gameplay_report(time) `

## How to use this skill
When the user asks to modify or update the graphs/charts in the Excel file for a specific month (e.g., "Feb", "March"), run the bundled Python script:

```bash
python3 scripts/update_graphs.py "Aji_game copy.xlsx" "<Target Month>"
```

### Example
If the user asks to update the graph for February:
```bash
python3 scripts/update_graphs.py "Aji_game copy.xlsx" "Feb"
```

The script will automatically:
1. Load the original file and identify the new row ranges for the given month.
2. Unzip the file and update the `c:f` and `cx:f` reference formulas in all chart XMLs.
3. Clear Excel's cached values so it is forced to recalculate based on the new formulas.
4. Save the result as a **new file** (e.g., `Aji_game copy_Feb.xlsx`).
