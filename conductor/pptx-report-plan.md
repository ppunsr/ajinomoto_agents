# Implementation Plan: PowerPoint Report Automation

## Objective
Extend the project's capabilities to automatically generate a PowerPoint report based on a provided template (`Merkle Thailand -Ajipanda's Kitchen report- 260331  copy.pptx`). The report will reflect the data and graphs from the processed Excel file (`Aji_game copy_<Month>.xlsx`) for a user-specified month.

## Key Files & Context
- **PPTX Template:** `Merkle Thailand -Ajipanda's Kitchen report- 260331  copy.pptx`
- **Excel Data Source:** `Aji_game copy_<Month>.xlsx` (e.g., `Aji_game copy_March.xlsx`)
- **Key Findings:** Derived from the Excel sheets (User Funnel, User Engagement, Gameplay Reports).

## Implementation Steps

### 1. Research & Analysis
- **Inspect Template:** Manually or via a one-time script, identify the names and indices of charts and text placeholders (shapes) in the 7-page template.
- **Map Excel to PPTX:** Create a mapping between Excel sheets/ranges and PPTX slides/charts.
  - Slide 1: Title/Date
  - Slide 2: User Funnel (Funnel Chart)
  - Slide 3: User Engagement (Line Chart - New vs Returning)
  - Slide 4: Gameplay Report Score (Line Chart)
  - Slide 5: Gameplay Report Time (Line Chart)
  - Slide 6: Static Sheets (State/Menu/Score summaries)
  - Slide 7: Conclusions/Key Findings

### 2. Update Automation Script
- Create `scripts/generate_pptx_report.py` (to be bundled with the skill or kept in the project).
- **Libraries:** Use `python-pptx` to manipulate the template.
- **Chart Data:** Use `chart.replace_data()` with `CategoryChartData` to update charts without losing the template's formatting.
- **Text Replacement:** Identify "Key Finding" text boxes and programmatically update them based on data trends (e.g., "User funnel increased by X% compared to previous month").

### 3. Integrate into Skill
- Update the `excel-graph-updater` skill (or rename it to `aji-report-generator`) to include the PowerPoint generation step.
- The skill will now:
  1. Generate/Update the Excel file for the requested month.
  2. Read the updated Excel file.
  3. Load the PPTX template.
  4. Inject the data into the PPTX.
  5. Save the final report as `Report_<Month>.pptx`.

### 4. Verification & Testing
- Run the full workflow for a sample month (e.g., March).
- Verify that:
  - All 7 pages of the PPTX are preserved.
  - Graphs in PPTX match the graphs in Excel.
  - Key findings are logically updated.
  - Visual layout/template style remains identical.

## Migration & Rollback
- Keep the original PPTX template untouched.
- All generated reports will have unique filenames (`Report_March.pptx`) to prevent accidental data loss.
