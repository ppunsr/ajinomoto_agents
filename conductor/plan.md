# Implementation Plan: Excel Graph Updater Skill

## Objective
Create a specialized Gemini CLI skill (`excel-graph-updater`) that can automatically modify the graphs for a user-specified month in `Aji_game copy.xlsx`. The skill will target specific sheets ("user funnel", "user_engagement", "gameplay_report(score)", "gane play_report(time)") and will be located in the `.gemini/skills` folder.

## Key Files & Context
- Target Excel file: `Aji_game copy.xlsx`
- Target Sheets: 
  - `user funnel`
  - `user_engagement`
  - `gameplay_report(score)`
  - `gane play_report(time)`
- Skill output location: `.gemini/skills/`

## Implementation Steps
1. **Analyze the Excel File**:
   - Write a Python script to inspect `Aji_game copy.xlsx`.
   - Identify the structure of the charts and how data series are labeled by month (e.g., "Feb", "Jan", etc.) in the specified sheets.
   
2. **Initialize the Skill**:
   - Run the skill-creator initialization script:
     ```bash
     node /usr/local/lib/node_modules/@google/gemini-cli/bundle/builtin/skill-creator/scripts/init_skill.cjs excel-graph-updater --path .gemini/skills/excel-graph-updater
     ```

3. **Develop the Automation Script**:
   - Create `scripts/update_graphs_by_month.py` inside the skill directory.
   - Implement chart updating logic using a library like `openpyxl`.
   - The script will accept the target month (e.g., "Feb", "Mar") as a command-line argument.
   - It will locate the data series for the provided month and update the charts dynamically as requested by the user.

4. **Document the Skill (SKILL.md)**:
   - Provide clear instructions in the skill's `SKILL.md` frontmatter and body on how to trigger and use the skill.
   - Ensure the instructions explain that Gemini CLI should pass the target month to the Python script when executing it.

5. **Package and Install**:
   - Package the skill into a `.skill` file.
   - Provide the option to install it to the workspace or keep it in the folder as requested.

## Verification & Testing
- Use a backup copy of the Excel file.
- Run the python script to update the charts for a sample month (e.g., "Feb").
- Confirm the update was successful across all target sheets.