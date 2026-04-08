# Ajinomoto Agents Reporting Project

## Project Overview
This project contains automation scripts and Gemini CLI skills designed to parse data from Ajinomoto game statistics Excel reports (e.g., `Aji_game copy.xlsx`) and inject it into pre-formatted PowerPoint templates (e.g., `Merkle Thailand -Ajipanda's Kitchen report.pptx`). 

It generates monthly 7-page PowerPoint reports displaying user funnels, user engagement metrics, gameplay scores, and state data. Instead of using standard presentation generation libraries which often strip out custom formatting, the system employs direct XML manipulation of the PPTX archive to achieve 100% style preservation and support for complex, unsupported chart types.

## Main Technologies
- **Python 3**
- **openpyxl**: Used for extracting and analyzing raw numbers and dates from the source Excel files.
- **zipfile / re / xml**: Used for unzipping `.pptx` and `.xlsx` archives and performing surgical regex replacements on XML nodes (like `<c:cat>` and `<c:val>`) to update chart data while maintaining template styles.

## Building and Running

The core workflow is encapsulated in the `aji-report-generator` Gemini skill. The standard process for generating a monthly report is as follows:

1. **Update Graph Data (Excel):**
   Run the update script to point the Excel charts to the targeted month's data, creating a new Excel copy:
   ```bash
   python .gemini/skills/aji-report-generator/scripts/update_graphs.py <input_excel> <target_month> <output_excel>
   # Example: python update_graphs.py "Aji_game copy.xlsx" "Feb" "Aji_game copy_Feb.xlsx"
   ```

2. **Generate PowerPoint Report:**
   Run the PowerPoint generation script, passing the newly created Excel file, the PPTX template, the target month, and the output path:
   ```bash
   python .gemini/skills/aji-report-generator/scripts/generate_pptx.py <input_excel> <template_pptx> <target_month> <output_pptx>
   # Example: python generate_pptx.py "Aji_game copy_Feb.xlsx" "Merkle Template.pptx" "Feb" "Report_Feb.pptx"
   ```

### Other Utilities
- **Data Extraction Testing:** `python extract_data.py <excel_file> <month>` to debug and preview the exact JSON payload being read from Excel.

## Development Conventions
- **Direct XML Patching:** Instead of using libraries like `python-pptx`, the project unzips the PowerPoint template into a temporary directory (e.g., `temp_pptx_gen`), modifies specific `chart*.xml` and `slide*.xml` files using regex strings (`<c:numRef>`, `<c:v>`, `<a:t>`), and zips them back into a valid `.pptx` file.
- **Trend Calculation:** Scripts calculate comparisons (e.g., sticking rate improvements, MAU declines) dynamically based on the target month versus the previous month.
- **Safety:** Always use copies of the original `.xlsx` and `.pptx` templates. Do not overwrite the source template.
