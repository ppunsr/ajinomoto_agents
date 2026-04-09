---
name: excel-data-analyzer
description: Analyzes Ajinomoto Game Excel reports and formats the metrics into a clear JSON structure to easily understand the data before building a PowerPoint.
---
# Excel Data Analyzer Skill

This skill automates the extraction and summarization of monthly data from the Ajinomoto Game Excel reports, outputting it in a clean, standardized JSON format. This allows users to review the calculated metrics (like Conversion Rate, Drop Off, and Month-over-Month differences) before injecting them into a presentation.

## Main Workflow
1. Analyze Excel Data: Run scripts/analyze_excel_to_json.py

## Usage Example
`python .gemini/skills/excel-data-analyzer/scripts/analyze_excel_to_json.py "Aji_game copy.xlsx" "Feb"`
