import openpyxl
import json
import sys

def extract_data(excel_path, month):
    wb = openpyxl.load_workbook(excel_path, data_only=True)
    m_target = month.lower()[:3]
    
    data = {
        "Month": month,
        "User_funnel": {},
        "User_Engagement": {}
    }
    
    # 1. User Funnel
    ws_f = wb['User_funnel']
    # Find the column for the target month
    target_col = None
    for c in range(2, 10):
        val = ws_f.cell(row=5, column=c).value  # In previous checks, month was in row 5 ('Month', 'feb', 'march')
        if val and isinstance(val, str) and val.lower().startswith(m_target):
            target_col = c
            break
            
    if target_col is None:
        # Fallback to hardcoded if not found in row 5
        target_col = 3 if m_target == 'mar' else 2
        
    data["User_funnel"]["Totalclick"] = ws_f.cell(row=2, column=target_col).value
    data["User_funnel"]["Register"] = ws_f.cell(row=3, column=target_col).value
    data["User_funnel"]["Player"] = ws_f.cell(row=4, column=target_col).value

    # 2. User Engagement
    ws_ue = wb['User Engagement']
    
    # Search for the target month column for DAU, MAU, Stickiness
    # In earlier checks, 'Daily actuve user', 'Monthly active user', 'user stickiness' were in column H (col 8)
    # with Jan in col I (9), Feb in col J (10), March in col K (11)
    
    target_col_ue = None
    for r in range(1, 15):
        for c in range(1, 15):
            val = ws_ue.cell(row=r, column=c).value
            if isinstance(val, str) and val.lower().startswith(m_target):
                # We found a column header for the month, but is it the right one?
                # Usually row 6 has 'Month', 'Dec', 'Jan', 'feb', 'march'
                if ws_ue.cell(row=r, column=c-1).value == 'Month' or ws_ue.cell(row=r, column=c-2).value == 'Month' or ws_ue.cell(row=r, column=c-3).value == 'Month':
                     target_col_ue = c
                     break
        if target_col_ue:
             break
             
    if not target_col_ue:
         # Fallback based on previous analysis
         target_col_ue = 10 if m_target == 'feb' else 11 # J for Feb, K for March
         
    for r in range(1, 15):
        label = ws_ue.cell(row=r, column=8).value # Column H
        if isinstance(label, str):
            if 'Daily actuve' in label or 'Daily active' in label:
                data["User_Engagement"]["Daily_active_user"] = ws_ue.cell(row=r, column=target_col_ue).value
            elif 'Monthly active' in label:
                data["User_Engagement"]["Monthly_active_user"] = ws_ue.cell(row=r, column=target_col_ue).value
            elif 'user stickiness' in label:
                data["User_Engagement"]["user_stickiness"] = ws_ue.cell(row=r, column=target_col_ue).value

    # Save to JSON
    output_filename = f"value_data_{month}.json"
    with open(output_filename, 'w') as f:
        json.dump(data, f, indent=4)
        
    print(f"Data extracted to {output_filename}:")
    print(json.dumps(data, indent=4))

if __name__ == "__main__":
    if len(sys.argv) < 3:
        print("Usage: python extract_data.py <excel_path> <month>")
        sys.exit(1)
    extract_data(sys.argv[1], sys.argv[2])
