import zipfile, re, os, shutil, openpyxl, sys
from datetime import datetime, date
import calendar

def to_excel_date(dt):
    if not isinstance(dt, datetime): return 0
    return (dt - datetime(1899, 12, 30)).days

def build_num_ref(sheet, col_letter, start_row, end_row, data, format_code='General'):
    formula = f"'{sheet}'!${col_letter}${start_row}:${col_letter}${end_row}"
    xml = f'<c:numRef><c:f>{formula}</c:f><c:numCache><c:formatCode>{format_code}</c:formatCode><c:ptCount val="{len(data)}"/>'
    for i, v in enumerate(data):
        xml += f'<c:pt idx="{i}"><c:v>{v}</c:v></c:pt>'
    xml += '</c:numCache></c:numRef>'
    return xml

def build_str_ref(sheet, col_letter, start_row, end_row, data):
    formula = f"'{sheet}'!${col_letter}${start_row}:${col_letter}${end_row}"
    xml = f'<c:strRef><c:f>{formula}</c:f><c:strCache><c:ptCount val="{len(data)}"/>'
    for i, v in enumerate(data):
        xml += f'<c:pt idx="{i}"><c:v>{v}</c:v></c:pt>'
    xml += '</c:strCache></c:strRef>'
    return xml

def build_cx_lvl(data, is_num=True, format_code='General'):
    tag = 'num' if is_num else 'str'
    fmt = f' formatCode="{format_code}"' if is_num else ''
    xml = f'<cx:lvl ptCount="{len(data)}"{fmt}>'
    for i, v in enumerate(data):
        xml += f'<cx:pt idx="{i}">{v}</cx:pt>'
    xml += '</cx:lvl>'
    return xml

def get_excel_data(wb, sheet_name, start_row, cols):
    if sheet_name not in wb.sheetnames: return []
    ws = wb[sheet_name]; data = []
    for r in range(start_row, ws.max_row + 1):
        vals = [ws.cell(row=r, column=c).value for c in cols]
        if not any(v is not None for v in vals): break
        data.append(vals)
    return data

def analyze_data(excel_path, month_str):
    wb = openpyxl.load_workbook(excel_path, data_only=True)
    m_target = month_str.lower()[:3]
    m_prev = 'feb' if m_target == 'mar' else 'jan'
    year = 2026 
    month_num = list(calendar.month_abbr).index(month_str[:3].capitalize())
    last_day = calendar.monthrange(year, month_num)[1]
    report_end_date = date(year, month_num, last_day)
    res = {'target_month': m_target, 'prev_month': m_prev, 'end_date': report_end_date}
    
    ws = wb['User Engagement']
    ue_rows = []
    for r in range(2, ws.max_row + 1):
        d = ws.cell(row=r, column=1).value
        if isinstance(d, datetime) and d.strftime('%b').lower() in [m_prev, m_target]:
            ue_rows.append({'row': r, 'date': to_excel_date(d), 'new': ws.cell(row=r, column=2).value, 'ret': ws.cell(row=r, column=3).value})
    res['ue'] = ue_rows
    
    ws_sc = wb['gameplay_report(score) ']
    score_rows = []; total_score_sum = 0
    for r in range(2, ws_sc.max_row + 1):
        d = ws_sc.cell(row=r, column=1).value
        if isinstance(d, datetime):
            val = ws_sc.cell(row=r, column=2).value or 0
            if d.date() <= report_end_date: total_score_sum += val
            if d.strftime('%b').lower() in [m_prev, m_target]: score_rows.append({'row': r, 'date': to_excel_date(d), 'val': val})
    res['score'] = score_rows
    
    ws_ti = wb['gameplay_report(time) ']
    time_rows = []; total_time_sum = 0
    for r in range(2, ws_ti.max_row + 1):
        d = ws_ti.cell(row=r, column=1).value
        if isinstance(d, datetime):
            val = ws_ti.cell(row=r, column=2).value or 0
            if d.date() <= report_end_date: total_time_sum += val
            if d.strftime('%b').lower() in [m_prev, m_target]: time_rows.append({'row': r, 'date': to_excel_date(d), 'val': val})
    res['time'] = time_rows
    
    ws_f = wb['User_funnel']
    f_col = 3 if m_target == 'mar' else 2
    res['funnel'] = {'cats': ['Totalclick', 'Register', 'Player'], 'vals': [ws_f.cell(row=2, column=f_col).value, ws_f.cell(row=3, column=f_col).value, ws_f.cell(row=4, column=f_col).value]}
    
    res['stats'] = {'dau': 0, 'prev_dau': 0, 'mau': 0, 'prev_mau': 0, 'stickiness': 0, 'prev_stickiness': 0, 'total_score': total_score_sum, 'avg_score': 0, 'score_change': 0, 'total_time': total_time_sum, 'avg_time': 0, 'time_change': 0}
    
    ws_ue = wb['User Engagement']
    for r in range(1, 15):
        for c in range(1, 15):
            v = ws_ue.cell(row=r, column=c).value
            if isinstance(v, str):
                if v == 'Monthly active user': 
                    res['stats']['mau'] = ws_ue.cell(row=r, column=c + (4 if m_target == 'mar' else 3)).value
                    res['stats']['prev_mau'] = ws_ue.cell(row=r, column=c + (3 if m_target == 'mar' else 2)).value
                if v == 'user stickiness': 
                    res['stats']['stickiness'] = ws_ue.cell(row=r, column=c + (4 if m_target == 'mar' else 3)).value
                    res['stats']['prev_stickiness'] = ws_ue.cell(row=r, column=c + (3 if m_target == 'mar' else 2)).value
                if 'Daily active' in v or 'Daily actuve' in v:
                    res['stats']['dau'] = ws_ue.cell(row=r, column=c + (4 if m_target == 'mar' else 3)).value
                    res['stats']['prev_dau'] = ws_ue.cell(row=r, column=c + (3 if m_target == 'mar' else 2)).value
            
    click, reg, play = res['funnel']['vals']
    res['stats']['conv_reg'] = (reg/click*100) if click else 0
    res['stats']['drop_off'] = ((reg-play)/reg*100) if reg else 0
    
    for c in range(1, 10):
        v = ws_sc.cell(row=1, column=c).value
        if v == f'AVG({month_str.capitalize()})': res['stats']['avg_score'] = ws_sc.cell(row=2, column=c).value
    prev_sc_label = f'AVG({m_prev.capitalize()})'
    for c in range(1, 10):
        if ws_sc.cell(row=1, column=c).value == prev_sc_label:
            ps = ws_sc.cell(row=2, column=c).value
            if ps: res['stats']['score_change'] = (res['stats']['avg_score'] - ps)/ps*100
            
    for c in range(1, 10):
        v = ws_ti.cell(row=1, column=c).value
        if v == f'Avg ({month_str.capitalize()})': res['stats']['avg_time'] = ws_ti.cell(row=2, column=c).value
    prev_ti_label = f'Avg ({m_prev.capitalize()})'
    for c in range(1, 10):
        if ws_ti.cell(row=1, column=c).value == prev_ti_label:
            pt = ws_ti.cell(row=2, column=c).value
            if pt: res['stats']['time_change'] = (res['stats']['avg_time'] - pt)/pt*100
            
    return res

def update_pptx(excel_path, template_path, output_path, month):
    res = analyze_data(excel_path, month)
    
    import json
    json_path = f"value_data_{month}.json"
    if os.path.exists(json_path):
        with open(json_path, 'r', encoding='utf-8') as jf:
            jdata = json.load(jf)
        
        p3 = next((p for p in jdata.get("pages", []) if p["page_number"] == 3), None)
        p4 = next((p for p in jdata.get("pages", []) if p["page_number"] == 4), None)
        p5 = next((p for p in jdata.get("pages", []) if p["page_number"] == 5), None)

        prev_m = None
        curr_m = None
        if p3:
            ue_c = p3["sections"]["User Engagement"]["comparison"]
            prev_m = ue_c["previous_month"]
            curr_m = ue_c["current_month"]
            
            funnel_m = p3["sections"]["User Funnel"]["metrics"]
            ue_m = p3["sections"]["User Engagement"]["metrics"]
            
            res['stats']['dau'] = float(ue_m["Daily Active Users (Avg.)"][curr_m])
            res['stats']['prev_dau'] = float(ue_m["Daily Active Users (Avg.)"][prev_m])
            res['stats']['mau'] = float(ue_m["Monthly Active Users"][curr_m])
            res['stats']['prev_mau'] = float(ue_m["Monthly Active Users"][prev_m])
            
            res['stats']['stickiness'] = float(str(ue_m["User Stickiness"][curr_m]).replace('%', ''))
            res['stats']['prev_stickiness'] = float(str(ue_m["User Stickiness"][prev_m]).replace('%', ''))
            
            res['stats']['conv_reg'] = float(str(funnel_m["Conversion rate"]).replace('%', ''))
            res['stats']['drop_off'] = float(str(funnel_m["Drop off"]).replace('%', ''))
            
        if p4 and curr_m:
            sc_m = p4["metrics"]["AVG Score per Day"]
            res['stats']['avg_score'] = float(sc_m[curr_m])
            res['stats']['score_change'] = float(str(sc_m["difference"]).replace('%', ''))
            
        if p5 and curr_m:
            ti_m = p5["metrics"]["AVG Time per Day"]
            res['stats']['avg_time'] = float(str(ti_m[curr_m]).replace(' minute', ''))
            res['stats']['time_change'] = float(str(ti_m["difference"]).replace('%', ''))
            
    wb = openpyxl.load_workbook(excel_path, data_only=True)
    temp_dir = 'temp_pptx_gen'
    if os.path.exists(temp_dir): shutil.rmtree(temp_dir)
    with zipfile.ZipFile(template_path, 'r') as zip_ref: zip_ref.extractall(temp_dir)
    
    charts_dir = os.path.join(temp_dir, 'ppt', 'charts')
    for filename in os.listdir(charts_dir):
        if not filename.endswith('.xml'): continue
        path = os.path.join(charts_dir, filename)
        with open(path, 'r', encoding='utf-8') as f: content = f.read()
        orig = content
        
        if filename != 'chartEx1.xml':
            content = re.sub(r'<c:(num|str)Cache>.*?</c:\1Cache>', '', content, flags=re.DOTALL)
            content = re.sub(r'<(c|cx):externalData[^>]*>.*?</\1:externalData>|<(c|cx):externalData[^>]*/>', '', content, flags=re.DOTALL)
            
        if 'User Engagement' in content and res['ue']:
            s, e = res['ue'][0]['row'], res['ue'][-1]['row']
            content = re.sub(r'<c:cat>.*?</c:cat>', f'<c:cat>{build_num_ref("User Engagement", "A", s, e, [r["date"] for r in res["ue"]], "m/d/yy")}</c:cat>', content, flags=re.DOTALL)
            content = re.sub(r'(<c:ser>.*?<c:idx val="0".*?<c:val>).*?(</c:val>)', r'\g<1>' + build_num_ref('User Engagement', 'B', s, e, [r['new'] for r in res['ue']]) + r'\g<2>', content, flags=re.DOTALL)
            content = re.sub(r'(<c:ser>.*?<c:idx val="1".*?<c:val>).*?(</c:val>)', r'\g<1>' + build_num_ref('User Engagement', 'C', s, e, [r['ret'] for r in res['ue']]) + r'\g<2>', content, flags=re.DOTALL)
            content = re.sub(r'<c:strRef>\s*<c:f>.*?\$B\$1</c:f>.*?</c:strRef>', '<c:v>New user</c:v>', content)
            content = re.sub(r'<c:strRef>\s*<c:f>.*?\$C\$1</c:f>.*?</c:strRef>', '<c:v>returning User</c:v>', content)
            
        elif 'gameplay_report(score) ' in content and res['score']:
            s, e = res['score'][0]['row'], res['score'][-1]['row']
            content = re.sub(r'<c:cat>.*?</c:cat>', f'<c:cat>{build_num_ref("gameplay_report(score) ", "A", s, e, [r["date"] for r in res["score"]], "m/d/yy")}</c:cat>', content, flags=re.DOTALL)
            content = re.sub(r'<c:val>.*?</c:val>', f'<c:val>{build_num_ref("gameplay_report(score) ", "B", s, e, [r["val"] for r in res["score"]], '0,"k"')}</c:val>', content, flags=re.DOTALL)
            content = content.replace('<a:schemeClr val="lt1"/>', '<a:srgbClr val="C00000"/>').replace('<a:schemeClr val="dk1"/>', '<a:schemeClr val="bg1"/>')
            
        elif 'gameplay_report(time) ' in content and res['time']:
            s, e = res['time'][0]['row'], res['time'][-1]['row']
            content = re.sub(r'<c:cat>.*?</c:cat>', f'<c:cat>{build_num_ref("gameplay_report(time) ", "A", s, e, [r["date"] for r in res["time"]], "m/d/yy")}</c:cat>', content, flags=re.DOTALL)
            content = re.sub(r'<c:val>.*?</c:val>', f'<c:val>{build_num_ref("gameplay_report(time) ", "B", s, e, [r["val"] for r in res["time"]])}</c:val>', content, flags=re.DOTALL)
            
        elif 'state!' in content:
            rows = get_excel_data(wb, 'state', 12, (1, 2))
            if rows:
                content = re.sub(r'<c:cat>.*?</c:cat>', f'<c:cat>{build_str_ref("state", "A", 12, 11+len(rows), [r[0] for r in rows])}</c:cat>', content, flags=re.DOTALL)
                content = re.sub(r'<c:val>.*?</c:val>', f'<c:val>{build_num_ref("state", "B", 12, 11+len(rows), [r[1] for r in rows])}</c:val>', content, flags=re.DOTALL)
                
        elif 'menu!' in content:
            rows = sorted(get_excel_data(wb, 'menu', 2, (1, 2)), key=lambda x: x[1] or 0, reverse=True)[::-1]
            if rows:
                content = re.sub(r'<c:cat>.*?</c:cat>', f'<c:cat>{build_str_ref("menu", "A", 2, 1+len(rows), [r[0] for r in rows])}</c:cat>', content, flags=re.DOTALL)
                content = re.sub(r'<c:val>.*?</c:val>', f'<c:val>{build_num_ref("menu", "B", 2, 1+len(rows), [r[1] for r in rows])}</c:val>', content, flags=re.DOTALL)
                
        elif 'score!' in content:
            rows = sorted(get_excel_data(wb, 'score', 2, (2, 3)), key=lambda x: x[1] or 0, reverse=True)[:10]
            if rows:
                content = re.sub(r'<c:cat>.*?</c:cat>', f'<c:cat>{build_str_ref("score", "B", 2, 1+len(rows), [r[0] for r in rows])}</c:cat>', content, flags=re.DOTALL)
                content = re.sub(r'<c:val>.*?</c:val>', f'<c:val>{build_num_ref("score", "C", 2, 1+len(rows), [r[1] for r in rows])}</c:val>', content, flags=re.DOTALL)
                
        elif filename == 'chartEx1.xml':
            f_col_letter = 'C' if month.lower().startswith('mar') else 'B'
            content = re.sub(r'<cx:strDim[^>]*>.*?</cx:strDim>', f'<cx:strDim type="cat"><cx:f>User_funnel!$A$2:$A$4</cx:f>{build_cx_lvl(res["funnel"]["cats"], False)}</cx:strDim>', content, flags=re.DOTALL)
            content = re.sub(r'<cx:numDim[^>]*>.*?</cx:numDim>', f'<cx:numDim type="val"><cx:f>User_funnel!${f_col_letter}$2:${f_col_letter}$4</cx:f>{build_cx_lvl(res["funnel"]["vals"], True, "General")}</cx:numDim>', content, flags=re.DOTALL)
            
        if content != orig:
            with open(path, 'w', encoding='utf-8') as f: f.write(content)
            
    slides_dir = os.path.join(temp_dir, 'ppt', 'slides')
    for filename in os.listdir(slides_dir):
        if not filename.endswith('.xml') or '_' in filename: continue
        path = os.path.join(slides_dir, filename)
        with open(path, 'r', encoding='utf-8') as f: content = f.read()
        orig = content
        
        content = content.replace("2025/11/28 – 2026/03/31", f"2025/11/28 - {res['end_date'].strftime('%Y/%m/%d')}")
        target_full = month.capitalize()
        prev_full = calendar.month_name[list(calendar.month_abbr).index(res['prev_month'].capitalize())]
        
        content = re.sub(r'February(</a:t>.*?<a:t>)-2026', fr'{prev_full}\1-2026', content)
        content = content.replace("February-2026", f"{prev_full}-2026")
        
        content = re.sub(r'March(</a:t>.*?<a:t>)-2026', fr'{target_full}\1-2026', content)
        content = content.replace("March-2026", f"{target_full}-2026")
            
        s = res['stats']
        
        # Dynamic Analysis Paragraph replacement
        stick_trend = "improved" if s["stickiness"] >= s["prev_stickiness"] else "declined"
        mau_trend = "increase" if s["mau"] >= s["prev_mau"] else "decline"
        mau_context = "a growing overall user base" if s["mau"] >= s["prev_mau"] else "a shrinking overall user base"
        
        content = re.sub(r'Although user stickiness improved \(11% → ', f'Although user stickiness {stick_trend} ({s["prev_stickiness"]:.0f}% → ', content)
        content = re.sub(r'the decline in Monthly Active Users indicates a shrinking overall user base', f'the {mau_trend} in Monthly Active Users indicates {mau_context}', content)
        
        if 'Daily Active Users' in content:
            import uuid
            idx = content.find('Daily Active Users')
            start_idx = content.rfind('<p:sp>', 0, idx)
            end_idx = content.find('</p:sp>', idx) + len('</p:sp>')
            orig_box = content[start_idx:end_idx]
            
            # Create previous month box (Left side, aligned exactly under the prev month header)
            new_box = orig_box.replace('id="34"', 'id="1034"').replace('name="TextBox 33"', 'name="TextBox Prev"')
            # 6590000 perfectly centers it under the February-2026 header at x=7076002
            new_box = re.sub(r'<a:off x="([0-9]+)" y="([0-9]+)"', r'<a:off x="6590000" y="1840950"', new_box)
            new_box = re.sub(r'id="\{[A-F0-9\-]+\}"', f'id="{{{str(uuid.uuid4()).upper()}}}"', new_box)
            new_box = new_box.replace('3.0</a:t>', f'{s["prev_dau"]:.1f}</a:t>')
            new_box = new_box.replace('>21</a:t>', f'>{s["prev_mau"]:.0f}</a:t>')
            new_box = new_box.replace('>14%</a:t>', f'>{s["prev_stickiness"]:.2f}%</a:t>')
            
            # Update target month box
            target_box = orig_box.replace('3.0</a:t>', f'{s["dau"]:.1f}</a:t>')
            target_box = target_box.replace('>21</a:t>', f'>{s["mau"]:.0f}</a:t>')
            target_box = target_box.replace('>14%</a:t>', f'>{s["stickiness"]:.2f}%</a:t>')
            
            content = content[:start_idx] + new_box + target_box + content[end_idx:]
            
        content = re.sub(r'88% conversion rate', f'{s["conv_reg"]:.0f}% conversion rate', content)
        content = re.sub(r'55% drop off', f'{s["drop_off"]:.0f}% drop off', content)
        content = re.sub(r'55% drop-off', f'{s["drop_off"]:.0f}% drop-off', content)
        
        content = content.replace('7,796,142', f'{s["total_score"]:,}')
        content = content.replace('28,986', f'{s["avg_score"]:,.0f}')
        content = content.replace('(-57.7%)', f'({s["score_change"]:.1f}%)')
        
        content = content.replace('7,876', f'{s["total_time"]:,}')
        content = content.replace('32 minute', f'{s["avg_time"]:.0f} minute')
        content = content.replace('(- 53.1%)', f'({s["time_change"]:.1f}%)')
        
        if 'Hours' in content: 
            content = re.sub(r'\([0-9.]+ Hours\)', f'({s["total_time"]/60:.2f} Hours)', content)
            
        if content != orig:
            with open(path, 'w', encoding='utf-8') as f: f.write(content)
            
    with zipfile.ZipFile(output_path, 'w', zipfile.ZIP_DEFLATED) as zipf:
        for root, ds, fs in os.walk(temp_dir):
            for file in fs:
                fpath = os.path.join(root, file); zipf.write(fpath, os.path.relpath(fpath, temp_dir))
    shutil.rmtree(temp_dir)
    print(f'Saved to {output_path}')

if __name__ == "__main__":
    if len(sys.argv) < 5: sys.exit(1)
    update_pptx(sys.argv[1], sys.argv[2], sys.argv[4], sys.argv[3])
