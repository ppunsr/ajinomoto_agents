import re

content = """
<c:f>'User Engagement'!$B$67:$B$123</c:f>
<cx:f>User_funnel!$A$2:$A$10</cx:f>
<c:f>'gameplay_report(score) '!$C$67:$C$123</c:f>
<c:numCache>
  <c:v>10</c:v>
</c:numCache>
<cx:strCache>
  <cx:v>Foo</cx:v>
</cx:strCache>
"""

ranges = {
    'User Engagement': (1, 10),
    'User_funnel': (20, 30),
    'gameplay_report(score) ': (40, 50),
    'gameplay_report(time) ': (60, 70)
}

def replace_formula(match):
    tag = match.group(1)
    full_ref = match.group(2)
    if '!' in full_ref:
        sheet_part, cell_part = full_ref.rsplit('!', 1)
        sheet_name_clean = sheet_part.strip("'")
        
        # Exact match or match without trailing spaces?
        # Let's try exact first. If not found, try stripping spaces.
        target_range = None
        if sheet_name_clean in ranges:
            target_range = ranges[sheet_name_clean]
        else:
            for k, v in ranges.items():
                if k.strip() == sheet_name_clean.strip():
                    target_range = v
                    break
                    
        if target_range:
            new_start, new_end = target_range
            if ':' in cell_part:
                left, right = cell_part.split(':')
                left = re.sub(r'\d+', str(new_start), left)
                right = re.sub(r'\d+', str(new_end), right)
                new_cell_part = f"{left}:{right}"
            else:
                new_cell_part = re.sub(r'\d+', str(new_start), cell_part)
            return f"<{tag}>{sheet_part}!{new_cell_part}</{tag}>"
    return match.group(0)

content = re.sub(r'<(c|cx):f>(.*?)</\1:f>', replace_formula, content)
content = re.sub(r'<(c|cx):numCache>.*?</\1:numCache>', '', content, flags=re.DOTALL)
content = re.sub(r'<(c|cx):strCache>.*?</\1:strCache>', '', content, flags=re.DOTALL)

print(content)
