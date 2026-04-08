import re

def build_cx_lvl(data, is_num=True, format_code='General'):
    tag = 'num' if is_num else 'str'
    fmt = f' formatCode="{format_code}"' if is_num else ''
    xml = f'<cx:lvl ptCount="{len(data)}"{fmt}>'
    for i, v in enumerate(data):
        xml += f'<cx:pt idx="{i}">{v}</cx:pt>'
    xml += '</cx:lvl>'
    return xml

content = """<cx:strDim type="cat"><cx:f>User_funnel!$A$2:$A$4</cx:f><cx:lvl ptCount="3"><cx:pt idx="0">Totalclick</cx:pt><cx:pt idx="1">Register</cx:pt><cx:pt idx="2">Player</cx:pt></cx:lvl></cx:strDim><cx:numDim type="val"><cx:f>User_funnel!$C$2:$C$4</cx:f><cx:lvl ptCount="3" formatCode="General"><cx:pt idx="0">644</cx:pt><cx:pt idx="1">565</cx:pt><cx:pt idx="2">252</cx:pt></cx:lvl></cx:numDim>"""

cats = ['Totalclick', 'Register', 'Player']
vals = [100, 50, 25]
f_col_letter = 'B'

content = re.sub(r'<cx:strDim[^>]*>.*?</cx:strDim>', f'<cx:strDim type="cat"><cx:f>User_funnel!$A$2:$A$4</cx:f>{build_cx_lvl(cats, False)}</cx:strDim>', content, flags=re.DOTALL)
content = re.sub(r'<cx:numDim[^>]*>.*?</cx:numDim>', f'<cx:numDim type="val"><cx:f>User_funnel!${f_col_letter}$2:${f_col_letter}$4</cx:f>{build_cx_lvl(vals, True, "General")}</cx:numDim>', content, flags=re.DOTALL)

print(content)
