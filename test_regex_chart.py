import re
content = """<c:dLbls><c:spPr><a:solidFill><a:schemeClr val="lt1"/></a:solidFill><a:ln w="19050" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:srgbClr val="C00000"/></a:solidFill><a:prstDash val="solid"/><a:miter lim="800000"/></a:ln><a:effectLst/></c:spPr><c:txPr><a:bodyPr rot="0" spcFirstLastPara="1" vertOverflow="ellipsis" vert="horz" wrap="square" lIns="38100" tIns="19050" rIns="38100" bIns="19050" anchor="ctr" anchorCtr="1"><a:spAutoFit/></a:bodyPr><a:lstStyle/><a:p><a:pPr><a:defRPr sz="900" b="0" i="0" u="none" strike="noStrike" kern="1200" baseline="0"><a:solidFill><a:schemeClr val="dk1"/></a:solidFill><a:latin typeface="+mn-lt"/><a:ea typeface="+mn-ea"/><a:cs typeface="+mn-cs"/></a:defRPr></a:pPr><a:endParaRPr lang="en-TH"/></a:p></c:txPr>"""

content = content.replace('<a:schemeClr val="lt1"/>', '<a:srgbClr val="C00000"/>')
content = re.sub(r'(<a:ln[^>]*>.*?<a:solidFill>\s*)<a:srgbClr val="C00000"/>(\s*</a:solidFill>)', r'\g<1><a:schemeClr val="accent3"/>\g<2>', content)
content = content.replace('<a:schemeClr val="dk1"/>', '<a:schemeClr val="bg1"/>')
print(content)
