#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""打印 docx 里各段落的实际样式与首行缩进，用来验收模板"""
import sys, zipfile, xml.etree.ElementTree as ET
W='{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'
ns={'w':W[1:-1]}
if len(sys.argv) < 2 or sys.argv[1] in ("-h", "--help"):
    print("用法: python3 show_styles.py 报告.docx")
    print("打印每段的样式 / 字体 / 字号 / 缩进，用于定位「这段格式为什么不对」。")
    raise SystemExit(0)
docx = sys.argv[1]
z=zipfile.ZipFile(docx)
sty={}
root=ET.fromstring(z.read('word/styles.xml'))
for s in root.findall('w:style',ns):
    sid=s.get(W+'styleId'); ind=s.find('w:pPr/w:ind',ns); rp=s.find('w:rPr',ns)
    f=rp.find('w:rFonts',ns) if rp is not None else None
    sz=rp.find('w:sz',ns) if rp is not None else None
    it=rp.find('w:i',ns) if rp is not None else None
    sty[sid]=dict(ind=(ind.get(W+'firstLineChars') if ind is not None else None),
                  ea=(f.get(W+'eastAsia') if f is not None else None),
                  sz=(int(sz.get(W+'val'))/2 if sz is not None else None),
                  italic=(it is not None and it.get(W+'val')!='0'))
body=ET.fromstring(z.read('word/document.xml')).find('w:body',ns)
print(f"{'样式':16s} {'字体':8s} {'字号':5s} {'缩进(字符)':10s} 文本")
print('-'*80)
for el in body:
    if el.tag!=W+'p': continue
    ps=el.find('w:pPr/w:pStyle',ns); sid=ps.get(W+'val') if ps is not None else 'Normal'
    d=el.find('w:pPr/w:ind',ns)
    direct=d.get(W+'firstLineChars') if d is not None else None
    info=sty.get(sid,{})
    ch = direct if direct else info.get('ind')
    txt=''.join(t.text or '' for t in el.iter(W+'t'))
    print(f"{sid:16s} {str(info.get('ea')):8s} {str(info.get('sz')):5s} {str(ch):10s} {txt[:40]}")
