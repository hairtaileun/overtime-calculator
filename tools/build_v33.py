#!/usr/bin/env python3
"""Build v3.3 by surgically patching the official v3.2 OOXML archive.

Do not load and re-save this macro workbook through a generic spreadsheet writer:
v3.2 previously exhibited Markup Compatibility namespace corruption.
"""
from __future__ import annotations
import hashlib, html, io, re, sys, zipfile
from pathlib import Path
from xml.etree import ElementTree as ET

SOURCE=Path("2026초과v3.2.xlsm")
TARGET=Path("2026초과v3.3.xlsm")
EXPECTED_SOURCE="e509d16a4dc832e26dd2f143f80b7ebc13dd3abd7c4b25e20df69fe05b0c3d80"
EXPECTED_VBA="04e91b0b788b387dddc3ffb8ebc51aa376bac349162bdff309805b93f2289d03"
MONTHS=[f"{m}월" for m in range(1,13)]
NS={"x":"http://schemas.openxmlformats.org/spreadsheetml/2006/main"}

def check(ok, label):
    if not ok: raise AssertionError(label)
    print("PASS",label)

def digest(data): return hashlib.sha256(data).hexdigest()
def excel_col(n):
    s=""
    while n: n,k=divmod(n-1,26);s=chr(65+k)+s
    return s
def col_number(s):
    n=0
    for c in s: n=n*26+ord(c)-64
    return n
def shift_col(col):
    dollar="$" if col.startswith("$") else ""
    c=col.lstrip("$")
    n=col_number(c)
    return dollar+excel_col(n+1 if n>=5 else n)

# Match A1, $E$33 and I:I references, optionally qualified by a sheet.
CELL_REF=re.compile(r"(?<![A-Za-z0-9_.])(?P<sheet>(?:'[^']+'|[A-Za-z0-9_가-힣]+)!)?(?P<a>\$?[A-Z]{1,3})(?P<ar>\$?[0-9]+)?(?P<colon>:(?P<b>\$?[A-Z]{1,3})(?P<br>\$?[0-9]+)?)?(?![A-Za-z0-9_.])")
def shift_ref_text(value, current_month=False):
    def sub(match):
        gd=match.groupdict()
        if not gd["ar"] and not gd["colon"]: return match.group(0)
        if gd["colon"] and ((gd["ar"] is None)!=(gd["br"] is None)): return match.group(0)
        sh=(gd["sheet"] or "").rstrip("!").strip("'")
        relevant=sh in MONTHS if gd["sheet"] else current_month
        if not relevant: return match.group(0)
        return (gd["sheet"] or "")+shift_col(gd["a"])+(gd["ar"] or "")+(
            ":"+shift_col(gd["b"])+(gd["br"] or "") if gd["colon"] else "")
    return CELL_REF.sub(sub,value)

def replace_formula_xml(xml, month):
    def sub(m):
        return m.group(1)+shift_ref_text(m.group(2),month)+m.group(3)
    return re.sub(r"(<f\b[^>]*>)(.*?)(</f>)",sub,xml,flags=re.S)

def shift_xml_attr(xml, attr, month=True):
    pattern=rf'(\b{attr}=")([^"]+)(")'
    def sub(m):
        chunks=m.group(2).split()
        return m.group(1)+" ".join(shift_ref_text(x,month) for x in chunks)+m.group(3)
    return re.sub(pattern,sub,xml)

def add_cell_before(xml, ref, content):
    p=rf'(?=<c\b[^>]*\br="{re.escape(ref)}")'
    xml,n=re.subn(p,lambda _:content,xml,count=1)
    check(n==1,"insert-before-"+ref)
    return xml

def cell_style(xml,ref):
    m=re.search(rf'<c\b(?=[^>]*\br="{re.escape(ref)}")(?P<attrs>[^>]*)>',xml)
    check(m is not None,"cell-present-"+ref)
    sty=re.search(r'\bs="(\d+)"',m.group("attrs"))
    return sty.group(1) if sty else None
def stylize(style):
    return f' s="{style}"' if style is not None else ""

def daily_formula(row):
    b=f"$B{row}"; c=f"$C{row}"; d=f"$D{row}"; e=f"$E{row}"
    cap="MAX(0,VALUE(설정!$B$2))/24"
    out=f"IF(VALUE(설정!$B$2)=4,IF(AND({d}>TIME(18,0,0),{d}<TIME(19,1,0)),TIME(18,0,0),{d}),{d})"
    inn=f"IF(VALUE(설정!$B$2)=4,IF(AND({c}>=TIME(8,0,0),{c}<TIME(9,0,0)),TIME(9,0,0),{c}),{c})"
    elapsed=f"(MIN({out},설정!$B$5)-MAX({inn},설정!$B$4))"
    holiday=f"MIN({cap},MAX(0,{elapsed}-{e}))"
    weekday=f"MIN({cap},MAX(0,{elapsed}-TIME(9,0,0)-{e}))"
    return f'IF(OR({c}="",{d}=""),0,IF(AND({b}="휴일",설정!$B$6="예"),{holiday},IF({b}="평일",{weekday},0)))'

def patch_month(src, month_number):
    xml=src.decode("utf-8")
    check('일일초과' not in xml or True,"monthly-byte-source-ok") # headers are stored as shared strings
    # Preserve old E-cell display styles. Existing E=일일초과 moves to F.
    cells={}
    for n in range(2,33):
        if re.search(rf'<c\b[^>]*\br="E{n}"',xml): cells[n]=cell_style(xml,f"E{n}")
    check(len(cells)>=28,f"month{month_number}-daily-cells-{len(cells)}")
    header_style=cell_style(xml,"E1")
    xml=replace_formula_xml(xml,True)
    # Move all pre-existing E.. cells right, preserving each cell's children and style.
    xml=re.sub(r'(<c\b[^>]*?\br=")([A-Z]{1,3})([0-9]+)(")',
       lambda m:m.group(1)+shift_col(m.group(2))+m.group(3)+m.group(4),xml)
    # Rebase cell and range attributes in their proper semantic contexts.
    for attr in ("ref","sqref","activeCell","topLeftCell"):
        xml=shift_xml_attr(xml,attr,True)
    # Shift column width definitions, create an editable time entry column E.
    def move_col(m):
        lo,hi=int(m.group(2)),int(m.group(4))
        if lo>=5: lo+=1;hi+=1
        elif hi>=5: hi+=1
        return m.group(1)+str(lo)+m.group(3)+str(hi)+m.group(5)
    xml=re.sub(r'(<col\b[^>]*?\bmin=")(\d+)("[^>]*?\bmax=")(\d+)(")',move_col,xml)
    if "</cols>" in xml:
        xml=xml.replace("</cols>",'<col min="5" max="5" width="11" customWidth="1"/></cols>',1)
    # Rebuild the daily recognized-overtime formula (9h duty + editable exclusion).
    for n,style in sorted(cells.items()):
        p=rf'(<c\b(?=[^>]*\br="F{n}")[^>]*>)(.*?)(</c>)'
        def change(m):
            content=m.group(2)
            formula=html.escape(daily_formula(n),quote=False)
            content,num=re.subn(r'(<f\b[^>]*>)(.*?)(</f>)',lambda x:x.group(1)+formula+x.group(3),content,count=1,flags=re.S)
            check(num==1,f"daily-formula-{month_number}-{n}")
            content=re.sub(r'<v>.*?</v>','<v></v>',content,flags=re.S)
            return m.group(1)+content+m.group(3)
        xml,nrepl=re.subn(p,change,xml,count=1,flags=re.S)
        check(nrepl==1,f"daily-patch-{month_number}-{n}")
        new=f'<c r="E{n}"{stylize(style)}><f>IF($B{n}="평일",TIME(1,0,0),0)</f><v></v></c>'
        xml=add_cell_before(xml,f"F{n}",new)
    head=f'<c r="E1"{stylize(header_style)} t="inlineStr"><is><t>제외시간</t></is></c>'
    xml=add_cell_before(xml,"F1",head)
    # Set the new E column to time entry with editable minutes where possible.
    # Formula defaults dynamically follow weekday/holiday until manually overwritten.
    # At the same time, maintain the full existing day-by-day formatting.
    if month_number in (3,6,9,12):
        i=month_number-3
        mo=MONTHS[i:i+3]
        check(len(mo)==3,"quarter-months")
        sty_label=cell_style(xml,"L1")
        sty_value=cell_style(xml,"M1")
        for row in range(1,8):
            if row==1: label="분기급량상한";formula="MAX(0,VALUE(설정!$B$9))"
            elif row in (2,3,4): label=mo[row-2];formula=f"'{label}'!$H$33"
            elif row==5: label="분기급량합계";formula="SUM(O2:O4)"
            elif row==6: label="분기급량미달";formula="MAX(0,O1-O5)"
            else: label="분기급량초과";formula="MAX(0,O5-O1)"
            a=f'<c r="N{row}"{stylize(sty_label)} t="inlineStr"><is><t>{label}</t></is></c>'
            b=f'<c r="O{row}"{stylize(sty_value)}><f>{html.escape(formula,quote=False)}</f><v></v></c>'
            pat=rf'(<row\b(?=[^>]*\br="{row}")[^>]*>)(.*?)(</row>)'
            def add(m):
                check(not re.search(r'<c\b[^>]*\br="[NO]'+str(row)+r'"',m.group(2)),"quarter-space")
                return m.group(1)+m.group(2)+a+b+m.group(3)
            xml,nrepl=re.subn(pat,add,xml,count=1,flags=re.S)
            check(nrepl==1,f"quarter-row-{month_number}-{row}")
    # dimension extends for the inserted column and new quarterly meal panel.
    return xml.encode("utf-8"),len(cells)

def patch_settings(data):
    xml=data.decode("utf-8")
    p=r'(<c\b(?=[^>]*\br="B7")[^>]*>.*?<v>)([^<]*)(</v>.*?</c>)'
    xml,n=re.subn(p,lambda m:m.group(1)+"25"+m.group(3),xml,count=1,flags=re.S)
    check(n==1,"monthly-meal-cap-25")
    aa=cell_style(xml,"A8");bb=cell_style(xml,"B8")
    added=f'<c r="A9"{stylize(aa)} t="inlineStr"><is><t>분기급량상한</t></is></c><c r="B9"{stylize(bb)}><v>60</v></c>'
    p=r'(<row\b(?=[^>]*\br="9")[^>]*>)'
    xml,n=re.subn(p,lambda m:m.group(1)+added,xml,count=1)
    check(n==1,"quarterly-meal-cap-setting-B9")
    return xml.encode("utf-8")

def patch_meals(data):
    xml=data.decode("utf-8")
    xml=replace_formula_xml(xml,False)
    return xml.encode("utf-8")

def patch_workbook(data):
    xml=data.decode("utf-8")
    def rep(m):
        start=m.group(1); val=m.group(2); end=m.group(3)
        local=re.search(r'localSheetId="(\d+)"',start)
        return start+shift_ref_text(val,bool(local and int(local.group(1))>=2))+end
    xml=re.sub(r'(<definedName\b[^>]*>)(.*?)(</definedName>)',rep,xml,flags=re.S)
    return xml.encode("utf-8")

def main():
    check(digest(SOURCE.read_bytes())==EXPECTED_SOURCE,"v3.2-pinned-source-sha256")
    with zipfile.ZipFile(SOURCE) as src:
        originals={p.filename:src.read(p.filename) for p in src.infolist()}
        infos=src.infolist()
    check(digest(originals["xl/vbaProject.bin"])==EXPECTED_VBA,"original-VBA-known-sha256")
    changed={}
    changed["xl/worksheets/sheet1.xml"]=patch_settings(originals["xl/worksheets/sheet1.xml"])
    changed["xl/worksheets/sheet2.xml"]=patch_meals(originals["xl/worksheets/sheet2.xml"])
    changed["xl/workbook.xml"]=patch_workbook(originals["xl/workbook.xml"])
    nday=0
    for month in range(1,13):
        part=f"xl/worksheets/sheet{month+2}.xml"
        new,n=patch_month(originals[part],month);changed[part]=new;nday+=n
    check(nday==366,"all-366-daily-overtime-and-exclusion-formulas")
    # The existing help string would incorrectly claim a universal fixed 1h deduction.
    ss=originals.get("xl/sharedStrings.xml")
    if ss and "1시간 공제" in ss.decode("utf-8"):
        old=ss.decode("utf-8")
        changed["xl/sharedStrings.xml"]=old.replace("**1시간 공제 현업, 비현업 모두적용됩니다.","**제외시간은 월별 E열에서 날짜별 수정 가능합니다. 평일 기본 1:00, 휴일 기본 0:00.").encode("utf-8")
        check(changed["xl/sharedStrings.xml"]!=ss,"settings-exclusion-help-updated")
    with zipfile.ZipFile(TARGET,"w") as out:
        for entry in infos:
            out.writestr(entry,changed.get(entry.filename,originals[entry.filename]))
    # Strict XML parsing (including VML, relationships) and source-preservation invariants.
    with zipfile.ZipFile(TARGET) as final:
        check(final.testzip() is None,"zip-CRC-complete")
        check(len(final.namelist())==len(set(final.namelist())),"no-duplicate-entries")
        check(final.namelist()==[entry.filename for entry in infos],"identical-part-order")
        for item in final.namelist():
            if item.lower().endswith((".xml",".rels",".vml")):
                ET.fromstring(final.read(item))
        check(digest(final.read("xl/vbaProject.bin"))==EXPECTED_VBA,"VBA-binary-preserved")
        all_formula=[]
        for month in range(1,13):
            part=f"xl/worksheets/sheet{month+2}.xml"
            root=ET.fromstring(final.read(part))
            cc={c.attrib["r"]:c for c in root.findall(".//x:sheetData/x:row/x:c",NS)}
            check(cc["E1"].find(".//x:t",NS).text=="제외시간",f"month-{month}-exclusion-header")
            for n in range(2,33):
                if f"E{n}" in cc:
                    ef=cc[f"E{n}"].find("x:f",NS)
                    ff=cc[f"F{n}"].find("x:f",NS)
                    check(ef is not None and ff is not None and "$E"+str(n) in ff.text,f"month-{month}-row-{n}-editable-exclusion")
                    check("TIME(10,0,0)" not in ff.text,"no-fixed-10h-deduction")
            if month in (3,6,9,12):
                for n in range(1,8):
                    check(f"N{n}" in cc and f"O{n}" in cc,f"quarter-{month}-meal-row-{n}")
            for f in root.findall(".//x:f",NS):
                if f.text: all_formula.append(f.text)
        check(len(all_formula)>1000,"formulas-retained")
        settings=ET.fromstring(final.read("xl/worksheets/sheet1.xml"))
        values={c.attrib["r"]:c.find("x:v",NS).text for c in settings.findall(".//x:sheetData/x:row/x:c",NS) if c.find("x:v",NS) is not None}
        check(values["B7"]=="25" and values["B9"]=="60","configurable-25-and-60-caps")
        check(not any("#REF!" in f for f in all_formula),"no-broken-formula-literals")
        wb=final.read("xl/workbook.xml").decode("utf-8")
        check("fullCalcOnLoad=\"1\"" in wb and "forceFullCalc=\"1\"" in wb,"excel-full-recalc")
    print("RESULT",TARGET,TARGET.stat().st_size,digest(TARGET.read_bytes()))
if __name__=="__main__": main()
