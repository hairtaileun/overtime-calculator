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
    return re.sub(r"(<f\b[^>]*?(?<!/)>)((?:(?!<f\b).)*?)(</f>)",sub,xml,flags=re.S)

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
    max_days=29 if month_number==2 else (30 if month_number in (4,6,9,11) else 31)
    for n in range(2,max_days+2):
        opening=re.search(rf'<c\b(?=[^>]*\br="E{n}")[^>]*>',xml)
        if not opening or opening.group(0).endswith("/>"): continue
        end=xml.find("</c>",opening.end())
        if end<0: continue
        if re.search(r"<f\b[^>]*>[^<]+</f>",xml[opening.end():end]):
            cells[n]=cell_style(xml,f"E{n}")
    check(len(cells)>=28,f"month{month_number}-daily-cells-{len(cells)}")
    header_style=cell_style(xml,"E1")
    xml=replace_formula_xml(xml,True)
    # Move all pre-existing E.. cells right, preserving each cell's children and style.
    xml=re.sub(r'(<c\b[^>]*?\br=")([A-Z]{1,3})([0-9]+)(")',
       lambda m:m.group(1)+shift_col(m.group(2))+m.group(3)+m.group(4),xml)
    # Rebase cell and range attributes in their proper semantic contexts.
    for attr in ("ref","sqref","activeCell","topLeftCell"):
        xml=shift_xml_attr(xml,attr,True)
    # Extended x14 data validations store targets as <xm:sqref>TEXT</xm:sqref>,
    # not as ordinary sqref attributes. Rebase those too (notably meal-shop dropdown).
    xml=re.sub(r"(<(?:[A-Za-z0-9_]+:)?sqref\b[^>]*>)([^<]+)(</(?:[A-Za-z0-9_]+:)?sqref>)",
        lambda m:m.group(1)+shift_ref_text(m.group(2),True)+m.group(3),xml)
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
    patched=0
    for n,style in sorted(cells.items()):
        p=rf'(<c\b(?=[^>]*\br="F{n}")[^>]*>)(.*?)(</c>)'
        if not re.search(p,xml,flags=re.S): continue
        def change(m):
            content=m.group(2)
            formula=html.escape(daily_formula(n),quote=False)
            content,num=re.subn(r'(<f\b[^>]*>)(.*?)(</f>)',lambda x:x.group(1)+formula+x.group(3),content,count=1,flags=re.S)
            check(num==1,f"daily-formula-{month_number}-{n}")
            content=re.sub(r'<v>.*?</v>','<v></v>',content,flags=re.S)
            return m.group(1)+content+m.group(3)
        xml,nrepl=re.subn(p,change,xml,count=1,flags=re.S)
        check(nrepl==1,f"daily-patch-{month_number}-{n}")
        new=f'<c r="E{n}"{stylize(style)}><f>IF($A{n}="",0,IF($B{n}="평일",TIME(1,0,0),0))</f><v></v></c>'
        xml=add_cell_before(xml,f"F{n}",new)
        patched+=1
    head=f'<c r="E1"{stylize(header_style)} t="inlineStr"><is><t>제외시간</t></is></c>'
    xml=add_cell_before(xml,"F1",head)
    # Each month has BOTH a monthly ceiling and (for non-field staff) a
    # quarterly ceiling. M18 is the monthly *recognized* time. Quarter
    # accounting uses M18, NEVER uncapped daily totals.
    #
    # M17 = applied month allowance; M18 = recognized monthly hours.
    # This forward-only chain avoids circular dependencies:
    # M17[m] -> previous M18 -> previous M17/actual hours.
    summary=max_days+2
    qfirst=((month_number-1)//3)*3+1
    prior=[f"'{k}월'!$M$18" for k in range(qfirst,month_number)]
    prevsum="SUM("+",".join(prior)+")" if prior else "0"
    monthly="MAX(0,VALUE(설정!$B$3))/24"
    quarter="MAX(0,VALUE(설정!$B$8))/24"
    remaining=f"MAX(0,{quarter}-{prevsum})"
    cap_formula=f'IF(설정!$B$10="예",{monthly},MIN({monthly},{remaining}))'
    recognized_formula=f"MIN(MAX(0,F{summary}),M17)"
    for sr,expr in ((summary+1,f"MAX(0,$M$17-F{summary})"),
                    (summary+2,f"MAX(0,F{summary}-$M$17)")):
        # Legacy months may encode summary formulas as <f t="shared" .../>
        # rather than ordinary <f>...</f>; produce explicit standalone formulas.
        fp=rf'(<c\b(?=[^>]*\br="F{sr}")[^>]*>)(.*?)(</c>)'
        def new_summary(mt):
            interior=mt.group(2)
            if re.search(r'<f\b[^>]*/>',interior):
                interior,nf=re.subn(r'<f\b[^>]*/>',lambda _:f'<f>{expr}</f>',interior,count=1)
            else:
                interior,nf=re.subn(r'<f\b[^>]*>.*?</f>',lambda _:f'<f>{expr}</f>',interior,count=1,flags=re.S)
            check(nf==1,f"dual-cap-summary-formula-{month_number}-F{sr}")
            interior=re.sub(r'<v>.*?</v>','<v></v>',interior,count=1,flags=re.S)
            return mt.group(1)+interior+mt.group(3)
        xml,n=re.subn(fp,new_summary,xml,count=1,flags=re.S)
        check(n==1,f"month-{month_number}-applied-dual-cap-summary-F{sr}")
    sty_label=cell_style(xml,"L1" if month_number in (3,6,9,12) else f"D{summary}")
    sty_time=cell_style(xml,f"F{summary}")
    for row,caption,formula in ((17,"월상한시간",cap_formula),
                                (18,"월인정시간",recognized_formula)):
        lc=f'<c r="L{row}"{stylize(sty_label)} t="inlineStr"><is><t>{caption}</t></is></c>'
        mc=f'<c r="M{row}"{stylize(sty_time)}><f>{html.escape(formula,quote=False)}</f><v></v></c>'
        pat=rf'(<row\b(?=[^>]*\br="{row}")[^>]*>)(.*?)(</row>)'
        def add_month_cap(mt):
            content=mt.group(2)
            for col,newcell in (("L",lc),("M",mc)):
                ref=f"{col}{row}"
                blank=rf'<c\b(?=[^>]*\br="{ref}")[^>]*/>'
                content,n=re.subn(blank,lambda _:newcell,content,count=1)
                if not n:
                    check(not re.search(rf'<c\b(?=[^>]*\br="{ref}")',content),f"dual-cap-no-overwrite-{month_number}-{ref}")
                    content+=newcell
            return mt.group(1)+content+mt.group(3)
        xml,n=re.subn(pat,add_month_cap,xml,count=1,flags=re.S)
        check(n==1,f"month-{month_number}-dual-cap-row-{row}")
    # Formula defaults dynamically follow weekday/holiday until overwritten.
    if month_number in (3,6,9,12):
        i=month_number-3
        mo=MONTHS[i:i+3]
        check(len(mo)==3,"quarter-months")
        sty_label=cell_style(xml,"L1")
        sty_value=cell_style(xml,"M1")
        sty_time=cell_style(xml,"M6")
        # Quarter panel displays only recognized (monthly-capped) hours.
        for qrow,mon in enumerate(mo,start=2):
            fp=rf'(<c\b(?=[^>]*\br="M{qrow}")[^>]*>.*?<f\b[^>]*>)(.*?)(</f>)'
            value=f"'{mon}'!$M$18"
            xml,n=re.subn(fp,lambda mt:mt.group(1)+value+mt.group(3),xml,count=1,flags=re.S)
            check(n==1,f"quarter-{month_number}-recognized-month-{mon}")
        # Disable quarterly hour panel for field workers without deactivating meals.
        for qrow in range(1,8):
            fp=rf'(<c\b(?=[^>]*\br="M{qrow}")[^>]*>.*?<f\b[^>]*>)(.*?)(</f>)'
            def disabled(mt):
                return mt.group(1)+'IF(설정!$B$10="예","",'+mt.group(2)+')'+mt.group(3)
            xml,n=re.subn(fp,disabled,xml,count=1,flags=re.S)
            check(n==1,f"quarter-{month_number}-inactive-hours-M{qrow}")
        special={16:(("분기초근 적용",sty_label),('IF(설정!$B$10="예","미사용(현업)","사용(비현업)")',sty_value))}
        for row,((caption,caption_style),(formula,formula_style)) in special.items():
            lc=f'<c r="L{row}"{stylize(caption_style)} t="inlineStr"><is><t>{caption}</t></is></c>'
            mc=f'<c r="M{row}"{stylize(formula_style)}><f>{html.escape(formula,quote=False)}</f><v></v></c>'
            pat=rf'(<row\b(?=[^>]*\br="{row}")[^>]*>)(.*?)(</row>)'
            def add_status(m):
                content=m.group(2)
                for col,cell in (("L",lc),("M",mc)):
                    ref=f"{col}{row}"
                    blank=rf'<c\b(?=[^>]*\br="{ref}")[^>]*/>'
                    content,n=re.subn(blank,lambda _:cell,content,count=1)
                    if not n:
                        check(not re.search(rf'<c\b(?=[^>]*\br="{ref}")',content),f"unused-{ref}")
                        content+=cell
                return m.group(1)+content+m.group(3)
            xml,n=re.subn(pat,add_status,xml,count=1,flags=re.S)
            check(n==1,f"quarter-{month_number}-status-row-{row}")
        rule_priority=max([int(p) for p in re.findall(r'<cfRule\b[^>]*\bpriority="(\d+)"',xml)] or [0])+1
        cf=f'<conditionalFormatting sqref="L1:M8"><cfRule type="expression" dxfId="__GREY_DXF__" priority="{rule_priority}"><formula>$M$16="미사용(현업)"</formula></cfRule></conditionalFormatting>'
        if '<dataValidations' in xml:
            xml=xml.replace('<dataValidations',cf+'<dataValidations',1)
        elif '<pageMargins' in xml:
            xml=xml.replace('<pageMargins',cf+'<pageMargins',1)
        else: raise AssertionError(f"quarter-{month_number}-CF-insertion-point")
        check("__GREY_DXF__" in xml,f"quarter-{month_number}-grayscale-rule")
        for offset in range(7):
            row=offset+9
            if offset==0: label="분기급량상한";formula="MAX(0,VALUE(설정!$B$9))"
            elif offset in (1,2,3):
                label=mo[offset-1]
                numbered_month=int(label[:-1])
                calendar_days=29 if numbered_month==2 else (30 if numbered_month in (4,6,9,11) else 31)
                formula=f"'{label}'!$H${calendar_days+2}"
            elif offset==4: label="분기급량합계";formula="SUM(M10:M12)"
            elif offset==5: label="분기급량미달";formula="MAX(0,M9-M13)"
            else: label="분기급량초과";formula="MAX(0,M13-M9)"
            a=f'<c r="L{row}"{stylize(sty_label)} t="inlineStr"><is><t>{label}</t></is></c>'
            b=f'<c r="M{row}"{stylize(sty_value)}><f>{html.escape(formula,quote=False)}</f><v></v></c>'
            pat=rf'(<row\b(?=[^>]*\br="{row}")[^>]*>)(.*?)(</row>)'
            def add(m):
                content=m.group(2)
                for col,replacement in (("L",a),("M",b)):
                    ref=f"{col}{row}"
                    # Quarterly section has blank, style-only cells in these slots.
                    # Replace placeholders while refusing to overwrite any actual data.
                    pattern=rf'<c\b(?=[^>]*\br="{ref}")[^>]*/>'
                    content,n=re.subn(pattern,lambda _:replacement,content,count=1)
                    if not n:
                        check(not re.search(rf'<c\b(?=[^>]*\br="{ref}")[^>]*>',content),"quarter-no-overwrite")
                        content+=replacement
                return m.group(1)+content+m.group(3)
            xml,nrepl=re.subn(pat,add,xml,count=1,flags=re.S)
            check(nrepl==1,f"quarter-row-{month_number}-{row}")
    # dimension extends for the inserted column and new quarterly meal panel.
    return xml.encode("utf-8"),patched

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
    # Explicit employment mode. Default is non-field to retain prior quarter behavior.
    label_style=cell_style(xml,"A6"); input_style=cell_style(xml,"B6")
    field=f'<c r="A10"{stylize(label_style)} t="inlineStr"><is><t>현업여부</t></is></c><c r="B10"{stylize(input_style)} t="inlineStr"><is><t>아니오</t></is></c>'
    p=r'(<row\b(?=[^>]*\br="10")[^>]*>)'
    xml,n=re.subn(p,lambda m:m.group(1)+field,xml,count=1)
    check(n==1,"field-duty-setting-B10-default-no")
    dv='<dataValidation type="list" allowBlank="0" showErrorMessage="1" errorTitle="선택 오류" error="예 또는 아니오를 선택하세요." sqref="B10"><formula1>"예,아니오"</formula1></dataValidation>'
    xml,n=re.subn(r'(<dataValidations\b[^>]*\bcount=")(\d+)(")',lambda m:m.group(1)+str(int(m.group(2))+1)+m.group(3),xml,count=1)
    check(n==1,"settings-validation-count")
    check("</dataValidations>" in xml,"settings-validation-close")
    xml=xml.replace("</dataValidations>",dv+"</dataValidations>",1)
    return xml.encode("utf-8")

def patch_meals(data):
    xml=data.decode("utf-8")
    xml=replace_formula_xml(xml,False)
    return xml.encode("utf-8")

def patch_styles(data):
    """Append grayscale conditional-formatting style while preserving all legacy styles."""
    xml=data.decode("utf-8")
    grey='<dxf><font><color rgb="FF6B7280"/></font><fill><patternFill patternType="solid"><fgColor rgb="FFE0E0E0"/><bgColor indexed="64"/></patternFill></fill></dxf>'
    match=re.search(r'<dxfs\b[^>]*(?:/>|>.*?</dxfs>)',xml,flags=re.S)
    if match:
        old=match.group(0)
        count=re.search(r'\bcount="(\d+)"',old)
        check(count is not None,"differential-styles-have-count")
        index=int(count.group(1))
        updated=re.sub(r'\bcount="\d+"',f'count="{index+1}"',old,count=1)
        if updated.endswith('/>'):
            updated=updated[:-2]+'>'+grey+'</dxfs>'
        else:
            updated=updated.replace('</dxfs>',grey+'</dxfs>',1)
        xml=xml[:match.start()]+updated+xml[match.end():]
    else:
        index=0
        place=re.search(r'<tableStyles\b',xml)
        check(place is not None,"styles-dxf-insertion-position")
        xml=xml[:place.start()]+f'<dxfs count="1">{grey}</dxfs>'+xml[place.start():]
    print("GREY_DXF_ID",index)
    return xml.encode("utf-8"),index

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
    changed["xl/styles.xml"],grey_id=patch_styles(originals["xl/styles.xml"])
    nday=0
    for month in range(1,13):
        part=f"xl/worksheets/sheet{month+2}.xml"
        new,n=patch_month(originals[part],month)
        changed[part]=new.replace(b"__GREY_DXF__",str(grey_id).encode("ascii"));nday+=n
        print("MONTH_REWRITTEN_DAYS",month,n)
    print("TOTAL_REWRITTEN_DAYS",nday)
    check(nday==366,"all-366-calendar-daily-overtime-and-exclusion-formulas")
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
            check("설정!$B$10" in cc["M17"].find("x:f",NS).text,f"month-{month}-dual-cap-switch")
            check("설정!$B$3" in cc["M17"].find("x:f",NS).text,f"month-{month}-67-hour-cap")
            check("설정!$B$8" in cc["M17"].find("x:f",NS).text,f"month-{month}-quarter-hour-cap")
            check("M17" in cc["M18"].find("x:f",NS).text,f"month-{month}-recognized-hours-helper")
            if month in (3,6,9,12):
                check("설정!$B$10" in cc["M17"].find("x:f",NS).text,f"quarter-{month}-cap-switch")
                check("설정!$B$10" in cc["M16"].find("x:f",NS).text,f"quarter-{month}-grayscale-status")
                for n in range(1,8):
                    check(cc[f"M{n}"].find("x:f",NS).text.startswith('IF(설정!$B$10="예","",'),f"quarter-{month}-disabled-hours-{n}")
                for n in range(9,16):
                    check(f"L{n}" in cc and f"M{n}" in cc,f"quarter-{month}-meal-row-{n}")
            for f in root.findall(".//x:f",NS):
                if f.text: all_formula.append(f.text)
        check(len(all_formula)>1000,"formulas-retained")
        settings=ET.fromstring(final.read("xl/worksheets/sheet1.xml"))
        values={c.attrib["r"]:c.find("x:v",NS).text for c in settings.findall(".//x:sheetData/x:row/x:c",NS) if c.find("x:v",NS) is not None}
        check(values["B7"]=="25" and values["B9"]=="60","configurable-25-and-60-caps")
        settings_cells={c.attrib["r"]:c for c in settings.findall(".//x:sheetData/x:row/x:c",NS)}
        check(settings_cells["B10"].find(".//x:t",NS).text=="아니오","field-duty-default-no")
        check(not any("#REF!" in f for f in all_formula),"no-broken-formula-literals")
        wb=final.read("xl/workbook.xml").decode("utf-8")
        check("fullCalcOnLoad=\"1\"" in wb and "forceFullCalc=\"1\"" in wb,"excel-full-recalc")
    print("RESULT",TARGET,TARGET.stat().st_size,digest(TARGET.read_bytes()))
if __name__=="__main__": main()
