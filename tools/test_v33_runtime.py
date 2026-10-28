#!/usr/bin/env python3
"""Independent v3.3 behavior regression. Never save/modify the release artifact.

Create two isolated scenario copies, recalculate using LibreOffice (on copies),
and assert numeric results independently of v3.3's source builder/auditor.
LibreOffice is a cross-engine behavioral check, not Microsoft Excel COM.
"""
from __future__ import annotations
import argparse, math, re, sys, zipfile
from pathlib import Path
from xml.etree import ElementTree as ET

BASE=Path("2026초과v3.3.xlsm")
MAIN_NS={"m":"http://schemas.openxmlformats.org/spreadsheetml/2006/main"}
MONTH_DAY={m:(29 if m==2 else 30 if m in (4,6,9,11) else 31) for m in range(1,13)}

def rewrite_cell(xml: str, ref: str, inside: str) -> str:
    p=rf'<c\b(?=[^>]*\br="{re.escape(ref)}")[^>]*(?:/>|>.*?</c>)'
    def sub(m):
        opening=m.group(0).split(">",1)[0].rstrip("/")
        opening=re.sub(r'\s+t="[^"]+"','',opening)
        if inside.startswith("<is>"):opening+=' t="inlineStr"'
        return opening+">"+inside+"</c>"
    xml,count=re.subn(p,sub,xml,count=1,flags=re.S)
    if count==0:
        # The template intentionally omits many empty C/D input cells.
        column,rownum=re.fullmatch(r'([A-Z]+)([0-9]+)',ref).groups()
        new=f'<c r="{ref}"'+(' t="inlineStr"' if inside.startswith("<is>") else '')+'>'+inside+'</c>'
        rp=rf'(<row\b(?=[^>]*\br="{rownum}")[^>]*>)(.*?)(</row>)'
        def insert(m):
            body=m.group(2)
            def index(col):
                out=0
                for char in col:out=out*26+ord(char)-64
                return out
            for c in re.finditer(rf'<c\b[^>]*\br="([A-Z]+){rownum}"',body):
                if index(c.group(1))>index(column):
                    body=body[:c.start()]+new+body[c.start():]
                    break
            else:body+=new
            return m.group(1)+body+m.group(3)
        xml,n=re.subn(rp,insert,xml,count=1,flags=re.S)
        if n!=1:raise AssertionError("Missing row for "+ref)
    elif count!=1:raise AssertionError("Duplicated cell "+ref)
    return xml

def numeric(xml,ref,value):
    return rewrite_cell(xml,ref,f"<v>{value:.14g}</v>")

def prepare(target: Path,mode: str) -> None:
    assert mode in ("예","아니오")
    assert BASE.is_file(),BASE
    with zipfile.ZipFile(BASE) as z:
        parts={i.filename:z.read(i.filename) for i in z.infolist()}
        meta=z.infolist()
    settings=parts["xl/worksheets/sheet1.xml"].decode("utf-8")
    settings=rewrite_cell(settings,"B10",f"<is><t>{mode}</t></is>")
    parts["xl/worksheets/sheet1.xml"]=settings.encode("utf-8")
    for m in range(1,13):
        part=f"xl/worksheets/sheet{m+2}.xml"
        xml=parts[part].decode("utf-8")
        summary=MONTH_DAY[m]+2
        # 100h then 60h, closing month alternating between 15h and 30h.
        quarter_position=(m-1)%3
        hours=(100,60,15 if ((m-1)//3)%2==0 else 30)[quarter_position]
        xml=numeric(xml,f"F{summary}",hours/24)
        # Every month has 25 recognized meals (a quarter totals 75 > 60).
        xml=numeric(xml,f"H{summary}",25)
        if m==3:
            # Validate three distinct exclusion cases independently of total.
            xml=rewrite_cell(xml,"B2",'<is><t>평일</t></is>')
            xml=rewrite_cell(xml,"C2",'<v>0.25</v>') # 06:00
            xml=rewrite_cell(xml,"D2",'<v>0.791666666666667</v>') # 19:00
            xml=rewrite_cell(xml,"B3",'<is><t>휴일</t></is>')
            xml=rewrite_cell(xml,"C3",'<v>0.416666666666667</v>') # 10:00
            xml=rewrite_cell(xml,"D3",'<v>0.625</v>') # 15:00
            xml=rewrite_cell(xml,"E3",'<v>0.0625</v>') # manually 1:30 excluded
            xml=rewrite_cell(xml,"B4",'<is><t>휴일</t></is>')
            xml=rewrite_cell(xml,"C4",'<v>0.416666666666667</v>')
            xml=rewrite_cell(xml,"D4",'<v>0.625</v>')
        parts[part]=xml.encode("utf-8")
    target.parent.mkdir(parents=True,exist_ok=True)
    with zipfile.ZipFile(target,"w") as z:
        for info in meta:
            z.writestr(info,parts[info.filename])
    print("SCENARIO_PREPARED",target,mode)

def get_cell(path: Path,sheet_number: int,ref: str):
    with zipfile.ZipFile(path) as z:
        root=ET.fromstring(z.read(f"xl/worksheets/sheet{sheet_number}.xml"))
        c=root.find(f'.//m:sheetData/m:row/m:c[@r="{ref}"]',MAIN_NS)
        if c is None:raise AssertionError(f"No cell {sheet_number}!{ref}")
        v=c.find("m:v",MAIN_NS)
        if v is not None and v.text is not None:
            if c.get("t") in ("str","inlineStr","s"):return v.text
            try:return float(v.text)
            except ValueError:return v.text
        i=c.find("m:is",MAIN_NS)
        if i is not None:return "".join(i.itertext())
        return ""

def near(got,expected,label,tol=1e-7):
    if not isinstance(got,(int,float)) or not math.isclose(got,expected,rel_tol=0,abs_tol=tol):
        raise AssertionError(f"{label}: expected {expected!r}, got {got!r}")
    print("SCENARIO_PASS",label,round(got,10))

def verify(folder: Path):
    for mode in ("아니오","예"):
        stem="nonfield" if mode=="아니오" else "field"
        file=folder/f"{stem}.xlsx"
        assert file.is_file(),file
        for m in (3,6,9,12):
            sheet=m+2
            current_hours=15 if m in (3,9) else 30
            cap=(17 if mode=="아니오" else 67)
            summary=MONTH_DAY[m]+2
            near(get_cell(file,sheet,"M17"),cap/24,f"{mode}-{m}month-dynamic-cap")
            near(get_cell(file,sheet,f"F{summary+1}"),max(0,cap-current_hours)/24,f"{mode}-{m}month-under-cap")
            near(get_cell(file,sheet,f"F{summary+2}"),max(0,current_hours-cap)/24,f"{mode}-{m}month-over-cap")
            near(get_cell(file,sheet,"M9"),60,f"{mode}-{m}quarter-meal-cap")
            for row in (10,11,12):near(get_cell(file,sheet,f"M{row}"),25,f"{mode}-{m}meal-month-{row}")
            near(get_cell(file,sheet,"M13"),75,f"{mode}-{m}quarter-meal-total")
            near(get_cell(file,sheet,"M14"),0,f"{mode}-{m}quarter-meal-under")
            near(get_cell(file,sheet,"M15"),15,f"{mode}-{m}quarter-meal-over")
            for row in range(1,8):
                val=get_cell(file,sheet,f"M{row}")
                if mode=="예":
                    if val not in ("",None):raise AssertionError(f"Active quarterly hours in field mode {m} M{row}: {val!r}")
                elif row<=4:
                    near(val, (177,100,60,current_hours)[row-1]/24,f"{mode}-{m}quarter-hours-{row}")
            print("SCENARIO_PASS",mode,m,"quarter-hour-activation")
        # A regular month retains independent 67h monthly cap in both modes.
        near(get_cell(file,3,"F34"),0,f"{mode}-january-monthly-under")
        near(get_cell(file,3,"F35"),33/24,f"{mode}-january-monthly-over")
        near(get_cell(file,4,"F32"),7/24,f"{mode}-february-monthly-under")
        # Editable weekday and holiday exclusion formula scenarios on March.
        near(get_cell(file,5,"E2"),1/24,f"{mode}-weekday-default-exclusion")
        near(get_cell(file,5,"F2"),3/24,f"{mode}-weekday-net-3h")
        near(get_cell(file,5,"E3"),1.5/24,f"{mode}-holiday-manual-exclusion")
        near(get_cell(file,5,"F3"),3.5/24,f"{mode}-holiday-net-3.5h")
        near(get_cell(file,5,"E4"),0,f"{mode}-holiday-default-zero-exclusion")
        near(get_cell(file,5,"F4"),5/24,f"{mode}-holiday-default-5h")
    print("ALL_SCENARIO_TESTS_PASSED")

if __name__=="__main__":
    p=argparse.ArgumentParser()
    p.add_argument("--prepare",type=Path)
    p.add_argument("--verify",type=Path)
    args=p.parse_args()
    if args.prepare:
        prepare(args.prepare/"nonfield.xlsm","아니오")
        prepare(args.prepare/"field.xlsm","예")
    elif args.verify:verify(args.verify)
    else:p.error("Need --prepare or --verify")
