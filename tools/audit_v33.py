#!/usr/bin/env python3
"""Read-only, fail-closed regression audit for the official v3.3 binary."""
import hashlib, zipfile
from xml.etree import ElementTree as ET
from pathlib import Path

P=Path("2026초과v3.3.xlsm")
NS={"x":"http://schemas.openxmlformats.org/spreadsheetml/2006/main"}
VBA_SHA="04e91b0b788b387dddc3ffb8ebc51aa376bac349162bdff309805b93f2289d03"
def need(x,msg):
    if not x: raise AssertionError(msg)
    print("V33_AUDIT_PASS",msg)
def cells(doc):
    return {c.attrib['r']:c for c in doc.findall(".//x:sheetData/x:row/x:c",NS)}
def formula(cell):
    f=cell.find("x:f",NS) if cell is not None else None
    return (f.text or "") if f is not None else ""
def labels(z):
    if "xl/sharedStrings.xml" not in z.namelist(): return []
    return ["".join(n.itertext()) for n in ET.fromstring(z.read("xl/sharedStrings.xml"))]
def content(cell,shared):
    if cell is None:return ""
    i=cell.find("x:is",NS)
    if i is not None:return "".join(i.itertext())
    v=cell.find("x:v",NS)
    if v is None:return ""
    return shared[int(v.text)] if cell.attrib.get("t")=="s" else (v.text or "")
def main():
    need(P.is_file(),"binary-exists")
    with zipfile.ZipFile(P) as z:
        need(z.testzip() is None,"archive-CRC")
        need(len(z.namelist())==len(set(z.namelist())),"archive-unique-names")
        for name in z.namelist():
            if name.endswith((".xml",".rels",".vml")):ET.fromstring(z.read(name))
        need(hashlib.sha256(z.read("xl/vbaProject.bin")).hexdigest()==VBA_SHA,"vba-preserved")
        shared=labels(z)
        settings=cells(ET.fromstring(z.read("xl/worksheets/sheet1.xml")))
        need(content(settings["B7"],shared)=="25","monthly-meal-cap-25")
        need(content(settings["B8"],shared)=="177","quarter-hours-cap-177-retained")
        need(content(settings["B9"],shared)=="60","quarter-meals-cap-60")
        need("분기급량" in content(settings["A9"],shared),"quarter-meal-setting-caption")
        year_total=0
        for m in range(1,13):
            daycount=29 if m==2 else 30 if m in (4,6,9,11) else 31
            summary=daycount+2
            doc=ET.fromstring(z.read(f"xl/worksheets/sheet{m+2}.xml"))
            d=cells(doc)
            allrefs=[c.attrib["r"] for c in doc.findall(".//x:sheetData/x:row/x:c",NS)]
            need(len(allrefs)==len(set(allrefs)),f"{m}-no-duplicate-cell-addresses")
            need(content(d["E1"],shared)=="제외시간",f"{m}-exclusion-header")
            need(content(d["F1"],shared)=="일일초과",f"{m}-daily-header")
            for row in range(2,daycount+2):
                a,b=formula(d.get(f"E{row}")),formula(d.get(f"F{row}"))
                need("IF($A" in a and "평일" in a and "TIME(1,0,0)" in a and "$B" in a, f"{m}-{row}-editable-default")
                need(f"$E{row}" in b and "TIME(9,0,0)" in b and "TIME(10,0,0)" not in b and "설정!$B$2" in b,f"{m}-{row}-deduct-and-cap")
                year_total+=1
            need(f"SUM(F2:F{summary-1})" in formula(d.get(f"F{summary}")),f"{m}-monthly-hours-total")
            need("설정!$B$7" in formula(d.get(f"H{summary}")),f"{m}-monthly-meal-cap-linked")
            need(f"$F2" in formula(d.get("G2")),f"{m}-holiday-meal-uses-recognized-hours")
            sqrefs=" ".join(x.attrib.get("sqref","") or (x.text or "") for x in doc.iter() if x.tag.rsplit("}",1)[-1] in ("dataValidation","sqref"))
            need("J1:J1048576" in sqrefs,f"{m}-meal-shop-dropdown-shifted")
            if m in (3,6,9,12):
                previous=[f"{n}월" for n in range(m-2,m+1)]
                need("설정!$B$8" in formula(d["M1"]),f"{m}-hours-quarter-cap-preserved")
                need("설정!$B$9" in formula(d["M9"]),f"{m}-meal-quarter-cap")
                for i,n in enumerate(previous,start=10):
                    monthnum=int(n[:-1])
                    ndays=29 if monthnum==2 else 30 if monthnum in (4,6,9,11) else 31
                    need(formula(d[f"M{i}"])==f"'{n}'!$H${ndays+2}",f"{m}-quarter-month-{n}-recognized-meal-row-{ndays+2}")
                need("SUM(M10:M12)" in formula(d["M13"]),f"{m}-quarter-meal-total")
                need("M9-M13" in formula(d["M14"]),f"{m}-quarter-meal-under")
                need("M13-M9" in formula(d["M15"]),f"{m}-quarter-meal-over")
        need(year_total==366,"calendar-366-formulas")
        meal=cells(ET.fromstring(z.read("xl/worksheets/sheet2.xml")))
        u=formula(meal.get("O2"))
        for month in range(1,13):
            need(f"'{month}월'!J:J" in u and f"'{month}월'!K:K" in u,f"{month}-meal-consumption-range-remapped")
        w=z.read("xl/workbook.xml").decode("utf-8")
        need('fullCalcOnLoad="1"' in w and 'forceFullCalc="1"' in w,"excel-auto-full-recalc")
    print("V33_AUDIT_RESULT",hashlib.sha256(P.read_bytes()).hexdigest(),P.stat().st_size)
if __name__=="__main__":main()
