#!/usr/bin/env python3
"""Read-only, independent numeric regression of the official v3.3 workbook.

The paired UNO driver loads and edits official XLSM in memory without saving
it. This verifier reads its temporary recalculated XLSX output only.
"""
from __future__ import annotations
import argparse,hashlib,math,zipfile
from pathlib import Path
from xml.etree import ElementTree as ET

BASE=Path("2026초과v3.3.xlsm")
EXPECTED_SHA="604684b58bb0c9676d987816ad466fbf6aa10fd74b68de1c791f1138fce4f7a5"
NS={"m":"http://schemas.openxmlformats.org/spreadsheetml/2006/main"}
DAYS={m:29 if m==2 else 30 if m in (4,6,9,11) else 31 for m in range(1,13)}

def cell(path: Path,month: int,ref: str):
    sheet_number=month+2 if month else 1
    with zipfile.ZipFile(path) as z:
        xml=ET.fromstring(z.read(f"xl/worksheets/sheet{sheet_number}.xml"))
        el=xml.find(f'.//m:sheetData/m:row/m:c[@r="{ref}"]',NS)
        if el is None:return None
        v=el.find("m:v",NS)
        if v is not None and v.text is not None:
            if el.get("t") in ("str","s","inlineStr"):return v.text
            try:return float(v.text)
            except ValueError:return v.text
        i=el.find("m:is",NS)
        if i is not None:return "".join(i.itertext())
        return ""

def near(actual,expected,label):
    if not isinstance(actual,(int,float)) or not math.isclose(actual,expected,rel_tol=0,abs_tol=1e-7):
        raise AssertionError(f"{label} expected={expected}, actual={actual!r}")
    print("SCENARIO_PASS",label,round(actual,10))

def verify(folder: Path):
    assert BASE.is_file(),"original XLSM is missing"
    digest=hashlib.sha256(BASE.read_bytes()).hexdigest()
    assert digest==EXPECTED_SHA,(digest,EXPECTED_SHA)
    print("PASS_RELEASE_BINARY_UNCHANGED",digest)
    for mode in ("아니오","예"):
        stem="nonfield" if mode=="아니오" else "field"
        p=folder/f"{stem}.xlsx"
        assert p.is_file(),p
        for month in (3,6,9,12):
            current=15 if month in (3,9) else 30
            first=200 if month==9 else 100
            preceding=first+60
            cap=(max(0,177-preceding) if mode=="아니오" else 67)
            summary=DAYS[month]+2
            near(cell(p,month,"M17"),cap/24,f"{mode}-{month}-applied-month-cap")
            near(cell(p,month,f"F{summary+1}"),max(0,cap-current)/24,f"{mode}-{month}-monthly-under")
            near(cell(p,month,f"F{summary+2}"),max(0,current-cap)/24,f"{mode}-{month}-monthly-over")
            near(cell(p,month,"M9"),60,f"{mode}-{month}-quarter-meal-cap")
            for row in (10,11,12):
                near(cell(p,month,f"M{row}"),25,f"{mode}-{month}-recognized-meal-{row}")
            near(cell(p,month,"M13"),75,f"{mode}-{month}-quarter-meal-total")
            near(cell(p,month,"M14"),0,f"{mode}-{month}-quarter-meal-under")
            near(cell(p,month,"M15"),15,f"{mode}-{month}-quarter-meal-over")
            if mode=="예":
                for row in range(1,8):
                    v=cell(p,month,f"M{row}")
                    if v not in ("",None):raise AssertionError(f"field-mode unused panel M{row}={v!r}")
                print("SCENARIO_PASS",mode,month,"hours-panel-unused")
            else:
                for row,x in enumerate((177,first/24,60/24,current/24),start=1):
                    near(cell(p,month,f"M{row}"),x,f"{mode}-{month}-quarter-hours-{row}")
                total=preceding+current
                near(cell(p,month,"M5"),total/24,f"{mode}-{month}-quarter-hours-total")
                near(cell(p,month,"M6"),max(0,177-total)/24,f"{mode}-{month}-quarter-hours-under")
                near(cell(p,month,"M7"),max(0,total-177)/24,f"{mode}-{month}-quarter-hours-over")
        # Regular months retain standalone 67h cap regardless of employment mode.
        near(cell(p,1,"F34"),0,f"{mode}-january-under-67")
        near(cell(p,1,"F35"),33/24,f"{mode}-january-over-67")
        near(cell(p,2,"F32"),7/24,f"{mode}-february-under-67")
        # Directly edited daily input cells: weekday, holiday and override.
        near(cell(p,3,"E2"),1/24,f"{mode}-weekday-default-1hour")
        near(cell(p,3,"F2"),3/24,f"{mode}-weekday-net-3hours")
        near(cell(p,3,"E3"),1.5/24,f"{mode}-holiday-exclusion-override")
        near(cell(p,3,"F3"),3.5/24,f"{mode}-holiday-net-3.5hours")
        near(cell(p,3,"E4"),0,f"{mode}-holiday-default-zero")
        near(cell(p,3,"F4"),5/24,f"{mode}-holiday-net-5hours")
        # 27 daily meals: 25 recognized, 2 above monthly cap.
        near(cell(p,3,"H33"),25,f"{mode}-march-meal-recognized-25")
        near(cell(p,3,"H35"),2,f"{mode}-march-meal-over-2")
    print("ALL_SCENARIO_TESTS_PASSED")

if __name__=="__main__":
    parser=argparse.ArgumentParser()
    parser.add_argument("--verify",type=Path,required=True)
    args=parser.parse_args()
    verify(args.verify)
