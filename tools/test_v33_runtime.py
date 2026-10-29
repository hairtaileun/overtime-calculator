#!/usr/bin/env python3
"""Independent read-only numerical verification of official macro-enabled v3.3.

Eight scenarios exercise the SAME release binary with values edited through UNO.
The XLSM is never saved and the verifier examines only temporary XLSX copies.
"""
from __future__ import annotations
import argparse,hashlib,math,zipfile
from pathlib import Path
from xml.etree import ElementTree as ET

BASE=Path("2026초과v3.3.xlsm")
EXPECTED_SHA="1a5f1cc9cd28d68c3e8151c838b97a62bd37396ad77c2066f730cdc7b6d8d716"
NS={"m":"http://schemas.openxmlformats.org/spreadsheetml/2006/main"}
DAYS={m:29 if m==2 else 30 if m in (4,6,9,11) else 31 for m in range(1,13)}
SCENARIOS=(("nonfield","아니오",177,25,60),
           ("field","예",177,25,60),
           ("nonfield_quarter100","아니오",100,25,60),
           ("nonfield_low_prior","아니오",177,25,60),
           ("nonfield_fullquarter","아니오",177,25,60),
           ("nonfield_meal40","아니오",177,25,40),
           ("field_meal40","예",177,25,40),
           ("nonfield_meal20","아니오",177,20,60))

def cell(path: Path,month: int,ref: str):
    with zipfile.ZipFile(path) as z:
        x=ET.fromstring(z.read(f"xl/worksheets/sheet{month+2}.xml"))
        c=x.find(f'.//m:sheetData/m:row/m:c[@r="{ref}"]',NS)
        if c is None:return None
        v=c.find("m:v",NS)
        if v is not None and v.text is not None:
            if c.get("t") in ("str","s","inlineStr"):return v.text
            try:return float(v.text)
            except ValueError:return v.text
        i=c.find("m:is",NS)
        if i is not None:return "".join(i.itertext())
        return ""

def near(actual,expected,label):
    if not isinstance(actual,(int,float)) or not math.isclose(actual,expected,abs_tol=1e-7,rel_tol=0):
        raise AssertionError(f"{label}: expected={expected!r} got={actual!r}")
    print("SCENARIO_PASS",label,round(actual,10))

def raw_hours(stem,month):
    pos=(month-1)%3
    h=(100,60,15 if ((month-1)//3)%2==0 else 70)[pos]
    if month==7:h=200
    if stem=="nonfield_low_prior" and month<=3:h={1:20,2:20,3:75}[month]
    if stem=="nonfield_fullquarter" and month<=3:h=67
    return h

def run(folder):
    assert BASE.is_file(),"Official v3.3 release file not present"
    sha=hashlib.sha256(BASE.read_bytes()).hexdigest()
    assert sha==EXPECTED_SHA,f"Release digest differs: {sha}"
    print("PASS_OFFICIAL_XLSM_SHA256_UNMODIFIED",sha)
    for stem,mode,qcap,meal_month_cap,meal_quarter_cap in SCENARIOS:
        path=folder/f"{stem}.xlsx"
        assert path.is_file(),f"Missing UNO recalculated result {path}"
        recognized=[]
        meal_recognized=[]
        for m in range(1,13):
            pos=(m-1)%3
            if pos==0:
                recognized=[]
                meal_recognized=[]
            raw=raw_hours(stem,m)
            prior=sum(recognized)
            cap=67 if mode=="예" else min(67,max(0,qcap-prior))
            approved=min(raw,cap)
            summary=DAYS[m]+2
            near(cell(path,m,f"F{summary}"),raw/24,f"{stem}-{m}-raw-month-hours")
            near(cell(path,m,"M17"),cap/24,f"{stem}-{m}-applied-month-cap")
            near(cell(path,m,"M18"),approved/24,f"{stem}-{m}-recognized-month-hours")
            near(cell(path,m,f"F{summary+1}"),max(0,cap-raw)/24,f"{stem}-{m}-monthly-under")
            near(cell(path,m,f"F{summary+2}"),max(0,raw-cap)/24,f"{stem}-{m}-monthly-over")
            recognized.append(approved)
            raw_meal=27 if m==3 else 25
            applied_meal_cap=min(meal_month_cap,max(0,meal_quarter_cap-sum(meal_recognized)))
            accepted_meal=min(raw_meal,applied_meal_cap)
            near(cell(path,m,f"H{summary}"),accepted_meal,f"{stem}-{m}-meal-recognized")
            near(cell(path,m,f"H{summary+1}"),max(0,applied_meal_cap-accepted_meal),f"{stem}-{m}-meal-under")
            near(cell(path,m,f"H{summary+2}"),max(0,raw_meal-applied_meal_cap),f"{stem}-{m}-meal-over")
            meal_recognized.append(accepted_meal)
            if pos==2:
                near(cell(path,m,"M9"),meal_quarter_cap,f"{stem}-{m}-quarter-meal-cap")
                for row,expected_meal in enumerate(meal_recognized,start=10):
                    near(cell(path,m,f"M{row}"),expected_meal,f"{stem}-{m}-meal-month-{row}")
                meal_total=sum(meal_recognized)
                near(cell(path,m,"M13"),meal_total,f"{stem}-{m}-quarter-meals-accepted")
                near(cell(path,m,"M14"),max(0,meal_quarter_cap-meal_total),f"{stem}-{m}-quarter-meals-left")
                meal_only=sum(min(27 if x==3 else 25,meal_month_cap) for x in (m-2,m-1,m))
                near(cell(path,m,"M15"),max(0,meal_only-meal_quarter_cap),f"{stem}-{m}-quarter-only-meal-over")
                if mode=="예":
                    for row in range(1,8):
                        value=cell(path,m,f"M{row}")
                        if value not in ("",None):
                            raise AssertionError(f"{stem}-{m}-hours-panel-not-disabled M{row}={value!r}")
                    print("SCENARIO_PASS",stem,m,"field-quarter-hours-panel-unused")
                else:
                    near(cell(path,m,"M1"),qcap,f"{stem}-{m}-quarter-ceiling")
                    for row,accepted in enumerate(recognized,start=2):
                        near(cell(path,m,f"M{row}"),accepted/24,f"{stem}-{m}-quarter-accepted-row{row}")
                    total=sum(recognized)
                    near(cell(path,m,"M5"),total/24,f"{stem}-{m}-quarter-accepted-total")
                    near(cell(path,m,"M6"),max(0,qcap-total)/24,f"{stem}-{m}-quarter-left")
                    near(cell(path,m,"M7"),max(0,total-qcap)/24,f"{stem}-{m}-quarter-over")
        # Test TRUE daily-formula recomputation with manually editable exclusion.
        near(cell(path,3,"E2"),1/24,f"{stem}-weekday-default-exclusion")
        near(cell(path,3,"F2"),3/24,f"{stem}-weekday-after-exclusion")
        near(cell(path,3,"E3"),1.5/24,f"{stem}-holiday-manual-exclusion")
        near(cell(path,3,"F3"),3.5/24,f"{stem}-holiday-after-manual-exclusion")
        near(cell(path,3,"E4"),0,f"{stem}-holiday-zero-default")
        near(cell(path,3,"F4"),5/24,f"{stem}-holiday-after-zero-exclusion")
        print("SCENARIO_MODE_PASS",stem,mode,qcap,meal_month_cap,meal_quarter_cap)
    print("ALL_DUAL_MONTHLY_AND_QUARTERLY_SCENARIOS_PASS")

if __name__=="__main__":
    p=argparse.ArgumentParser()
    p.add_argument("--verify",type=Path,required=True)
    run(p.parse_args().verify)
