#!/usr/bin/python3
"""Independent read-only boundary tests for the exact released macro workbook.

Unlike conventional cap regressions, this never injects F/G/H formulas or
recognition totals: it populates only user-input B,C,D,E and setting cells.
"""
import sys,time
from pathlib import Path
import uno
from com.sun.star.beans import PropertyValue

def prop(k,v):
    p=PropertyValue();p.Name=k;p.Value=v;return p
def main():
    root=Path(sys.argv[1]).resolve()
    context=uno.getComponentContext()
    resolver=context.ServiceManager.createInstanceWithContext("com.sun.star.bridge.UnoUrlResolver",context)
    session=None
    for _ in range(80):
        try:
            session=resolver.resolve("uno:socket,host=127.0.0.1,port=20270;urp;StarOffice.ComponentContext")
            break
        except Exception:time.sleep(.2)
    if session is None:raise RuntimeError("UNO not reachable")
    desktop=session.ServiceManager.createInstanceWithContext("com.sun.star.frame.Desktop",session)
    def load():
        doc=desktop.loadComponentFromURL(
          uno.systemPathToFileUrl(str(root/"2026초과v3.3.xlsm")),"_blank",0,
          (prop("Hidden",True),prop("ReadOnly",True),prop("UpdateDocMode",3),prop("MacroExecutionMode",4)))
        assert doc is not None
        doc.enableAutomaticCalculation(True)
        return doc
    def setcell(sheet,ref,value):
        c=sheet.getCellRangeByName(ref)
        if isinstance(value,str):c.setString(value)
        else:c.setValue(float(value))
    def get(sheet,ref):return sheet.getCellRangeByName(ref).getValue()
    def near(actual,expected,label,tol=2e-7):
        assert abs(actual-expected)<=tol, f"{label}: expected {expected!r}, actual {actual!r}"
        print("BOUNDARY_PASS",label,round(actual,10),flush=True)
    def scenario(name,mode,holiday,start,end,expected_f,expected_g,expected_h,exclusion=None,day_cap=8,b6="예"):
        doc=load()
        try:
            config=doc.Sheets.getByName("설정")
            setcell(config,"B10",mode);setcell(config,"B6",b6);setcell(config,"B2",day_cap)
            sh=doc.Sheets.getByName("3월")
            setcell(sh,"B2","휴일" if holiday else "평일")
            setcell(sh,"C2",start/86400);setcell(sh,"D2",end/86400)
            if exclusion is not None:setcell(sh,"E2",exclusion/86400)
            doc.calculateAll();doc.calculateAll()
            expected_e=(exclusion if exclusion is not None else (0 if holiday else 3600))/86400
            near(get(sh,"E2"),expected_e,name+"-exclusion")
            near(get(sh,"F2"),expected_f/86400,name+"-hours")
            near(get(sh,"G2"),expected_g,name+"-am-or-holiday-meal")
            near(get(sh,"H2"),expected_h,name+"-pm-meal")
        finally:doc.close(True)
    scenarios=[
       ("weekday-early-075959","아니오",False,7*3600+59*60+59,18*3600,1,1,0,None,8,"예"),
       ("weekday-boundary-080000","아니오",False,8*3600,18*3600,0,0,0,None,8,"예"),
       ("evening-before-1901","아니오",False,9*3600,19*3600+59,59,0,0,None,8,"예"),
       ("evening-at-1901","아니오",False,9*3600,19*3600+60,60,0,1,None,8,"예"),
       ("holiday-just-below-0101","아니오",True,10*3600,11*3600+59,3659,0,0,None,8,"예"),
       ("holiday-exact-0101","아니오",True,10*3600,11*3600+60,3660,1,0,None,8,"예"),
       ("holiday-manual-90min","아니오",True,10*3600,15*3600,12600,1,0,5400,8,"예"),
       ("weekday-manual-90min","아니오",False,6*3600,19*3600,9000,1,0,5400,8,"예"),
       ("holiday-b6-no","아니오",True,10*3600,15*3600,0,0,0,None,8,"아니오"),
       ("holiday-daily-cap-4","아니오",True,6*3600,23*3600,14400,1,0,None,4,"예"),
       ("holiday-daily-cap-8","예",True,6*3600,23*3600,28800,1,0,None,8,"예"),
       ("weekday-field-exclusion","예",False,6*3600,19*3600,10800,1,0,None,8,"예"),
    ]
    for case in scenarios:scenario(*case)
    # Integrated recognition: all monthly meals must arise from the *actual*
    # daily attendance formulas, never from injected G/H/H-summary values.
    for mode in ("예","아니오"):
        doc=load()
        try:
            setting=doc.Sheets.getByName("설정")
            for ref,val in (("B10",mode),("B6","예"),("B7",25),("B9",60)):
                setcell(setting,ref,val)
            for month in range(1,7):
                sh=doc.Sheets.getByName(str(month)+"월")
                for row in range(2,15):
                    setcell(sh,f"B{row}","평일")
                    setcell(sh,f"C{row}",7/24)
                    setcell(sh,f"D{row}",20/24)
            doc.calculateAll();doc.calculateAll()
            for m in (1,2,3,4,5,6):
                sh=doc.Sheets.getByName(str(m)+"월")
                summary=(29 if m==2 else 30 if m in (4,6,9,11) else 31)+2
                near(get(sh,"G2"),1,f"{mode}-month{m}-actual-morning-meal")
                near(get(sh,"H2"),1,f"{mode}-month{m}-actual-evening-meal")
                near(get(sh,"F2"),3/24,f"{mode}-month{m}-actual-overtime")
                expected=10 if m in (3,6) else 25
                near(get(sh,f"H{summary}"),expected,f"{mode}-month{m}-meal-recognized")
                near(get(sh,f"H{summary+2}"),26-expected,f"{mode}-month{m}-meal-unrecognized")
            for q in (3,6):
                sh=doc.Sheets.getByName(str(q)+"월")
                for ref,exp in (("M10",25),("M11",25),("M12",10),("M13",60),("M14",0),("M15",15)):
                    near(get(sh,ref),exp,f"{mode}-quarter-{q}-{ref}")
        finally:doc.close(True)
    # Fully independent overtime cap audit: user enters clock times for
    # 25 days, which must generate 75h per month through unmodified F rows.
    # Both staff types have monthly 67h; only nonfield enforces quarter 177h.
    for mode in ("예","아니오"):
        doc=load()
        try:
            setting=doc.Sheets.getByName("설정")
            for ref,val in (("B10",mode),("B2",8),("B3",67),("B8",177),
                            ("B6","예"),("B7",25),("B9",60)):
                setcell(setting,ref,val)
            for m in range(1,7):
                sh=doc.Sheets.getByName(str(m)+"월")
                for row in range(2,27):
                    setcell(sh,f"B{row}","평일")
                    setcell(sh,f"C{row}",6/24)
                    setcell(sh,f"D{row}",19/24)
            doc.calculateAll();doc.calculateAll()
            for m in range(1,7):
                sh=doc.Sheets.getByName(str(m)+"월")
                summary=(29 if m==2 else 30 if m in (4,6,9,11) else 31)+2
                pos=(m-1)%3
                cap=67 if mode=="예" or pos<2 else 43
                near(get(sh,"F2"),3/24,f"{mode}-actual-fullchain-month{m}-daily-3h")
                near(get(sh,f"F{summary}"),75/24,f"{mode}-actual-fullchain-month{m}-raw-75h")
                near(get(sh,"M17"),cap/24,f"{mode}-actual-fullchain-month{m}-cap")
                near(get(sh,"M18"),cap/24,f"{mode}-actual-fullchain-month{m}-accepted")
                near(get(sh,f"F{summary+2}"),(75-cap)/24,f"{mode}-actual-fullchain-month{m}-excess")
                near(get(sh,f"H{summary}"),10 if pos==2 else 25,
                     f"{mode}-actual-fullchain-month{m}-meal-accepted")
            if mode=="아니오":
                for qm in (3,6):
                    sh=doc.Sheets.getByName(str(qm)+"월")
                    near(get(sh,"M2"),67/24,f"{mode}-quarter{qm}-overtime-month1")
                    near(get(sh,"M3"),67/24,f"{mode}-quarter{qm}-overtime-month2")
                    near(get(sh,"M4"),43/24,f"{mode}-quarter{qm}-overtime-month3")
                    near(get(sh,"M5"),177/24,f"{mode}-quarter{qm}-overtime-recognized-177h")
            else:
                for qm in (3,6):
                    sh=doc.Sheets.getByName(str(qm)+"월")
                    assert sh.getCellRangeByName("M5").getString() in ("",), "field quarter hours should be hidden"
                    print("BOUNDARY_PASS",mode,"quarter",qm,"hours panel hidden",flush=True)
        finally:doc.close(True)
    desktop.terminate()
    print("ALL_INDEPENDENT_DAILY_AND_ATTENDANCE_END_TO_END_TESTS_PASSED",flush=True)

if __name__=="__main__":main()
