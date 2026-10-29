#!/usr/bin/python3
"""Independent LibreOffice cross-version behavioral-equivalence regression.

Compare REAL v3.2 and v3.3 binary workbooks side by side using identical
settings and entry cells. No formula cell is overwritten. Monthly meal ceiling
is normalized to 20 and the new quarterly meal ceiling to 1000 so the new
rules cannot create expected differences in preserved legacy behavior.
"""
import sys,time
from pathlib import Path
import uno
from com.sun.star.beans import PropertyValue

def prop(key,val):
    p=PropertyValue();p.Name=key;p.Value=val;return p

def connect():
    ctx=uno.getComponentContext()
    resolver=ctx.ServiceManager.createInstanceWithContext("com.sun.star.bridge.UnoUrlResolver",ctx)
    for _ in range(100):
        try:return resolver.resolve("uno:socket,host=127.0.0.1,port=20271;urp;StarOffice.ComponentContext")
        except Exception:time.sleep(.2)
    raise RuntimeError("LibreOffice UNO service unavailable")

def setting(sheet,ref,value):
    c=sheet.getCellRangeByName(ref)
    if isinstance(value,str):c.setString(value)
    else:c.setValue(float(value))

def near(a,b,label):
    if abs(a-b)>1e-7:raise AssertionError(f"{label}: v32={a}, v33={b}")
def main():
    folder=Path(sys.argv[1]).resolve()
    ctx=connect()
    desk=ctx.ServiceManager.createInstanceWithContext("com.sun.star.frame.Desktop",ctx)
    scenarios=(
      (2026,8,"예"),
      (2026,4,"예"),
      (2026,8,"아니오"),
      (2028,8,"예"),
    )
    count=0
    for year,day_cap,b6 in scenarios:
        pair=[]
        for version in ("3.2","3.3"):
            path=folder/f"2026초과v{version}.xlsm"
            doc=desk.loadComponentFromURL(uno.systemPathToFileUrl(str(path)),"_blank",0,
                   (prop("Hidden",True),prop("ReadOnly",True),prop("UpdateDocMode",3),
                    prop("MacroExecutionMode",4)))
            if doc is None:raise RuntimeError(f"Cannot open {path}")
            pair.append(doc)
        old,new=pair
        try:
            for idx,doc in enumerate(pair):
                config=doc.Sheets.getByName("설정")
                for cell,val in (("B1",year),("B2",day_cap),("B3",67),("B6",b6),("B7",20),("B8",177)):
                    setting(config,cell,val)
                if idx==1:
                    setting(config,"B9",1000)
                    setting(config,"B10","아니오")
                for m in range(1,13):
                    sheet=doc.Sheets.getByName(str(m)+"월")
                    maxdays=29 if m==2 else 30 if m in (4,6,9,11) else 31
                    for n in range(1,maxdays+1):
                        # February 29 is intentionally blank in the non-leap year.
                        if m==2 and n==29 and year!=2028:continue
                        row=n+1
                        k=(n-1)%12
                        values=(
                          ("평일",6*3600,19*3600),
                          ("평일",7*3600+59*60+59,19*3600),
                          ("평일",8*3600,19*3600+60),
                          ("평일",9*3600,18*3600),
                          ("휴일",10*3600,15*3600),
                          ("휴일",10*3600,11*3600+60),
                          ("휴일",6*3600,23*3600),
                          ("평일",None,None),
                          ("평일",9*3600,23*3600),
                          ("휴일",6*3600,10*3600),
                          ("평일",7*3600,19*3600+60),
                          ("휴일",10*3600,11*3600+59),
                        )[k]
                        typ,start,end=values
                        setting(sheet,f"B{row}",typ)
                        if start is not None:
                            setting(sheet,f"C{row}",start/86400)
                            setting(sheet,f"D{row}",end/86400)
                # Independent spending source refs (I/J in v3.2 -> J/K in v3.3).
                jan=doc.Sheets.getByName("1월")
                ledger=doc.Sheets.getByName("급량")
                vendor=ledger.getCellRangeByName("A2").getString()
                assert vendor
                setting(ledger,"B2",40000)
                setting(jan,"I2" if idx==0 else "J2",vendor)
                setting(jan,"J2" if idx==0 else "K2",5000)
                setting(doc.Sheets.getByName("4월"),"I2" if idx==0 else "J2",vendor)
                setting(doc.Sheets.getByName("4월"),"J2" if idx==0 else "K2",3000)
                doc.enableAutomaticCalculation(True)
                doc.calculateAll();doc.calculateAll()
            matched=0
            for m in range(1,13):
                a=old.Sheets.getByName(str(m)+"월")
                b=new.Sheets.getByName(str(m)+"월")
                daycount=29 if m==2 else 30 if m in (4,6,9,11) else 31
                for row in range(2,daycount+2):
                    for prior,now in (("E","F"),("F","G"),("G","H")):
                        c0=a.getCellRangeByName(prior+str(row))
                        c1=b.getCellRangeByName(now+str(row))
                        near(c0.getValue(),c1.getValue(),f"{year}-daycap{day_cap}-b6{b6}-month{m}-row{row}-{prior}->{now}")
                        matched+=1
                    da=a.getCellRangeByName(f"A{row}").getString()
                    db=b.getCellRangeByName(f"A{row}").getString()
                    if da!=db:raise AssertionError(f"{year} month{m} calendar row{row}: {da!r} != {db!r}")
                summary=daycount+2
                for r in range(summary,summary+3):
                    # Monthly meal outcomes must be identical with B7=20 and
                    # B9=1000, even when v3.3 dynamically applies both ceilings.
                    near(a.getCellRangeByName("G"+str(r)).getValue(),
                         b.getCellRangeByName("H"+str(r)).getValue(),
                         f"{year} month{m} legacy monthly meal summary H{r}")
                    matched+=1
                # Uncapped daily overtime total should remain the same.
                near(a.getCellRangeByName("E"+str(summary)).getValue(),
                     b.getCellRangeByName("F"+str(summary)).getValue(),
                     f"{year} month{m} uncapped total")
                matched+=1
            l0=old.Sheets.getByName("급량");l1=new.Sheets.getByName("급량")
            for ref,expected in (("O2",8000),("P2",32000)):
                for sheet,label in ((l0,"v3.2"),(l1,"v3.3")):
                    near(sheet.getCellRangeByName(ref).getValue(),expected,f"{year}-{label}-ledger-{ref}")
            count+=matched
            print("DIFFERENTIAL_NUMERIC_CASE_PASS",year,day_cap,b6,matched,flush=True)
        finally:
            for doc in pair:doc.close(True)
    desk.terminate()
    print("ALL_V32_V33_PRESERVED_DAILY_MEAL_BEHAVIOR_PASS",count,flush=True)
if __name__=="__main__":main()
