#!/usr/bin/python3
"""Independent LibreOffice UNO scenario tests against the original v3.3 workbook.

Open release workbook read-only in a separate in-memory document, edit user-facing
cells through Calc exactly as a user would, recalculate, export only temp copies.
Never save or alter the release .xlsm on disk.
"""
import sys,time
from pathlib import Path
import uno
from com.sun.star.beans import PropertyValue

def prop(name,value):
    p=PropertyValue();p.Name=name;p.Value=value;return p

def connect():
    ctx=uno.getComponentContext()
    resolver=ctx.ServiceManager.createInstanceWithContext("com.sun.star.bridge.UnoUrlResolver",ctx)
    last=None
    for n in range(70):
        try:return resolver.resolve("uno:socket,host=127.0.0.1,port=20269;urp;StarOffice.ComponentContext")
        except Exception as e:
            last=e
            time.sleep(.2)
    raise RuntimeError("LibreOffice UNO socket unavailable") from last

def set_value(sheet,ref,v):
    cell=sheet.getCellRangeByName(ref)
    if isinstance(v,str):cell.setString(v)
    else:cell.setValue(float(v))

def populate(doc,mode):
    set_value(doc.Sheets.getByName("설정"),"B10",mode)
    for month in range(1,13):
        sh=doc.Sheets.getByName(f"{month}월")
        days=29 if month==2 else 30 if month in (4,6,9,11) else 31
        summary=days+2
        pos=(month-1)%3
        hours=(100,60,15 if ((month-1)//3)%2==0 else 30)[pos]
        # Set daily recognized hours, not just monthly display totals.
        # This keeps independent SUM and monthly-cap formulas consistent.
        for row in range(2,summary):
            if month==3 and row in (2,3,4): continue
            set_value(sh,f"F{row}",0)
        if month==3:
            assert hours==15
            set_value(sh,"B2","평일")
            set_value(sh,"C2",6/24)
            set_value(sh,"D2",19/24)
            set_value(sh,"B3","휴일")
            set_value(sh,"C3",10/24)
            set_value(sh,"D3",15/24)
            set_value(sh,"E3",1.5/24)
            set_value(sh,"B4","휴일")
            set_value(sh,"C4",10/24)
            set_value(sh,"D4",15/24)
            set_value(sh,"F5",3.5/24)
        else:
            row,remaining=2,hours
            while remaining:
                chunk=min(8,remaining)
                set_value(sh,f"F{row}",chunk/24)
                row+=1
                remaining-=chunk
        set_value(sh,f"H{summary}",25)

def main():
    source=Path(sys.argv[1]).resolve()
    output=Path(sys.argv[2]).resolve()
    output.mkdir(parents=True,exist_ok=True)
    input_path=source/"2026초과v3.3.xlsm"
    if not input_path.is_file():raise AssertionError(input_path)
    ctx=connect()
    desktop=ctx.ServiceManager.createInstanceWithContext("com.sun.star.frame.Desktop",ctx)
    for stem,mode in (("nonfield","아니오"),("field","예")):
        output_path=output/(stem+".xlsx")
        doc=desktop.loadComponentFromURL(
            uno.systemPathToFileUrl(str(input_path)),"_blank",0,
            (prop("Hidden",True),prop("ReadOnly",True),prop("UpdateDocMode",3),
             prop("MacroExecutionMode",4)))
        if doc is None:raise RuntimeError("Failed to open "+str(input_path))
        try:
            populate(doc,mode)
            doc.enableAutomaticCalculation(True)
            doc.calculateAll()
            doc.calculateAll()
            for sheet,ref in (("1월","F33"),("2월","F31"),("3월","E2"),("3월","F2"),("3월","E3"),("3월","F3"),("3월","E4"),("3월","F4"),("3월","F33"),("3월","M2"),("3월","M3"),("3월","M17")):
                cell=doc.Sheets.getByName(sheet).getCellRangeByName(ref)
                print("UNO_DIAGNOSTIC",stem,sheet,ref,cell.getValue(),cell.getFormula(),flush=True)
            # Store only a temporary XLSX of the in-memory scenario.
            doc.storeAsURL(
                uno.systemPathToFileUrl(str(output_path)),
                (prop("FilterName","Calc MS Excel 2007 XML"),prop("Overwrite",True)))
            print("UNO_RECALCULATED",stem,str(output_path),output_path.stat().st_size,flush=True)
        finally:doc.close(True)
    desktop.terminate()

if __name__=="__main__":
    main()
