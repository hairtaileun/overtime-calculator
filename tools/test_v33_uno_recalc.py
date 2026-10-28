#!/usr/bin/python3
"""Recalculate isolated XLSM scenarios via LibreOffice UNO calculateAll().

Runs against temporary copies only. Source release workbook is never changed.
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

def main():
    source=Path(sys.argv[1]).resolve()
    output=Path(sys.argv[2]).resolve()
    output.mkdir(parents=True,exist_ok=True)
    ctx=connect()
    desktop=ctx.ServiceManager.createInstanceWithContext("com.sun.star.frame.Desktop",ctx)
    for stem in ("nonfield","field"):
        input_path=source/(stem+".xlsm")
        output_path=output/(stem+".xlsx")
        if not input_path.is_file():raise AssertionError(input_path)
        doc=desktop.loadComponentFromURL(
            uno.systemPathToFileUrl(str(input_path)),"_blank",0,
            (prop("Hidden",True),prop("ReadOnly",False),prop("UpdateDocMode",3),
             prop("MacroExecutionMode",4)))
        if doc is None:raise RuntimeError("Failed to open scenario "+stem)
        try:
            doc.enableAutomaticCalculation(True)
            doc.calculateAll()
            doc.calculateAll()
            for sheet,ref in (("1월","F33"),("2월","F31"),("3월","M2"),("3월","M3"),("3월","M17")):
                cell=doc.Sheets.getByName(sheet).getCellRangeByName(ref)
                print("UNO_DIAGNOSTIC",stem,sheet,ref,cell.getValue(),cell.getFormula(),flush=True)
            doc.storeAsURL(
                uno.systemPathToFileUrl(str(output_path)),
                (prop("FilterName","Calc MS Excel 2007 XML"),prop("Overwrite",True)))
            print("UNO_RECALCULATED",stem,str(output_path),output_path.stat().st_size,flush=True)
        finally:doc.close(True)
    desktop.terminate()

if __name__=="__main__":
    main()
