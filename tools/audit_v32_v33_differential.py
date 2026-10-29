#!/usr/bin/env python3
"""Independent exhaustive differential audit: the released v3.2 XLSM vs v3.3.

Fail closed on dropped original cells, unplanned changes to formula semantics,
styles, worksheet feature counts, non-target binary parts or Excel package.
Workbooks are read only. Tests inspect actual binary XML, not regenerated code.
"""
from __future__ import annotations
import collections,hashlib,re,sys,zipfile
from pathlib import Path
from xml.etree import ElementTree as ET

sys.path.insert(0,str(Path(__file__).resolve().parent))
from build_v33 import shift_col,shift_ref_text
N={"x":"http://schemas.openxmlformats.org/spreadsheetml/2006/main"}
X="{"+N["x"]+"}"
OLD=Path("2026초과v3.2.xlsm");NEW=Path("2026초과v3.3.xlsm")
EXPECTED={
 "old":"e509d16a4dc832e26dd2f143f80b7ebc13dd3abd7c4b25e20df69fe05b0c3d80",
 "new":"1a5f1cc9cd28d68c3e8151c838b97a62bd37396ad77c2066f730cdc7b6d8d716"}
failures=[]
stats=collections.Counter()
def require(cond,msg):
    stats["assertions"]+=1
    if not cond:
        failures.append(msg)
        if len(failures)<=100:print("DIFF_FAILURE",msg,flush=True)
def getcells(r):
    return {c.get("r"):c for c in r.findall(".//x:sheetData/x:row/x:c",N)}
def formula(c):
    if c is None:return None
    f=c.find("x:f",N)
    return f.text if f is not None else None
def value(c):
    if c is None:return None
    v=c.find("x:v",N)
    return v.text if v is not None else None
def sha(b):return hashlib.sha256(b).hexdigest()
def parsed(z,p):return ET.fromstring(z.read(p))
def colshift(ref):
    m=re.fullmatch(r"([A-Z]+)(\d+)",ref)
    return shift_col(m.group(1))+m.group(2) if m else ref
def manual_formula(month,col,row,days):
    summary=days+2
    return (col=="F" and 2<=row<=days+1
       or col=="F" and row in (summary+1,summary+2)
       or col=="H" and row in (summary,summary+1,summary+2)
       or col=="M" and month in (3,6,9,12) and 1<=row<=7)
def inspect_month(z0,z1,m):
    path=f"xl/worksheets/sheet{m+2}.xml"
    x0=parsed(z0,path);x1=parsed(z1,path)
    a=getcells(x0);b=getcells(x1)
    days=29 if m==2 else (30 if m in (4,6,9,11) else 31)
    mapped={colshift(k):k for k in a}
    lost=[(old,colshift(old)) for old in a if colshift(old) not in b]
    require(not lost,f"MONTH{m} dropped source cells {lost[:15]}")
    new_only=sorted(set(b)-set(mapped),key=lambda x:(int(re.search(r'\d+',x).group()),x))
    expected_new={f"E{row}" for row in range(1,days+2)}
    expected_new|={f"L{row}" for row in (16,17,18)}
    expected_new|={f"M{row}" for row in (16,17,18)}
    if m in (3,6,9,12):
        expected_new|={f"L{row}" for row in range(9,16)}
        expected_new|={f"M{row}" for row in range(9,16)}
    surprises=sorted(set(new_only)-expected_new)
    require(not surprises,f"MONTH{m} unexpected inserted cells {surprises[:20]}")
    stats["old_cells"]+=len(a);stats["mapped_cells"]+=len(mapped)
    stats["new_cells"]+=len(new_only)
    unexpected=[]
    for old,source in a.items():
        target=b.get(colshift(old))
        if target is None:continue
        newref=colshift(old)
        match=re.fullmatch(r"([A-Z]+)(\d+)",newref)
        col,row=match.group(1),int(match.group(2))
        # The new L/M summaries intentionally occupy EMPTY original style-only
        # placeholders. Verify original value/formula/children were all blank;
        # changed styles are legitimate only in these explicit positions.
        overlay=(col in ("L","M") and
                 (row in (17,18) or
                  (m in (3,6,9,12) and 9<=row<=16)))
        if overlay:
            content=[e for e in source if e.tag not in (X+"f",X+"v")]
            if formula(source) is not None or value(source) is not None or content:
                unexpected.append((old,newref,"NONBLANK OVERWRITTEN BY SUMMARY",
                                   formula(source),value(source)))
            else:
                stats["verified_blank_overlays"]+=1
                continue
        # Structure and style: each old cell must keep its original type,
        # style, and metadata in its corresponding shifted location.
        p0=dict(source.attrib);p1=dict(target.attrib)
        p0.pop("r",None);p1.pop("r",None)
        if p0!=p1:unexpected.append((old,newref,"cell attrs",p0,p1))
        f0=formula(source);f1=formula(target)
        expected=shift_ref_text(f0,True) if f0 is not None else None
        if f1!=expected:
            if manual_formula(m,col,row,days):
                stats["intentional_formula_changes"]+=1
            else:
                unexpected.append((old,newref,"formula",expected,f1))
        v0=value(source);v1=value(target)
        if v0!=v1:
            if manual_formula(m,col,row,days):
                stats["expected_cache_clears"]+=1
            else:unexpected.append((old,newref,"cached value",v0,v1))
        # Every other child element except a formula/cache is not allowed to
        # change; this includes inline strings and legacy shared strings.
        children0=[ET.tostring(el,encoding="unicode") for el in source if el.tag not in (X+"f",X+"v")]
        children1=[ET.tostring(el,encoding="unicode") for el in target if el.tag not in (X+"f",X+"v")]
        if children0!=children1:unexpected.append((old,newref,"child contents",children0[:2],children1[:2]))
    require(not unexpected,f"MONTH{m} unexpected original-cell differences {unexpected[:12]}")
    if unexpected:
        stats["unplanned_cell_diffs"]+=len(unexpected)
    print("MONTH_DIFF",m,"old",len(a),"new",len(b),"new_only",new_only[:35],
          "intentional_f_changes",stats["intentional_formula_changes"],
          "unexpected",len(unexpected),flush=True)
    # Compare data validation count and preservation of opaque x14 XML.
    for tag in ("dataValidation","mergeCell","hyperlink","drawing","legacyDrawing","autoFilter","tablePart","ext"):
        n0=sum(x.tag.rsplit("}",1)[-1]==tag for x in x0.iter())
        n1=sum(x.tag.rsplit("}",1)[-1]==tag for x in x1.iter())
        require(n0==n1,f"MONTH{m} feature {tag} count {n0}->{n1}")
    old_dv=[(x.get("type"),x.get("sqref")) for x in x0.iter() if x.tag.rsplit("}",1)[-1]=="dataValidation"]
    new_dv=[(x.get("type"),x.get("sqref")) for x in x1.iter() if x.tag.rsplit("}",1)[-1]=="dataValidation"]
    require(len(old_dv)==len(new_dv),f"MONTH{m} legacy validation count")
    old_x14=[x.text for x in x0.iter() if x.tag.rsplit("}",1)[-1]=="sqref" and x.text]
    new_x14=[x.text for x in x1.iter() if x.tag.rsplit("}",1)[-1]=="sqref" and x.text]
    require(len(old_x14)==len(new_x14),f"MONTH{m} extended validation target count")
    for a0,b0 in zip(old_x14,new_x14):
        require(shift_ref_text(a0,True)==b0,f"MONTH{m} x14 target {a0}->{b0}")
    # Conditional formats: original rules retained, only gray-state rule added.
    n0=sum(x.tag.rsplit("}",1)[-1]=="cfRule" for x in x0.iter())
    n1=sum(x.tag.rsplit("}",1)[-1]=="cfRule" for x in x1.iter())
    require(n1==n0+(m in (3,6,9,12)),f"MONTH{m} CF rule count {n0}->{n1}")
    # Preserve every original row with old height/hidden/style settings.
    rows0={int(x.get("r")):x for x in x0.findall(".//x:sheetData/x:row",N)}
    rows1={int(x.get("r")):x for x in x1.findall(".//x:sheetData/x:row",N)}
    require(set(rows0)==set(rows1),f"MONTH{m} original rows changed")
    for rid,r0 in rows0.items():
        require(r0.attrib==rows1[rid].attrib,f"MONTH{m} row {rid} formatting")
    # Important problem detection: renumbered F/H should never lose existing
    # source cell named range or change user input C/D.
    for row in range(2,days+2):
        for col in "ABCD":
            ref=f"{col}{row}"
            require((formula(a.get(ref)),value(a.get(ref))) ==
                    (formula(b.get(ref)),value(b.get(ref))),
                    f"MONTH{m} original input/date {ref} changed")

def main():
    assert sha(OLD.read_bytes())==EXPECTED["old"]
    assert sha(NEW.read_bytes())==EXPECTED["new"]
    with zipfile.ZipFile(OLD) as z0,zipfile.ZipFile(NEW) as z1:
        n0=set(z0.namelist());n1=set(z1.namelist())
        require(n0==n1,f"ZIP parts changed added={sorted(n1-n0)} missing={sorted(n0-n1)}")
        changes=[p for p in sorted(n0&n1) if z0.read(p)!=z1.read(p)]
        allow={"xl/workbook.xml","xl/sharedStrings.xml","xl/styles.xml","xl/worksheets/sheet1.xml","xl/worksheets/sheet2.xml"}
        allow|={f"xl/worksheets/sheet{i}.xml" for i in range(3,15)}
        unexpected=set(changes)-allow
        require(not unexpected,f"Unexpected mutated OOXML binary components {sorted(unexpected)}")
        print("OOXML_PART_DIFF",len(changes),"changed",changes)
        for p in ("xl/vbaProject.bin","[Content_Types].xml",
                  "xl/_rels/workbook.xml.rels","xl/worksheets/_rels/sheet1.xml.rels"):
            if p in n0:
                require(z0.read(p)==z1.read(p),f"Protected binary/relationship modified: {p}")
        for m in range(1,13):inspect_month(z0,z1,m)
        # Settings and meal ledger: only intended setting cells and account
        # source formula reference shifts permitted.
        for sheet in (1,2):
            a=getcells(parsed(z0,f"xl/worksheets/sheet{sheet}.xml"))
            b=getcells(parsed(z1,f"xl/worksheets/sheet{sheet}.xml"))
            require(set(a)<=set(b),f"Sheet{sheet} dropped original cells")
            new=set(b)-set(a)
            require(new<={ "A9","B9","A10","B10" } if sheet==1 else not new,
                    f"Sheet{sheet} unexpected added cells {sorted(new)[:20]}")
            diff=[]
            for ref,c in a.items():
                d=b[ref]
                if formula(c)!=formula(d):
                    if sheet==2 and shift_ref_text(formula(c),False)==formula(d):
                        stats["ledger_correct_reference_shifts"]+=1
                    else:diff.append((ref,"formula",formula(c),formula(d)))
                if value(c)!=value(d):
                    if sheet==1 and ref=="B7" and value(d)=="25":pass
                    else:diff.append((ref,"value",value(c),value(d)))
                if c.attrib!=d.attrib:diff.append((ref,"attrs",c.attrib,d.attrib))
            require(not diff,f"sheet{sheet} old-cell unexpected changes {diff[:15]}")
    print("DIFF_STATS",dict(stats))
    print("DIFF_RESULT", "PASS" if not failures else "FAIL",len(failures))
    if failures:raise AssertionError(str(len(failures))+" differential errors")
if __name__=="__main__":main()
