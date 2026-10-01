import calendar, hashlib, io, re, zipfile
from pathlib import Path
import xml.etree.ElementTree as ET

SRC=Path('2026초과v3.1.xlsm'); OUT=Path('2026초과v3.2.xlsm')
M='http://schemas.openxmlformats.org/spreadsheetml/2006/main'; R='http://schemas.openxmlformats.org/officeDocument/2006/relationships'

def xparse(b):
    for _,(p,u) in ET.iterparse(io.BytesIO(b),events=('start-ns',)):
        try: ET.register_namespace(p or '',u)
        except ValueError: pass
    return ET.fromstring(b)
def xb(r): return ET.tostring(r,encoding='utf-8',xml_declaration=True)
def cn(a):
    n=0
    for c in re.match(r'[A-Z]+',a).group(): n=n*26+ord(c)-64
    return n
def cell(root,a):
    c=root.find(f".//{{{M}}}c[@r='{a}']")
    if c is not None:return c
    rn=int(re.search(r'\d+$',a).group()); sd=root.find(f'{{{M}}}sheetData'); row=sd.find(f"{{{M}}}row[@r='{rn}']")
    if row is None:
        row=ET.Element(f'{{{M}}}row',{'r':str(rn)}); kids=list(sd); i=next((i for i,e in enumerate(kids) if int(e.attrib['r'])>rn),len(kids)); sd.insert(i,row)
    c=ET.Element(f'{{{M}}}c',{'r':a}); kids=list(row); i=next((i for i,e in enumerate(kids) if cn(e.attrib.get('r','A1'))>cn(a)),len(kids)); row.insert(i,c); return c
def clear(c):
    for x in list(c):c.remove(x)
def inline(root,a,s,style=None):
    c=cell(root,a); clear(c); c.attrib['t']='inlineStr'
    if style:
        sc=cell(root,style)
        if 's' in sc.attrib:c.attrib['s']=sc.attrib['s']
    q=ET.SubElement(c,f'{{{M}}}is'); ET.SubElement(q,f'{{{M}}}t').text=s
def num(root,a,v,style=None):
    c=cell(root,a); clear(c); c.attrib.pop('t',None)
    if style:
        sc=cell(root,style)
        if 's' in sc.attrib:c.attrib['s']=sc.attrib['s']
    ET.SubElement(c,f'{{{M}}}v').text=str(v)
def formula(root,a,s):
    c=cell(root,a); clear(c); c.attrib.pop('t',None); ET.SubElement(c,f'{{{M}}}f').text=s; ET.SubElement(c,f'{{{M}}}v').text='0'
def fmap(root):
    return {c.attrib['r']:f.text for c in root.findall(f'.//{{{M}}}c') if (f:=c.find(f'{{{M}}}f')) is not None and f.text}

def settings(b):
    r=xparse(b); inline(r,'A3','월상한','A2'); num(r,'B3',67,'B2'); inline(r,'A7','월급량상한','A2'); num(r,'B7',20,'B2'); inline(r,'A8','분기상한','A2'); num(r,'B8',177,'B2')
    inline(r,'E10','5. 월급량상한(B7), 분기상한(B8)도 필요에 따라 수정할 수 있습니다.')
    inline(r,'E22','평일 출근시간 7:59까지 급량+1, 퇴근시간 19:01부터 급량+1, 월 급량상한은 설정값(B7) 적용')
    cols=r.find(f'{{{M}}}cols')
    if cols is not None:
        for c in cols:
            if c.attrib.get('min')=='1' and c.attrib.get('max')=='1':c.attrib.update(width='12.5',customWidth='1');break
    return xb(r)
def month(b,m):
    r=xparse(b); fm=fmap(r); days=29 if m==2 else calendar.monthrange(2026,m)[1]; end=days+1; sr,rr,orow=end+1,end+2,end+3
    oh='MIN(TIME(4,0,0),MAX(0,MIN('; nh='MIN(MAX(0,VALUE(설정!$B$2))/24,MAX(0,MIN('
    ow='IF(VALUE(설정!$B$2)=4,4/24,IF(VALUE(설정!$B$2)=8,8/24,9999))'; nw='MAX(0,VALUE(설정!$B$2))/24'
    for row in range(2,end+1):
        e=fm[f'E{row}']; e=e.replace(oh,nh,1).replace(ow,nw,1)
        if nh not in e or nw not in e:raise RuntimeError(f'{m}월 E{row} template')
        formula(r,f'E{row}',e)
        f=fm[f'F{row}'].replace('>=7260','>=3660')
        if '>=3660' not in f:raise RuntimeError(f'{m}월 F{row} template')
        formula(r,f'F{row}',f)
    formula(r,f'E{rr}',f'MAX(0,(MAX(0,VALUE(설정!$B$3))/24)-E{sr})'); formula(r,f'E{orow}',f'MAX(0,E{sr}-(MAX(0,VALUE(설정!$B$3))/24))')
    formula(r,f'G{sr}',f'MIN(SUM(F2:G{end}),MAX(0,VALUE(설정!$B$7)))'); formula(r,f'G{rr}',f'MAX(0,MAX(0,VALUE(설정!$B$7))-G{sr})'); formula(r,f'G{orow}',f'MAX(0,SUM(F2:G{end})-MAX(0,VALUE(설정!$B$7)))')
    if m in (3,6,9,12):
        formula(r,'L1','MAX(0,VALUE(설정!$B$8))'); formula(r,'L6','MAX(0,(MAX(0,VALUE(설정!$B$8))/24)-L5)'); formula(r,'L7','MAX(0,L5-(MAX(0,VALUE(설정!$B$8))/24))')
    return xb(r)
def workbook(b):
    r=xparse(b); c=r.find(f'{{{M}}}calcPr')
    if c is None:c=ET.SubElement(r,f'{{{M}}}calcPr')
    c.attrib.update(calcMode='auto',fullCalcOnLoad='1',forceFullCalc='1'); return xb(r)
def drop_chain(b,ctype=False):
    r=xparse(b)
    for x in list(r):
        if (ctype and x.attrib.get('PartName')=='/xl/calcChain.xml') or (not ctype and x.attrib.get('Type','').endswith('/calcChain')):r.remove(x)
    return xb(r)
def sheetmap(wb,rels):
    ids={x.attrib['Id']:x.attrib['Target'] for x in rels}; d={}
    for s in wb.find(f'{{{M}}}sheets'):
        t=ids[s.attrib[f'{{{R}}}id']]; d[s.attrib['name']]=('xl/'+t) if not t.startswith('/') else t[1:]
    return d

def main():
    with zipfile.ZipFile(SRC) as z:
        vba=z.read('xl/vbaProject.bin'); wb=xparse(z.read('xl/workbook.xml')); rel=xparse(z.read('xl/_rels/workbook.xml.rels')); sm=sheetmap(wb,rel)
        with zipfile.ZipFile(OUT.with_suffix('.tmp.xlsm'),'w') as o:
            for i in z.infolist():
                n=i.filename
                if n=='xl/calcChain.xml':continue
                b=z.read(n)
                if n=='[Content_Types].xml':b=drop_chain(b,True)
                elif n=='xl/_rels/workbook.xml.rels':b=drop_chain(b)
                elif n=='xl/workbook.xml':b=workbook(b)
                elif n==sm['설정']:b=settings(b)
                else:
                    for m in range(1,13):
                        if n==sm[f'{m}월']:b=month(b,m);break
                o.writestr(i,b)
        OUT.with_suffix('.tmp.xlsm').replace(OUT)
    with zipfile.ZipFile(OUT) as z:
        assert z.testzip() is None and hashlib.sha256(z.read('xl/vbaProject.bin')).digest()==hashlib.sha256(vba).digest() and 'xl/calcChain.xml' not in z.namelist()
        wb=xparse(z.read('xl/workbook.xml')); rel=xparse(z.read('xl/_rels/workbook.xml.rels')); sm=sheetmap(wb,rel); sheets=set(sm); fs=[]
        cp=wb.find(f'{{{M}}}calcPr'); assert cp.attrib.get('calcMode')=='auto' and cp.attrib.get('fullCalcOnLoad')=='1' and cp.attrib.get('forceFullCalc')=='1'
        for sn,p in sm.items():
            for a,f in fmap(xparse(z.read(p))).items():fs.append((sn,a,f))
        assert not any(any(t in f for t in ('#REF!','#DIV/0!','#VALUE!','#N/A','#NAME?')) for _,_,f in fs)
        assert not any('[' in f or ']' in f for _,_,f in fs)
        for sn,a,f in fs:
            refs=set(re.findall(r"'([^']+)'!",f)); refs.update(re.findall(r'(?<![\w.])([A-Za-z0-9가-힣_]+)!',f)); assert not refs-sheets,(sn,a,refs-sheets)
        assert sum('>=3660' in f for _,_,f in fs)==366 and sum('7260' in f for _,_,f in fs)==0
        assert sum('MIN(TIME(4,0,0)' in f for _,_,f in fs)==0 and sum('9999' in f for _,_,f in fs)==0
        assert sum('설정!$B$7' in f for _,_,f in fs)==36 and sum('설정!$B$8' in f for _,_,f in fs)==12
        q={'3월':{'L2':"'1월'!$E$33",'L3':"'2월'!$E$31",'L4':"'3월'!$E$33"},'6월':{'L2':"'4월'!$E$32",'L3':"'5월'!$E$33",'L4':"'6월'!$E$32"},'9월':{'L2':"'7월'!$E$33",'L3':"'8월'!$E$33",'L4':"'9월'!$E$32"},'12월':{'L2':"'10월'!$E$33",'L3':"'11월'!$E$32",'L4':"'12월'!$E$33"}}
        for sn,d in q.items():
            fm=fmap(xparse(z.read(sm[sn]))); assert all(fm[a]==f for a,f in d.items())
    print('PASS',hashlib.sha256(OUT.read_bytes()).hexdigest(),hashlib.sha256(vba).hexdigest())
if __name__=='__main__':main()
