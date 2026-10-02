"""Harvest every FactSet formula pattern from the refreshed client files with
its cached result, so the guide's 'verified' list is evidence-based."""
import re, json, glob
from collections import defaultdict
from openpyxl import load_workbook
from openpyxl.worksheet.formula import ArrayFormula
U='/path/to/uploads/'
FILES={
 '8/31 refreshed HBAN quarterly template': U+'64a8d193-Quarterly_Template_V3_8-31-26_Edits.xlsx',
 'HBAN master V1 (refreshed)': '/mnt/user-data/uploads/FactSet_MASTER_FILE_-HBAN_V1_Claude_Input.xlsx',
 'HBAN estimates template (refreshed)': U+'39b9df35-All_Estimates_Formulas.xlsx',
 'CAMELS original (refreshed)': U+'CAMELS_FDS_Code_Analysis.xlsx',
}
CODE=re.compile(r'\b((?:FFI|FF|FB|FG|FMA|FREF|P|FE)_[A-Z0-9_]+)\s*\(')
FEITEM=re.compile(r'FE_[A-Z_]+\(\s*([A-Z0-9_]+)\s*,')
out=defaultdict(lambda: {'n':0,'num':0,'na':0,'txt':0,'samples':[],'example':None,'labels':set(),'files':set()})
for tag,p in FILES.items():
    F=load_workbook(p); V=load_workbook(p,data_only=True)
    for sh in F.sheetnames:
        ws,wv=F[sh],V[sh]
        for row in ws.iter_rows():
            for c in row:
                v=c.value; t=v.text if isinstance(v,ArrayFormula) else v
                if not (isinstance(t,str) and t.startswith('=') and 'FDS' in t.upper()): continue
                cv=wv[c.coordinate].value
                keys=set(CODE.findall(t))
                keys|={'FE item: '+m for m in FEITEM.findall(t)}
                for k in keys:
                    d=out[k]; d['n']+=1; d['files'].add(tag)
                    if isinstance(cv,(int,float)):
                        d['num']+=1
                        if len(d['samples'])<4: d['samples'].append(round(cv,4))
                        if d['example'] is None or d['example'][2] is None: d['example']=(tag,sh,t.replace('_xll.',''))
                    elif cv in ('#N/A',None) or (isinstance(cv,str) and cv.startswith('#')): d['na']+=1
                    else: d['txt']+=1
                    if d['example'] is None: d['example']=(tag,sh,t.replace('_xll.',''))
                    # label guess: nearest text to the left on row or header rows
                    for cc in range(c.column-1,0,-1):
                        lv=ws.cell(c.row,cc).value
                        if isinstance(lv,str) and not lv.startswith('=') and len(lv)>3: d['labels'].add(lv[:60]); break
res={k:{**v,'labels':sorted(v['labels'])[:3],'files':sorted(v['files'])} for k,v in out.items()}
json.dump(res,open('harvest.json','w'),indent=1,default=str)
for k in sorted(res): 
    v=res[k]; print(f"{k:38} n={v['n']:4} num={v['num']:4} na={v['na']:4} txt={v['txt']:4} s={v['samples'][:3]}")
