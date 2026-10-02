import csv, json, re
from collections import OrderedDict
H=json.load(open('harvest.json'))
CODE_RX=re.compile(r'\b(?:FFI|FF|FB|FG|FMA|FREF|P|FE)_[A-Z0-9_]+\s*\(')
def ver(code):
    h=H.get(code) or H.get('FE item: '+code)
    if not h: return '', '', ''
    status = 'VERIFIED - returned numbers' if h['num']>0 else ('TESTED - returned no data (avoid)' if h['n']>0 else '')
    ex=h['example'][2] if h['example'] else ''
    samp='; '.join(map(str,h['samples'][:3]))
    if len(CODE_RX.findall(ex))>1 and samp: samp+='  (from a combined formula - may not be this item alone)'
    return status, samp, ex
rows=OrderedDict()
def add(code,family,desc,section,src,syntax):
    k=code
    if k in rows:
        r=rows[k]
        if src not in r['Source']: r['Source']+=' / '+src
        if desc and desc not in r['Description'] and len(r['Description'])<160: r['Description']+=' || '+desc
        return
    rows[k]=dict(Code=code,Family=family,Description=desc,Section=section,Source=src,Syntax=syntax)
TPL={'BKIND':'Banks','SFIND':'Specialty Finance','INSIND':'Insurance'}
for f,sh,c,l in csv.reader(open('/home/claude/codes/catalog.csv')):
    if c.startswith('FFI_'):
        parts=[p.strip() for p in l.split('|')]
        tid=next((p for p in parts if p.startswith('F.') and 'IND.' in p),'')
        tpl=next((v for k,v in TPL.items() if k in tid),'')
        desc=next((p for p in reversed(parts) if p and not p.startswith('F.') and not re.match(r'^[A-Z]{4}\d+',p) and p not in ('+','+/-','-/+','Point-in-time','Period','Bronze','Silver','Gold','As Reported Notes','As Reported Face','Calculated','added','krol')),'')
        if l.startswith(('Balance Sheet |','Income Statement |','Cash Flow |','Regulatory','Asset Quality','Loans and','Average')):
            desc=parts[2] if len(parts)>2 else desc; sec=' | '.join(parts[:2])
        else: sec=sh
        add(c,'FFI (FactSet Fundamentals Industry - bank/SF/insurance detail, global)',desc,sec,f'Industry Codes file ({tpl or "Supp"})',f'FFI_X(ANN_L,2025) / (QTR_L,0) ; add ,,,,USD for $ items')
    elif c.startswith('FF_'):
        add(c,'FF (FactSet Fundamentals - standardized, global)',l,'','FactSet_Fundamentals_Codes_SF file','FF_X(ANN_R,2025) / (QTR,"2025/1F") ; add ,,,,USD for $ items')
for t,sec,c,d,s in csv.reader(open('/home/claude/codes/fb.csv')):
    add(c.replace('()',''),'FB (FactSet regulatory - US bank holding cos. FR Y-9C / call reports)',d,sec,'Regulatory Bank Codes (FFB) file - '+('Holding Companies' if t=='Bank' else 'Banking Institutions'),'FB_X(QTR,"2025/1F") or FB_X(ANN,2025,,,RF[,USD]); items with a type arg e.g. FB_LOAN(\'TOT\') - confirm syntax')
for line in open('fe_items.txt'):
    fn,item,lab,val=[x.strip() for x in line.split('|')]
    add(item,'FE (FactSet Estimates item - goes INSIDE FE_ESTIMATE / FE_TIMESERIES / FE_VALUATION)',lab,fn,'All_Estimates_Formulas (HBAN)',f'{fn}({item},MEAN,...)')
MAN=json.load(open('manual_desc.json'))
FAMN={'FF':'FF (FactSet Fundamentals - standardized, global)','FG':'FG (FactSet global price/reference)','P':'P (FactSet prices)','FREF':'FREF (FactSet reference/market value)','FMA':'FMA (FactSet market aggregates)','FE':'FE (FactSet Estimates function)'}
for k,h in H.items():
    code=k.replace('FE item: ','')
    if code not in rows:
        fam='FE (FactSet Estimates item)' if k.startswith('FE item') else FAMN.get(code.split('_')[0],code.split('_')[0])
        add(code,fam,MAN.get(code,''),'','HBAN / CAMELS client template files','')
for code,d in MAN.items():
    if code in rows and not rows[code]['Description']: rows[code]['Description']=d
with open('/home/claude/kb/FactSet_Code_Catalog.csv','w',newline='') as fh:
    w=csv.writer(fh); w.writerow(['Code','Family','Description','Section / Statement','Source file','Verified in a refreshed client file?','Sample values returned','Working example formula','Syntax pattern'])
    for r in rows.values():
        v=ver(r['Code'])
        w.writerow([r['Code'],r['Family'],r['Description'],r['Section'],r['Source'],v[0],v[1],v[2],r['Syntax']])
print(len(rows))
from collections import Counter
print(Counter(r['Family'][:4] for r in rows.values()))
