"""grade.xlsx + tables.json(hwp) -> data/history.json  Usage: python3 -I build_dataset.py grade.xlsx tables.json out.json"""
import sys, json, re, openpyxl
xlsx, tjson, out = sys.argv[1:4]
YEARS=list(range(2019,2027))
ALIAS={'통상협정국내내책추진지원':'통상협정국내대책추진지원'}
def norm(s):
    s=ALIAS.get((s or '').strip(),s)
    return re.sub(r'[\s\(\)（）·ㆍ_\-]','',s or '').replace('기금','').replace('R&D','').replace('일반','').replace('정보화','').replace('융자','')

# ---- grades
wb=openpyxl.load_workbook(xlsx,data_only=True); ws=wb.worksheets[0]
hdr=[c.value for c in ws[1]]
progs=[]
for r in ws.iter_rows(min_row=2):
    v={hdr[i]:(c.value.strip() if isinstance(c.value,str) else c.value) for i,c in enumerate(r)}
    if not v.get('세부사업명'): continue
    progs.append({
      'id':str(v['세부사업ID']), 'ministry':v['부처'], 'fund':v['예산/기금'], 'name':v['세부사업명'],
      'start':str(v['사업시작연도'] or ''), 'end':str(v['사업종료연도'] or ''), 'budget_code':v['예산서코드'],
      'field':{str(y):v.get(f'{y%100}년분야') for y in YEARS},
      'grade':{str(y):v.get(f'{y%100}년 등급') for y in YEARS},
      'first_low':v['첫미흡연도'], 'low_count':v['미흡횟수'], 'exclude':v.get('제외사유'),
      'findings':{}, 'actions':{}
    })

# ---- hwp tables -> grid with rowspan fill
T=json.load(open(tjson))
def grid(t):
    g={}
    for c in t['cells']:
        for dr in range(c['rowspan']):
            for dc in range(c['colspan']):
                g.setdefault((c['row']+dr,c['col']+dc), c['text'])
    return [[g.get((r,c),'') for c in range(t['cols'])] for r in range(t['rows'])]

def split_items(txt):
    """'• title\n * (지적) ...\n * (권고) ...' -> list"""
    items=[]; cur=None
    for ln in [l.strip() for l in txt.split('\n') if l.strip()]:
        ln=ln.replace('­','').strip()
        if ln.startswith('•'):
            cur={'title':ln.lstrip('• ').strip(),'issue':'','recommend':'','detail':[]}; items.append(cur)
        elif cur is None:
            cur={'title':ln,'issue':'','recommend':'','detail':[]}; items.append(cur)
        elif '(지적)' in ln: cur['issue']=re.sub(r'^\*?\s*\(지적\)\s*','',ln.lstrip('* ').strip())
        elif '(권고)' in ln: cur['recommend']=re.sub(r'^\*?\s*\(권고\)\s*','',ln.lstrip('* ').strip())
        else: cur['detail'].append(re.sub(r'^[\-\u2010-\u2015\u00ad•*·ㆍ]+\s*','',ln))
    return items

def split_actions(txt):
    items=[]; cur=None
    for ln in [l.strip() for l in txt.split('\n') if l.strip()]:
        ln=ln.replace('­','-').strip()
        if ln.startswith('•'):
            body=ln.lstrip('• ').strip()
            m=re.search(r'[\(（]([’‘\'`]?\d{2}\.\s*\d{1,2}월[^)）]*)[\)）]\s*$',body)
            cur={'action':body,'done':m.group(1) if m else '','detail':[]}; items.append(cur)
        elif cur is None:
            m=re.search(r'[\(（]([’‘\'`]?\d{2}\.\s*\d{1,2}월[^)）]*)[\)）]\s*$',ln)
            cur={'action':ln,'done':m.group(1) if m else '','detail':[]}; items.append(cur)
        else: cur['detail'].append(re.sub(r'^[\-\u2010-\u2015\u00ad•*·ㆍ]+\s*','',ln))
    return items

MIN={'고용부':'고용노동부','과기부':'과학기술정보통신부','금융위':'금융위원회','농림부':'농림축산식품부','농식품부':'농림축산식품부','문체부':'문화체육관광부','산림청':'산림청','산업부':'산업통상자원부','해수부':'해양수산부','환경부':'환경부','기후부':'기후에너지환경부','중기부':'중소벤처기업부'}
findings=[]; actions=[]
for row in grid(T[0])[1:]:
    minis,year,name,dept,body=row
    findings.append({'ministry':minis.strip(),'year':year.strip(),'name':name.replace('\n','').strip(),'dept':dept.strip(),'items':split_items(body)})
for row in grid(T[1])[1:]:
    year,minis,name,body=row
    actions.append({'ministry':minis.strip(),'year':year.strip(),'name':name.replace('\n','').strip(),'items':split_actions(body)})

def attach(lst,key):
    un=[]
    for rec in lst:
        hit=[p for p in progs if norm(p['name'])==norm(rec['name'])] or \
            [p for p in progs if norm(rec['name']) and (norm(rec['name']) in norm(p['name']) or norm(p['name']) in norm(rec['name']))]
        if hit: hit[0][key].setdefault(rec['year'],[]).extend(rec['items']); rec['matched']=hit[0]['id']
        else: un.append(rec['name']); rec['matched']=None
    return un
u1=attach(findings,'findings'); u2=attach(actions,'actions')
json.dump({'years':YEARS,'programs':progs,'hwp_findings':findings,'hwp_actions':actions},open(out,'w',encoding='utf-8'),ensure_ascii=False,indent=1)
print('programs',len(progs),'findings rows',len(findings),'actions rows',len(actions))
print('unmatched findings:',u1); print('unmatched actions:',u2)
both=[p for p in progs if p['findings'] and p['actions']]
print('programs with findings+actions:',len(both))
for p in both: print(' ',p['ministry'],p['name'],'| grades',[p['grade'][str(y)] or '-' for y in YEARS],'| F',list(p['findings']),'A',list(p['actions']))
