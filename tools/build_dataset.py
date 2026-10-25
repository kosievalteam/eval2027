"""grade.xlsx + tables.json(hwp) [+ openfiscal_results.json] -> data/history.json
Usage: python3 -I build_dataset.py grade.xlsx tables.json out.json [openfiscal_results.json]"""
import sys, json, re, openpyxl
xlsx, tjson, out = sys.argv[1:4]
fiscal_path = sys.argv[4] if len(sys.argv) > 4 else None
YEARS=list(range(2019,2027))
ALIAS={'통상협정국내내책추진지원':'통상협정국내대책추진지원'}
def norm(s):
    s=ALIAS.get((s or '').strip(),s or '')
    s=re.sub(r'\((중기부|산업부|고용부)\)','',s)
    s=re.sub(r'[\s\(\)（）·ㆍ_\-]','',s).replace('R&D','').replace('일반','').replace('정보화','').replace('융자','').replace('기금','')
    return re.sub(r'(지원|개발|구축|강화|육성|조성|촉진|활성화|진출|운영|관리)사업$',r'\1',s)

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
      'findings':{}, 'actions':{}, 'fiscal':{}
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

# ---- 열린재정 재정사업평가결과 결합: 세부사업명(정규화) 일치, 복수 후보면 소관명 일치 우선
MINI_ALIAS={'산업통상자원부':['산업통상자원부','기후에너지환경부','산업통상부'],'중소벤처기업부':['중소벤처기업부'],'환경부':['환경부','기후에너지환경부']}
def mnorm(m): return re.sub(r'\s*\(.*?\)\s*','',m or '').strip()
fiscal_stats={'rows':0,'matched_rows':0,'programs_with_fiscal':0}
if fiscal_path:
    F=json.load(open(fiscal_path,encoding='utf-8'))
    by_name={}
    for p in progs: by_name.setdefault(norm(p['name']),[]).append(p)
    # 과거명 병기 사업명 "현재명_(과거)과거명" 도 색인
    for p in progs:
        m=re.match(r'(.+?)_\(과거\)(.+)',p['name'])
        if m:
            for part in m.groups(): by_name.setdefault(norm(part),[]).append(p)
    for r in F['rows']:
        fiscal_stats['rows']+=1
        cands=by_name.get(norm(r['name'])) or []
        if len(cands)>1:
            same=[p for p in cands if mnorm(r['ministry']) in MINI_ALIAS.get(mnorm(p['ministry']),[mnorm(p['ministry'])])]
            cands=same or cands
        if not cands: continue
        p=cands[0]
        rec={'system':r['system'],'result':r['result'],'type':r['type'],'ministry':r['ministry'],'account':r['account'],'program':r['program'],'src_name':r['name']}
        p['fiscal'].setdefault(r['year'],[]).append(rec); fiscal_stats['matched_rows']+=1
    fiscal_stats['programs_with_fiscal']=sum(1 for p in progs if p['fiscal'])
    print('fiscal:',fiscal_stats)
json.dump({'years':YEARS,'programs':progs,'hwp_findings':findings,'hwp_actions':actions,'fiscal_meta':(json.load(open(fiscal_path,encoding='utf-8'))['meta'] if fiscal_path else None),'fiscal_rows':fiscal_stats['rows']},open(out,'w',encoding='utf-8'),ensure_ascii=False,indent=1)
print('programs',len(progs),'findings rows',len(findings),'actions rows',len(actions))
print('unmatched findings:',u1); print('unmatched actions:',u2)
both=[p for p in progs if p['findings'] and p['actions']]
print('programs with findings+actions:',len(both))
for p in both: print(' ',p['ministry'],p['name'],'| grades',[p['grade'][str(y)] or '-' for y in YEARS],'| F',list(p['findings']),'A',list(p['actions']))
