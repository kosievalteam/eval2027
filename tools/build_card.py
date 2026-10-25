"""template.html + history.json [+ sample_findings_*.json] -> self-contained card HTML.
Usage: python3 -I build_card.py template.html history.json out.html [data/sample_findings_3612300.json ...] [--only <세부사업ID>]
--only: 해당 사업 하나만 담은 단일 카드(선택 상자 숨김). 예) --only 2030200 (고용노동부 고용유지지원금)
--from: 표시 시작 연도(기본 2020). 그 이전 연도의 등급·분야·지적·이행·열린재정 결과는 카드에서 제외되고 저성과 횟수도 표시 연도 기준으로 계산
sample 파일: 수합본에 원문이 없는 사업의 서식 시연용 예시 문안. 실제 데이터가 없는 연도에만 넣고 항목마다 sample=true 표식."""
import sys, json
args = sys.argv[1:]
only = args.pop(args.index('--only') + 1) if '--only' in args else None
if only: args.remove('--only')
y_from = args.pop(args.index('--from') + 1) if '--from' in args else '2020'
if '--from' in args: args.remove('--from')
tpl, data, out = args[:3]
D = json.load(open(data, encoding='utf-8'))
if only:
    D['programs'] = [p for p in D['programs'] if p['id'] == only]
    assert D['programs'], f'세부사업ID {only} 없음'
D['years'] = [y for y in D['years'] if str(y) >= y_from]
for p in D['programs']:
    p.setdefault('other_findings', {})
    for k in ('grade', 'field', 'findings', 'actions', 'fiscal', 'other_findings'):
        p[k] = {y: v for y, v in p.get(k, {}).items() if str(y) >= y_from}
    p.setdefault('budget', []); p.setdefault('abolished', [])
byid = {p['id']: p for p in D['programs']}
for f in args[3:]:
    S = json.load(open(f, encoding='utf-8'))
    p = byid.get(S.get('program_id'))
    if not p:
        continue
    for y, items in S.get('findings', {}).items():
        if y < y_from: continue
        if not p['findings'].get(y):
            p['findings'][y] = [dict(i, sample=True) for i in items]
    for y, items in S.get('actions', {}).items():
        if y < y_from: continue
        if not p['actions'].get(y):
            p['actions'][y] = [dict(i, sample=True) for i in items]
    p['sample_note'] = S.get('_comment', '')
slim = {'years': [str(y) for y in D['years']], 'programs': D['programs'], 'fiscal_meta': D.get('fiscal_meta'), 'fiscal_rows': D.get('fiscal_rows'), 'coverage': D.get('coverage'), 'jobs_meta': D.get('jobs_meta'), 'single': bool(only)}   # hwp 원표(raw)는 카드에 불필요
js = json.dumps(slim, ensure_ascii=False, separators=(',', ':')).replace('</', '<' + chr(92) + '/')
html = open(tpl, encoding='utf-8').read().replace('__DATA__', js)
open(out, 'w', encoding='utf-8').write(html)
print(out, len(html.encode()) // 1024, 'KB')
