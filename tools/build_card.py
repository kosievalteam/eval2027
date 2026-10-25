"""template.html + history.json [+ sample_findings_*.json] -> self-contained card HTML.
Usage: python3 -I build_card.py template.html history.json out.html [data/sample_findings_3612300.json ...]
sample 파일: 수합본에 원문이 없는 사업의 서식 시연용 예시 문안. 실제 데이터가 없는 연도에만 넣고 항목마다 sample=true 표식."""
import sys, json
tpl, data, out = sys.argv[1:4]
D = json.load(open(data, encoding='utf-8'))
byid = {p['id']: p for p in D['programs']}
for f in sys.argv[4:]:
    S = json.load(open(f, encoding='utf-8'))
    p = byid.get(S.get('program_id'))
    if not p:
        continue
    for y, items in S.get('findings', {}).items():
        if not p['findings'].get(y):
            p['findings'][y] = [dict(i, sample=True) for i in items]
    for y, items in S.get('actions', {}).items():
        if not p['actions'].get(y):
            p['actions'][y] = [dict(i, sample=True) for i in items]
    p['sample_note'] = S.get('_comment', '')
slim = {'years': [str(y) for y in D['years']], 'programs': D['programs'], 'fiscal_meta': D.get('fiscal_meta'), 'fiscal_rows': D.get('fiscal_rows')}   # hwp 원표(raw)는 카드에 불필요
js = json.dumps(slim, ensure_ascii=False, separators=(',', ':')).replace('</', '<' + chr(92) + '/')
html = open(tpl, encoding='utf-8').read().replace('__DATA__', js)
open(out, 'w', encoding='utf-8').write(html)
print(out, len(html.encode()) // 1024, 'KB')
