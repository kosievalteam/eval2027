"""template.html + history.json -> self-contained card HTML. Usage: python3 -I build_card.py template.html history.json out.html"""
import sys, json
tpl, data, out = sys.argv[1:4]
D = json.load(open(data, encoding='utf-8'))
slim = {'years': [str(y) for y in D['years']], 'programs': D['programs'], 'fiscal_meta': D.get('fiscal_meta'), 'fiscal_rows': D.get('fiscal_rows')}   # hwp 원표(raw)는 카드에 불필요
js = json.dumps(slim, ensure_ascii=False, separators=(',', ':')).replace('</', '<\\/')
html = open(tpl, encoding='utf-8').read().replace('__DATA__', js)
open(out, 'w', encoding='utf-8').write(html)
print(out, len(html.encode()) // 1024, 'KB')
