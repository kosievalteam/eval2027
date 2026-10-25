"""history.json + openfiscal_results.json -> card/index.html (전체 평가사업 목록)
Usage: python3 -I tools/build_index.py data/history.json data/openfiscal_results.json card/index.html [--from 2020]

① 이력카드 수록 사업: xlsx(개선필요사업 등급추이) 99개. 등급 추이·저성과 횟수·원문 수합·이행 회신·열린재정 ’26 결과, 카드 링크
② 열린재정 수록 중소기업지원사업평가 전체 사업: ’21~’25 중소기업지원사업평가 결과 + ’26 통합재정사업평가 중 평가유형에 '중소기업'이 포함된 사업
"""
import sys, json, re, html, collections
args = sys.argv[1:]
y_from = args.pop(args.index('--from') + 1) if '--from' in args else '2020'
if '--from' in args: args.remove('--from')
hist_path, fiscal_path, out = args[:3]
H = json.load(open(hist_path, encoding='utf-8'))
F = json.load(open(fiscal_path, encoding='utf-8'))
YEARS = [str(y) for y in H['years'] if str(y) >= y_from]
esc = html.escape


def mn(m):                                   # '산업통상자원부 (기후부 이관: 26년)' -> '산업통상자원부'
    return re.sub(r'\s*\(.*?\)\s*', '', m or '').strip()


# ---------- ① 카드 수록 사업
progs = sorted(H['programs'], key=lambda p: (mn(p['ministry']), p['name']))
link_of = {}                                 # (열린재정 소관, 세부사업명) -> 카드 사업 id
for p in progs:
    for y, rs in p.get('fiscal', {}).items():
        for r in rs:
            link_of[(r['ministry'], r['src_name'])] = p['id']

# ---------- ② 열린재정 중소기업지원사업평가 전체
uni = collections.OrderedDict()              # (ministry, name) -> {year: [result(+type)]}
for r in F['rows']:
    if r['system'] == '중소기업지원사업평가' or (r['system'] == '통합재정사업평가' and '중소기업' in (r['type'] or '')):
        k = (r['ministry'], r['name'])
        uni.setdefault(k, {'account': r['account'], 'program': r['program'], 'years': {}})
        uni[k]['years'].setdefault(r['year'], []).append(r['result'] + ('' if r['system'] == '중소기업지원사업평가' else ' · ' + (r['type'] or '')))
uni_rows = sorted(uni.items(), key=lambda kv: (kv[0][0], kv[0][1]))
UYEARS = [y for y in ['2021', '2022', '2023', '2024', '2025', '2026'] if y >= y_from]

LOW = {'미흡', '개선필요', '개선', '감액', '폐지'}
WARN = {'사업개선'}
OK = {'우수', '양호', '정상추진'}


def cls(g):
    g = (g or '').split(' · ')[0]
    return 'g-none' if not g else 'g-cut' if g == '감액' else 'g-bad' if g in LOW else 'g-warn' if g in WARN else 'g-ok' if g in OK else 'g-mid'


def chip(g):
    return f'<span class="chip {cls(g)}">{esc(g)}</span>' if g else '<span class="chip g-none">·</span>'


ministries = sorted({mn(p['ministry']) for p in progs} | {k[0] for k, _ in uni_rows})

rows1 = []
for p in progs:
    lows = [y for y in YEARS if p['grade'].get(y) in ('미흡', '개선필요')]
    fy = sorted(y for y, v in p['findings'].items() if v and not v[0].get('sample') and y >= y_from)
    ay = sorted(y for y, v in p['actions'].items() if v and not v[0].get('sample') and y >= y_from)
    f26 = [r['result'] + (' · ' + r['type'] if r.get('type') and r['type'] != '-' else '') for r in p.get('fiscal', {}).get('2026', []) if r['system'] == '통합재정사업평가']
    rows1.append(f'''<tr data-m="{esc(mn(p['ministry']))}" data-t="{esc(mn(p['ministry']) + ' ' + p['name'] + ' ' + p['id'])}">
<td class="m">{esc(mn(p['ministry']))}</td>
<td class="n"><a href="eval_history_card.html#p{p['id']}">{esc(p['name'])}</a><br><span class="sub">{esc(p['fund'] or '')} · {esc(p['id'])} · {esc(p['budget_code'] or '-')}</span></td>
{''.join(f'<td>{chip(p["grade"].get(y))}</td>' for y in YEARS)}
<td class="c"><b>{len(lows)}</b></td>
<td class="c">{('●' if fy else '') + (' ' + '·'.join(y[2:] for y in fy) if fy else '-')}</td>
<td class="c">{('·'.join(y[2:] for y in ay)) if ay else '-'}</td>
<td>{''.join(chip(x) for x in f26) or '<span class="chip g-none">·</span>'}</td>
<td class="sub">{esc(p.get('exclude') or '')}</td></tr>''')

rows2 = []
for (m, n), v in uni_rows:
    pid = link_of.get((m, n))
    name = f'<a href="eval_history_card.html#p{pid}">{esc(n)}</a> <span class="tag">카드</span>' if pid else esc(n)
    rows2.append(f'''<tr data-m="{esc(m)}" data-t="{esc(m + ' ' + n + ' ' + (v['program'] or ''))}" data-c="{1 if pid else 0}">
<td class="m">{esc(m)}</td>
<td class="n">{name}<br><span class="sub">{esc(v['account'] or '')} · {esc(v['program'] or '')}</span></td>
{''.join('<td>' + (''.join(chip(x) for x in v['years'].get(y, [])) or '<span class="chip g-none">·</span>') + '</td>' for y in UYEARS)}
</tr>''')

page = f'''<!doctype html>
<html lang="ko"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1">
<title>전체 평가사업 목록</title>
<link rel="stylesheet" href="https://fonts.googleapis.com/css2?family=Noto+Serif+KR:wght@700&family=Noto+Sans+KR:wght@400;500;700&family=IBM+Plex+Mono:wght@400;500&display=swap">
<style>
:root{{--paper:#f5f6f8;--sheet:#fff;--ink:#1b2330;--ink-2:#4a5566;--ink-3:#7b8596;--line:#d9dee6;--line-2:#eef1f5;--navy:#24456e;--navy-soft:#e8eef7;
--bad:#b23a3a;--bad-soft:#fbe9e9;--warn:#a86a12;--warn-soft:#fcf1dd;--ok:#2f7a4f;--ok-soft:#e5f3ea;--mid:#5c6b82;--mid-soft:#eceff4;--none:#c3c9d3;
--display:"Noto Serif KR",serif;--body:"Noto Sans KR","Malgun Gothic",sans-serif;--mono:"IBM Plex Mono",ui-monospace,monospace}}
@media (prefers-color-scheme:dark){{:root:not([data-theme="light"]){{--paper:#12161d;--sheet:#1a2029;--ink:#e7ebf1;--ink-2:#b5bdca;--ink-3:#8590a1;--line:#2f3845;--line-2:#242c37;--navy:#8fb0dc;--navy-soft:#213247;--bad:#e07a7a;--bad-soft:#3a2224;--warn:#e0b060;--warn-soft:#3a2f1c;--ok:#7fc49a;--ok-soft:#1f3328;--mid:#aab4c4;--mid-soft:#2a3240;--none:#4a5463;color-scheme:dark}}}}
:root[data-theme="dark"]{{--paper:#12161d;--sheet:#1a2029;--ink:#e7ebf1;--ink-2:#b5bdca;--ink-3:#8590a1;--line:#2f3845;--line-2:#242c37;--navy:#8fb0dc;--navy-soft:#213247;--bad:#e07a7a;--bad-soft:#3a2224;--warn:#e0b060;--warn-soft:#3a2f1c;--ok:#7fc49a;--ok-soft:#1f3328;--mid:#aab4c4;--mid-soft:#2a3240;--none:#4a5463;color-scheme:dark}}
*{{box-sizing:border-box}}
body{{margin:0;background:var(--paper);color:var(--ink);font-family:var(--body);font-size:13.5px;line-height:1.5;padding:24px max(16px,env(safe-area-inset-right)) 48px max(16px,env(safe-area-inset-left))}}
.wrap{{max-width:1280px;margin:0 auto;display:flex;flex-direction:column;gap:16px}}
header{{border-bottom:2px solid var(--navy);padding-bottom:12px}}
.eyebrow{{font-size:12px;letter-spacing:.08em;color:var(--navy);font-weight:500}}
h1{{font-family:var(--display);font-size:24px;margin:4px 0 0}}
h1 small{{display:block;font-family:var(--body);font-size:13px;font-weight:400;color:var(--ink-2);margin-top:4px}}
.ctl{{display:flex;flex-wrap:wrap;gap:10px 16px;align-items:center;color:var(--ink-2);font-size:13px}}
.ctl input,.ctl select{{font:inherit;padding:6px 8px;border:1px solid var(--line);border-radius:4px;background:var(--sheet);color:var(--ink)}}
.ctl input{{flex:1 1 240px;min-width:0}}
.ctl label{{display:flex;gap:6px;align-items:center;white-space:nowrap}}
.ctl .cnt{{margin-left:auto;font-family:var(--mono);font-size:12px;color:var(--ink-3)}}
h2{{font-size:15px;margin:8px 0 0;display:flex;gap:10px;align-items:baseline}}
h2 .n{{font-family:var(--mono);font-size:12px;color:var(--navy)}}
h2 .hint{{font-size:12px;font-weight:400;color:var(--ink-3);margin-left:auto}}
.box{{background:var(--sheet);border:1px solid var(--line);border-radius:6px;overflow:auto}}
table{{border-collapse:collapse;width:100%;min-width:980px}}
th,td{{padding:6px 6px;border-bottom:1px solid var(--line-2);text-align:center;vertical-align:top;font-size:12.5px}}
td:last-child{{max-width:9em}}
th{{position:sticky;top:0;background:var(--sheet);color:var(--ink-2);font-weight:500;font-family:var(--mono);font-size:11.5px;cursor:pointer;white-space:nowrap;z-index:1;box-shadow:0 1px 0 var(--line)}}
th.n,td.n,th.m,td.m{{text-align:left}}
td.m{{white-space:nowrap;color:var(--ink-2)}}
td.n{{min-width:16em}}
td.c{{font-family:var(--mono)}}
.sub{{font-size:11px;color:var(--ink-3)}}
a{{color:var(--navy);text-decoration:none;font-weight:500}} a:hover{{text-decoration:underline}}
.chip{{display:inline-block;min-width:3.6em;padding:2px 6px;border-radius:3px;font-size:11.5px;font-weight:500;line-height:1.3;border:1px solid transparent;margin:1px 0;white-space:nowrap}}
.g-bad{{background:var(--bad-soft);color:var(--bad);border-color:var(--bad)}} .g-ok{{background:var(--ok-soft);color:var(--ok)}} .g-mid{{background:var(--mid-soft);color:var(--mid)}}
.g-warn{{background:var(--warn-soft);color:var(--warn);border-color:var(--warn)}} .g-cut{{background:var(--bad);color:#fff}} .g-none{{color:var(--none);border:1px dashed var(--none);min-width:0}}
.tag{{font-size:10.5px;padding:0 5px;border-radius:3px;background:var(--navy-soft);color:var(--navy);vertical-align:1px}}
.note{{font-size:12px;color:var(--ink-3);line-height:1.6}}
tr.hide{{display:none}}
footer{{border-top:1px solid var(--line);padding-top:10px;font-size:11.5px;color:var(--ink-3);line-height:1.7}}
</style></head><body><div class="wrap">
<header><div class="eyebrow">중소기업 지원사업 성과평가 · 계속평가 사업 평가이력제(안)</div>
<h1>전체 평가사업 목록<small>① 이력카드 수록 사업 {len(progs)}개(’{YEARS[0][2:]}~’{YEARS[-1][2:]} 저성과 이력) · ② 열린재정 수록 중소기업지원사업평가 전체 사업 {len(uni_rows)}개(’21~’26) · 사업명을 누르면 이력카드로 이동</small></h1></header>

<div class="ctl">
<input id="q" type="search" placeholder="부처·사업명·세부사업ID 검색" aria-label="검색">
<select id="m" aria-label="부처"><option value="">전체 부처</option>{''.join(f'<option>{esc(m)}</option>' for m in ministries)}</select>
<label><input type="checkbox" id="onlycard"> ②에서 카드 수록 사업만</label>
<span class="cnt" id="cnt"></span>
</div>

<h2><span class="n">①</span>이력카드 수록 사업<span class="hint">열 머리글을 누르면 정렬 · 등급은 중소기업지원사업평가, ’26 열린재정은 통합재정사업평가 결과</span></h2>
<div class="box"><table id="t1"><thead><tr><th class="m">부처</th><th class="n">세부사업</th>{''.join(f'<th>{y[2:]}</th>' for y in YEARS)}<th>저성과</th><th>지적 원문</th><th>이행 회신</th><th>’26 열린재정</th><th>비고</th></tr></thead>
<tbody>{''.join(rows1)}</tbody></table></div>
<p class="note">저성과 = 미흡(~’21)·개선필요(’22~) 횟수(표시 연도 {YEARS[0]}~{YEARS[-1]} 기준). 지적 원문·이행 회신은 수합본(’26.10.23)에 실린 평가연도(● 표시, 예시 문안 제외). ’26 열린재정은 통합재정사업평가 결과와 평가유형.</p>

<h2><span class="n">②</span>열린재정 수록 중소기업지원사업평가 전체 사업<span class="hint">’21~’25 중소기업지원사업평가 결과 · ’26 통합재정사업평가 중 평가유형에 중소기업이 포함된 사업</span></h2>
<div class="box"><table id="t2"><thead><tr><th class="m">소관</th><th class="n">세부사업</th>{''.join(f'<th>{y[2:]}</th>' for y in UYEARS)}</tr></thead>
<tbody>{''.join(rows2)}</tbody></table></div>
<p class="note">열린재정 재정사업평가결과(openfiscaldata.go.kr, {esc(F['meta']['fetched_at'])} 수집) 중 중소기업지원사업평가 {len([r for r in F['rows'] if r['system']=='중소기업지원사업평가'])}행과 ’26년 통합재정사업평가 중소기업 유형 {len([r for r in F['rows'] if r['system']=='통합재정사업평가' and '중소기업' in (r['type'] or '')])}행을 소관·세부사업명으로 묶음. '개선'은 '개선필요'와 같은 등급. 한 연도에 복수 표기는 내역사업 또는 사업명 변경 전후 수록분. '카드' 표식은 ①의 사업과 결합된 행.</p>

<footer><b>자료</b> ① (’19~’26년) 개선필요사업 등급추이 xlsx, (2026.10.12) 저성과 사업 개선사항 파악 수합본 hwp ② 열린재정 재정사업평가결과 API(POST /op/ko/br/selectPgrvRsltList.do). 생성 tools/build_index.py</footer>
</div>
<script>
const q=document.getElementById('q'), m=document.getElementById('m'), oc=document.getElementById('onlycard'), cnt=document.getElementById('cnt');
function norm(s){{return (s||'').toLowerCase().replace(/\\s/g,'');}}
function apply(){{
  const t=norm(q.value), mm=m.value; let n1=0,n2=0;
  document.querySelectorAll('#t1 tbody tr').forEach(r=>{{const ok=(!mm||r.dataset.m===mm)&&(!t||norm(r.dataset.t).includes(t)); r.classList.toggle('hide',!ok); if(ok)n1++;}});
  document.querySelectorAll('#t2 tbody tr').forEach(r=>{{const ok=(!mm||r.dataset.m===mm)&&(!t||norm(r.dataset.t).includes(t))&&(!oc.checked||r.dataset.c==='1'); r.classList.toggle('hide',!ok); if(ok)n2++;}});
  cnt.textContent=`① ${{n1}}개 · ② ${{n2}}개`;
}}
[q,m,oc].forEach(e=>e.addEventListener('input',apply)); apply();
// 열 정렬(문자 기준, 칩은 텍스트)
document.querySelectorAll('th').forEach((th,i)=>th.addEventListener('click',()=>{{
  const tb=th.closest('table').tBodies[0], rows=[...tb.rows], idx=[...th.parentNode.children].indexOf(th), asc=th.dataset.asc!=='1';
  th.parentNode.querySelectorAll('th').forEach(h=>delete h.dataset.asc); th.dataset.asc=asc?'1':'0';
  rows.sort((a,b)=>a.cells[idx].textContent.trim().localeCompare(b.cells[idx].textContent.trim(),'ko',{{numeric:true}})*(asc?1:-1));
  rows.forEach(r=>tb.appendChild(r));
}}));
</script></body></html>'''
open(out, 'w', encoding='utf-8').write(page)
print(out, len(page.encode()) // 1024, 'KB', '| 카드 수록', len(progs), '| 열린재정 중소기업 전체', len(uni_rows), '| 그중 카드 결합', sum(1 for (m, n), _ in uni_rows if (m, n) in link_of))
