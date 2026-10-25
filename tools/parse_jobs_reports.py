"""재정지원 일자리사업 평가보고서(사업별 HWP) 폴더 -> data/jobs_eval_<year>.json
Usage: python3 -I tools/parse_jobs_reports.py <보고서 폴더> data/jobs_eval_2026.json [2026]

파일: 확장자는 .hwpx이나 대부분 HWP 5.0 OLE 바이너리(일부는 진짜 hwpx zip). 둘 다 지원.
파일명 '(유형) NN (부처약칭) 세부사업(내역사업).hwpx' 에서 부처·세부사업·내역사업을 읽고(가장 안정적),
본문 'Ⅲ. 평가등급 및 개선 권고사항' 절에서 '□ 평가 등급 : X', 종합의견(○), '□ 제도개선 권고사항' 아래 권고(○ 제목 / - 세부)를 읽음.
머리표(부처(부서)·세부사업·내역사업·일자리유형·세부유형·등급)는 등급이 Ⅲ절과 일치할 때만 부서·유형 보강에 사용.
한 파일에 내역사업 여러 개(파일명 NN~MM, NN-MM)면 Ⅲ절이 반복되므로 절마다 레코드 생성.
출력 레코드: category, file, year, system, ministry, dept, name, sub, job_type, sub_type, grade, grade_raw, summary[], items[{title, detail[]}], multi
"""
import sys, os, re, json, zlib, struct, zipfile, collections
import olefile
import xml.etree.ElementTree as ET

MIN = {'노동부': '고용노동부', '고용부': '고용노동부', '복지부': '보건복지부', '문체부': '문화체육관광부', '기후부': '기후에너지환경부', '환경부': '환경부',
       '산림청': '산림청', '성평등부': '성평등가족부', '여가부': '여성가족부', '지재처': '지식재산처', '특허청': '특허청', '해경청': '해양경찰청',
       '과기정통부': '과학기술정보통신부', '과기부': '과학기술정보통신부', '국토부': '국토교통부', '농식품부': '농림축산식품부', '방사청': '방위사업청',
       '산업부': '산업통상자원부', '식약처': '식품의약품안전처', '중기부': '중소벤처기업부', '해수부': '해양수산부', '교육부': '교육부', '국방부': '국방부',
       '법무부': '법무부', '보훈부': '국가보훈부', '통일부': '통일부', '기상청': '기상청', '국가유산청': '국가유산청'}
GRADES = ['개선필요', '개선', '감액', '우수', '양호', '등급제외', '등급 미부여', '등급미부여', '미부여']


# ---------- 텍스트 추출
def hwp5_text(path):
    ole = olefile.OleFileIO(path)
    hdr = ole.openstream('FileHeader').read()
    compressed = struct.unpack('<I', hdr[36:40])[0] & 1
    secs = sorted([s for s in ole.listdir() if s[0] == 'BodyText'], key=lambda s: int(s[1][7:]))
    out = []
    for s in secs:
        data = ole.openstream(s).read()
        if compressed:
            data = zlib.decompress(data, -15)
        i = 0
        while i + 4 <= len(data):
            h = struct.unpack('<I', data[i:i + 4])[0]; tag = h & 0x3ff; size = (h >> 20) & 0xfff; i += 4
            if size == 0xfff:
                size = struct.unpack('<I', data[i:i + 4])[0]; i += 4
            body = data[i:i + size]; i += size
            if tag == 67:
                txt = []; j = 0
                while j + 1 < len(body):
                    c = struct.unpack('<H', body[j:j + 2])[0]
                    if c in (1, 2, 3, 4, 5, 6, 7, 8, 11, 12, 14, 15, 16, 17, 18, 19, 20, 21, 22, 23):
                        j += 16; continue
                    if c == 9: txt.append('\t'); j += 16; continue
                    if c in (10, 13): txt.append('\n'); j += 2; continue
                    if c < 32: j += 2; continue
                    if 0xD800 <= c < 0xDC00 and j + 3 < len(body):
                        c2 = struct.unpack('<H', body[j + 2:j + 4])[0]
                        txt.append(chr(0x10000 + ((c - 0xD800) << 10) + (c2 - 0xDC00))); j += 4; continue
                    txt.append(chr(c)); j += 2
                out.append(''.join(txt))
            elif tag == 66:
                out.append('\n')
    return ''.join(out)


def hwpx_text(path):
    HP = '{http://www.hancom.co.kr/hwpml/2011/paragraph}'
    z = zipfile.ZipFile(path)
    secs = sorted([n for n in z.namelist() if re.match(r'Contents/section\d+\.xml', n)], key=lambda n: int(re.search(r'(\d+)', n).group(1)))
    out = []
    for n in secs:
        root = ET.fromstring(z.read(n))
        for p in root.iter(HP + 'p'):
            txt = ''.join(t.text or '' for t in p.findall('.//' + HP + 't'))
            if txt.strip():
                out.append(txt)
    return '\n'.join(out)


def lines_of(path):
    text = hwpx_text(path) if zipfile.is_zipfile(path) else hwp5_text(path)
    return [l.strip() for l in text.split('\n') if l.strip()]


# ---------- 파일명 → 부처·세부사업·내역사업
def parse_filename(f):
    base = re.sub(r'\.hwpx?$', '', f, flags=re.I)
    m = re.match(r'^\((.+?)\)\s*([\d~\-]+)\s*\((.+?)\)\s*(.+)$', base)
    if not m:
        return None
    cat, nums, minis, rest = m.groups()
    rest = re.sub(r'^평가보고서_', '', rest).strip()
    rest = re.sub(r'\((공개용|\d+종)\)', '', rest).strip()
    name, sub = rest, ''
    if rest.endswith(')'):                       # 마지막 최상위 괄호 = 내역사업
        depth = 0
        for i in range(len(rest) - 1, -1, -1):
            depth += 1 if rest[i] == ')' else -1 if rest[i] == '(' else 0
            if depth == 0:
                inner = rest[i + 1:-1]
                if inner not in ('융자', 'R&D', '정보화', '일반', '지특', '고보'):
                    name, sub = rest[:i].strip(), inner
                break
    minis_full = MIN.get(minis.split(',')[0].strip(), minis)
    return {'category': cat, 'no': nums, 'ministry': minis_full, 'ministry_abbr': minis, 'name': name, 'sub': sub,
            'multi': bool(re.search(r'[~\-]', nums))}


# ---------- 머리표(보강용)
def parse_header(L, year):
    try:
        k = next(i for i, l in enumerate(L) if '평가등급' in l and str(year % 100) in l)
    except StopIteration:
        return []
    end = next((i for i in range(k + 1, len(L)) if re.match(r'^Ⅰ\s*\.', L[i])), len(L))
    cells = L[k + 1:end]
    rows = []; i = 0
    while i < len(cells):
        minis = cells[i]; i += 1
        if i < len(cells) and cells[i].startswith('('):
            minis += ' ' + cells[i]; i += 1
        if i + 5 > len(cells):
            break
        name, sub, jt, st, grade = cells[i:i + 5]; i += 5
        m = re.match(r'(.+?)\s*\((.+)\)$', minis)
        rows.append({'ministry': m.group(1).strip() if m else minis, 'dept': m.group(2).strip() if m else '',
                     'name': name, 'sub': '' if sub == '-' else sub, 'job_type': jt, 'sub_type': st, 'grade': grade.strip()})
    return rows


def norm_grade(g):
    g = (g or '').strip()
    for k in GRADES:
        if g.startswith(k):
            return '등급 미부여' if '미부여' in k else k
    return ''


# ---------- Ⅲ절
def parse_sections(L):
    starts = [i for i, l in enumerate(L) if re.match(r'^Ⅲ\s*\.', l)]
    secs = []
    for n, s in enumerate(starts):
        nxt = starts[n + 1] if n + 1 < len(starts) else len(L)
        e = next((i for i in range(s + 1, nxt) if L[i].startswith('붙임') or re.match(r'^Ⅰ\s*\.', L[i])), nxt)
        body = L[s + 1:e]
        grade_raw = ''; summary = []; items = []; mode = 'summary'; cur = None
        for l in body:
            if l.startswith('□'):
                m = re.search(r'평가\s*등급\s*[:：]\s*(.+)$', l)
                if m: grade_raw = m.group(1).strip(); mode = 'summary'; continue
                if '권고' in l or '개선' in l: mode = 'rec'; continue
                continue
            head = re.match(r'^[○◦ㅇ]\s*(.*)', l)
            if mode == 'summary':
                if head: summary.append(head.group(1).strip())
                elif summary: summary[-1] += ' ' + re.sub(r'^[\-–—·ㆍ*]+\s*', '', l)
            else:
                if head:
                    cur = {'title': head.group(1).strip(), 'detail': []}; items.append(cur)
                elif cur is not None:
                    cur['detail'].append(re.sub(r'^[\-–—·ㆍ*]+\s*', '', l))
        secs.append({'grade_raw': grade_raw, 'grade': norm_grade(grade_raw), 'summary': summary, 'items': items})
    return secs


def main(folder, out, year=2026):
    year = int(year)
    recs = []; problems = []
    for root, _, files in os.walk(folder):
        for f in sorted(files):
            if not f.lower().endswith(('.hwpx', '.hwp')):
                continue
            fn = parse_filename(f)
            if not fn:
                problems.append((f, '파일명 형식 불일치')); continue
            path = os.path.join(root, f)
            try:
                L = lines_of(path)
            except Exception as e:
                problems.append((f, f'읽기 실패 {e}')); continue
            hdr = parse_header(L, year); secs = parse_sections(L)
            if any(s['grade'] for s in secs):
                secs = [s for s in secs if s['grade'] or s['grade_raw']]
            if not secs:
                problems.append((f, 'Ⅲ절 없음')); continue
            for n, sec in enumerate(secs):
                h = hdr[n] if n < len(hdr) and norm_grade(hdr[n]['grade']) and norm_grade(hdr[n]['grade']) == sec['grade'] else {}
                rec = {'category': os.path.basename(root), 'file': f, 'year': str(year), 'system': '일자리사업평가',
                       'ministry': fn['ministry'], 'dept': h.get('dept', ''),
                       'name': h.get('name') or fn['name'], 'sub': h.get('sub') if h else fn['sub'],
                       'job_type': h.get('job_type') or fn['category'], 'sub_type': h.get('sub_type', ''),
                       'grade': sec['grade'], 'grade_raw': sec['grade_raw'], 'summary': sec['summary'], 'items': sec['items'],
                       'multi': len(secs) > 1 or fn['multi'], 'seq': n}
                if len(secs) > 1 and not h:
                    m = re.match(r'^(.+?)\((.+?)\)', sec['grade_raw'])       # '장애인취업알선지원(양호), …' 형식
                    if m and norm_grade(m.group(2)): rec['sub'] = m.group(1).strip(); rec['grade'] = norm_grade(m.group(2))
                if not rec['grade']:
                    problems.append((f, f"등급 해석 불가: {sec['grade_raw'][:40]}"))
                recs.append(rec)
    json.dump({'year': str(year), 'source': '재정지원 일자리사업 평가보고서(사업별 HWP)', 'count': len(recs), 'records': recs},
              open(out, 'w', encoding='utf-8'), ensure_ascii=False, indent=1)
    sys.stdout.reconfigure(encoding='utf-8')
    print(f'{len(recs)} records from {len({r["file"] for r in recs})} files -> {out}')
    print('grades:', dict(collections.Counter(r['grade'] for r in recs)))
    print('권고 없음:', sum(1 for r in recs if not r['items']), '| 머리표 보강:', sum(1 for r in recs if r['dept']))
    for p in problems: print('  !', p)


if __name__ == '__main__':
    main(*sys.argv[1:4])
