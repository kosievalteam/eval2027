"""열린재정(openfiscaldata.go.kr) 재정사업 자율평가 결과 → data/openfiscal_grades.json

상태: 2026-10-25 세션에서는 www.openfiscaldata.go.kr 호스트가 네트워크 정책으로 차단되어 미검증.
      엔드포인트·파라미터명은 열린재정 Open API 안내(회원가입 후 인증키 발급)에서 확인 후 아래 상수를 수정.

사용: OPENFISCAL_KEY=<인증키> python3 -I tools/fetch_openfiscal.py 2019 2026 data/history.json data/openfiscal_grades.json
결합 키: 열린재정의 세부사업코드(또는 소관코드+사업명) ↔ history.json programs[].budget_code / name
"""
import sys, os, json, urllib.request, urllib.parse, re

BASE = os.environ.get('OPENFISCAL_BASE', 'https://www.openfiscaldata.go.kr/op/ko/br/UOPKOBRA03')  # 화면 URL. API URL 확인 후 교체
KEY = os.environ.get('OPENFISCAL_KEY', '')
PARAMS_TEMPLATE = {'evlAcntYr': None, 'Key': None, 'Type': 'json', 'pIndex': 1, 'pSize': 1000}
# 응답 레코드에서 읽을 필드명(추정치, API 명세 확인 후 수정)
FIELDS = {'year': 'evlAcntYr', 'ministry': 'OFFC_NM', 'code': 'PGM_CD', 'name': 'PGM_NM', 'grade': 'EVL_GRD_NM'}


def norm(s):
    return re.sub(r'[\s()（）·ㆍ_\-]', '', s or '').replace('R&D', '').replace('융자', '').replace('정보화', '')


def fetch_year(year):
    q = dict(PARAMS_TEMPLATE, evlAcntYr=year, Key=KEY)
    url = BASE + '?' + urllib.parse.urlencode({k: v for k, v in q.items() if v is not None})
    with urllib.request.urlopen(url, timeout=60) as r:
        body = r.read().decode('utf-8', 'ignore')
    try:
        return json.loads(body)
    except json.JSONDecodeError:
        sys.exit(f'[{year}] JSON 아님(HTML 화면 응답 가능) → BASE를 Open API 주소로 교체 필요. 응답 앞부분: {body[:200]!r}')


def records(payload):
    """응답 구조가 확정되지 않아 dict/list를 재귀 탐색해 FIELDS 키를 모두 가진 레코드만 수집."""
    want = set(FIELDS.values())
    out = []

    def walk(x):
        if isinstance(x, dict):
            if want <= set(x):
                out.append(x)
            for v in x.values():
                walk(v)
        elif isinstance(x, list):
            for v in x:
                walk(v)
    walk(payload)
    return out


def main(y0, y1, hist_path, out_path):
    if not KEY:
        sys.exit('OPENFISCAL_KEY 환경변수(열린재정 인증키) 필요')
    hist = json.load(open(hist_path, encoding='utf-8'))
    by_code = {p['budget_code']: p for p in hist['programs'] if p.get('budget_code')}
    by_name = {norm(p['name']): p for p in hist['programs']}
    result = {}
    for y in range(int(y0), int(y1) + 1):
        for rec in records(fetch_year(y)):
            code, name, grade = rec.get(FIELDS['code']), rec.get(FIELDS['name']), rec.get(FIELDS['grade'])
            p = by_code.get(code) or by_name.get(norm(name))
            if p:
                result.setdefault(p['id'], {})[str(y)] = {'grade': grade, 'src_code': code, 'src_name': name}
    json.dump(result, open(out_path, 'w', encoding='utf-8'), ensure_ascii=False, indent=1)
    print(f'matched programs: {len(result)} -> {out_path}')


if __name__ == '__main__':
    main(*sys.argv[1:5])
