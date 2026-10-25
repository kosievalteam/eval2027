"""열린재정 재정사업평가결과(UOPKOBRA03) 수집 → 연도별 전 사업 결과 JSON

화면 https://www.openfiscaldata.go.kr/op/ko/br/UOPKOBRA03?evlAcntYr=2026 이 내부적으로 호출하는
POST /op/ko/br/selectPgrvRsltList.do (JSON body {"searchDVO":{evlAcntYr, offcCd, pfinBizEvlCd}}) 를 그대로 사용.
인증키 불필요. 소관·평가제도를 비우면 해당 연도 전체 결과가 한 번에 반환됨(2026년 약 2,600행).

사용: python3 -I tools/fetch_openfiscal.py 2019 2026 data/openfiscal_results.json [cache_dir]
  cache_dir 에 pgrv_<year>.json 이 있으면 재요청하지 않음(서버가 간헐적으로 연결을 끊으므로 재시도 포함).

출력 레코드 필드: year, system(평가제도), ministry(소관), ministry_code, account(회계), program(프로그램),
  name(세부사업), result(결과), type(평가유형), note(비고)
"""
import sys, os, json, time, urllib.request

URL = 'https://www.openfiscaldata.go.kr/op/ko/br/selectPgrvRsltList.do'
HEADERS = {
    'Content-Type': 'application/json;charset=UTF-8',
    'X-Requested-With': 'XMLHttpRequest',
    'User-Agent': 'Mozilla/5.0',
}


def fetch_year(year, cache_dir=None, tries=5):
    cache = os.path.join(cache_dir, f'pgrv_{year}.json') if cache_dir else None
    if cache and os.path.exists(cache):
        try:
            return json.load(open(cache, encoding='utf-8'))
        except json.JSONDecodeError:
            pass
    body = json.dumps({'searchDVO': {'evlAcntYr': str(year), 'offcCd': '', 'pfinBizEvlCd': ''}}).encode()
    req = urllib.request.Request(URL, data=body, headers=dict(HEADERS, Referer=f'https://www.openfiscaldata.go.kr/op/ko/br/UOPKOBRA03?evlAcntYr={year}'))
    last = None
    for i in range(tries):
        try:
            with urllib.request.urlopen(req, timeout=180) as r:
                payload = json.loads(r.read().decode('utf-8'))
            if cache:
                os.makedirs(cache_dir, exist_ok=True)
                json.dump(payload, open(cache, 'w', encoding='utf-8'), ensure_ascii=False)
            return payload
        except Exception as e:  # 연결 리셋 등
            last = e
            time.sleep(3 * (i + 1))
    raise SystemExit(f'[{year}] 수집 실패: {last}')


def compact(payload, year):
    out = []
    for r in payload.get('opKoBrPgrvRsltList') or []:
        out.append({
            'year': str(r.get('evlAcntYr') or year), 'system': r.get('evlSstNm'),
            'ministry': r.get('offcNm'), 'ministry_code': r.get('offcCd'), 'account': r.get('acntNm'),
            'program': r.get('pgmNm'), 'name': r.get('evlSayNm'), 'result': r.get('pfinBizEvlCrtrNm'),
            'type': r.get('evlTyNm'), 'note': r.get('bizDivCd'),
        })
    return out


def main(y0, y1, out_path, cache_dir=None):
    rows = []
    for y in range(int(y0), int(y1) + 1):
        part = compact(fetch_year(y, cache_dir), y)
        print(f'{y}: {len(part)} rows')
        rows.extend(part)
    meta = {'source': 'https://www.openfiscaldata.go.kr/op/ko/br/UOPKOBRA03', 'endpoint': URL,
            'fetched_at': time.strftime('%Y-%m-%d'), 'years': [str(y) for y in range(int(y0), int(y1) + 1)]}
    json.dump({'meta': meta, 'rows': rows}, open(out_path, 'w', encoding='utf-8'), ensure_ascii=False, separators=(',', ':'))
    print(f'total {len(rows)} rows -> {out_path}')


if __name__ == '__main__':
    main(*sys.argv[1:5])
