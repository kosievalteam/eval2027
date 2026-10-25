# 평가이력카드 프로토타입 (계속평가 사업 평가이력제 신설안)

□ 계속 평가 사업의 연도별 지적사항·개선 여부를 한 장의 이력카드로 정리하여 평가위원에게 제공하는 시안

○ 카드 구성(5개 절): 사업 개요 → 등급 추이(’19~’26) → 전년도 권고 ↔ 금년도 이행실적 대비 → 반복 지적·미개선 사항 → 평가위원 확인 사항
 - 등급 추이는 '중소기업 지원사업 성과평가' 행과 '재정사업 자율평가(기획예산처·열린재정)' 행을 병렬 배치
 - 권고↔이행 대비는 저성과(미흡·개선필요) 연도별로 좌측 평가보고서 원문(지적·권고), 우측 부처 회신(이행조치·완료일)을 대응
 - 반복 지적은 연도 간 지적 제목·내용의 문자 2-gram 유사도(0.28 이상) 또는 공통 핵심어로 자동 탐지한 후보를 제시
 * 샘플: 중소벤처기업부 마케팅지원사업(세부사업 3612300) — ’19·’21 미흡, ’25 개선필요, 미흡횟수 3회

○ 자료 범위와 한계
 - 등급·분야·미흡횟수: 「(’19~’26년) 개선필요사업_등급추이_노란음영수정」 xlsx(99개 사업)
 - 지적·권고 및 이행조치·완료일: 「(2026.10.12) 저성과 사업_개선사항 파악_부처별정리_261023수합본」 hwp(34개 사업·연도, ’23~’25 평가분)
 - 수합본에 중소벤처기업부 회신분이 없어 마케팅지원사업의 지적·이행 란은 공란 상태로 표시(카드가 '미수합' 상태를 명시)
 - 이행조치 미회신 10건(문체부 6, 농림부 3, 금융위 1)은 '이행조치 미회신'으로 표시
 - 열린재정 재정사업 자율평가 등급은 호스트 차단으로 미연동('연동 대기' 표시)

## 파일

| 경로 | 내용 |
|---|---|
| `card/eval_history_card.html` | 완성 카드(데이터 내장, 단일 파일). 상단 선택 상자로 99개 사업 전환, `#p<세부사업ID>`로 직접 링크 |
| `card/template.html` | 카드 서식·렌더링 로직(`__DATA__` 자리에 데이터 삽입) |
| `data/history.json` | xlsx 등급 + hwp 지적·권고·이행조치를 세부사업 단위로 결합한 데이터 |
| `tools/hwp_tables.py` | HWP 5.0 바이너리 표 → 셀(행·열·병합) JSON |
| `tools/build_dataset.py` | xlsx + 표 JSON → `history.json` (사업명 정규화 결합, 지적/권고/이행/완료일 분리) |
| `tools/build_card.py` | 템플릿 + 데이터 → 카드 HTML |
| `tools/fetch_openfiscal.py` | 열린재정 자율평가 결과 수집 골격(인증키·API 주소 확인 후 사용) |

## 재생성

```bash
pip install olefile openpyxl
python3 -I tools/hwp_tables.py <개선사항 수합본.hwp> /tmp/tables.json
python3 -I tools/build_dataset.py <등급추이.xlsx> /tmp/tables.json data/history.json
python3 -I tools/build_card.py card/template.html data/history.json card/eval_history_card.html
```

## 열린재정 연동 절차(안)

□ 모든 사업의 재정사업 자율평가 등급을 카드 2절 두 번째 행에 반영

○ 선행 조건
 - 실행 환경의 네트워크 허용 목록에 `www.openfiscaldata.go.kr` 추가
 - 열린재정 회원가입 후 Open API 인증키 발급, 재정사업평가결과 API의 주소·요청변수·응답필드 확인
○ 결합 방식
 - 1차 키: 열린재정 세부사업코드 ↔ `history.json`의 예산서코드(`budget_code`, 예: 마케팅지원사업 4835-305)
 - 2차 키: 정규화한 세부사업명
 - 결과는 `data/openfiscal_grades.json`에 저장하고 `build_card.py`에서 `programs[].fiscal_grade`로 병합
