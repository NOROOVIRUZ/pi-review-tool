# PI / CI / PL 검토 툴 v2

PI PDF를 기준으로 CI(상업송장), PL(패킹리스트), 박스내용 엑셀을 동시에 비교하여
수정요청 목록을 포함한 검토 엑셀을 생성하는 도구입니다.

## 비교 구조

```
PI PDF (기준)
  ├── vs 박스내용.xlsx  : OA 내부 입고 계획 (PI번호로 연결)
  ├── vs CI 시트       : 공급사 상업송장
  └── vs PL 시트       : 공급사 패킹리스트
```

PI가 기준이며, 나머지 세 문서는 PI 기준으로 일치 여부를 검증합니다.

## 사용

```bash
# 기본 실행 (기본값: ~/Downloads/물류.zip → ~/Downloads/new/)
python3 pi_review_tool.py

# ZIP 또는 폴더 지정
python3 pi_review_tool.py /path/to/물류.zip

# 출력 폴더 지정
python3 pi_review_tool.py /path/to/input.zip -o /path/to/new

# 디버그 JSON 함께 저장
python3 pi_review_tool.py /path/to/input.zip --debug-json
```

## 입력 구조

ZIP 또는 폴더 안에 다음 파일이 있으면 자동 인식합니다.

| 파일 | 인식 기준 | 설명 |
|------|-----------|------|
| `*박스내용*.xlsx` | 파일명 포함 | OA 내부 입고 계획 |
| `*.xls` / `*.xlsx` (CI/PL 시트 보유) | 시트명 "CI", "PL" | 공급사 상업송장 + 패킹리스트 |
| `*.pdf` | 경로 내 모든 PDF | PI 계약서 |

## 결과 시트 (10개)

| 시트명 | 내용 |
|--------|------|
| `요약` | 전체 건수와 상태별 집계 |
| `수정요청` | 3개 비교에서 문제 항목만 통합 |
| `PI_vs_박스내용` | PI ↔ 박스내용 전체 비교 |
| `PI_vs_CI` | PI ↔ 상업송장 비교 (수량·단가·금액) |
| `PI_vs_PL` | PI ↔ 패킹리스트 비교 (수량·입수·카톤) |
| `PI추출` | PI PDF에서 추출한 구조화 데이터 |
| `CI추출` | CI에서 파싱한 항목 |
| `PL추출` | PL에서 파싱한 항목 |
| `박스내용원본` | 박스내용 엑셀 파싱 결과 |
| `파싱로그` | PDF별 파싱 상태 및 오류 |

## 상태 코드

| 상태 | 색상 | 의미 |
|------|------|------|
| 정상 | 🟢 녹색 | 일치 |
| 부분선적 | 🔵 파란색 | 박스내용과 일치하나 PI 전체 수량과 다름 (의도적 부분선적 가능) |
| 불일치 | 🔴 빨간색 | 수치 차이 → 수정 필요 |
| PI누락 | 🔴 빨간색 | 박스내용에 주문번호가 있으나 PI PDF 없음 |
| PI항목누락 | 🔴 빨간색 | PI PDF에서 해당 모델 항목을 찾지 못함 |
| PI미매칭 | 🔴 빨간색 | CI/PL 모델코드를 PI에서 찾지 못함 |
| OCR필요 | 🟡 노란색 | 스캔 PDF라 텍스트 추출 불가 |
| 확인필요 | 🟡 노란색 | 카톤수 불일치 등 경미한 차이 |
| 모델불명 | 🟣 보라색 | CI/PL 설명에서 모델코드 추출 실패 |

## 주의사항

- **다중 모델 PI** (PI 1장에 제품 여러 개): FOC 항목 추출이 불완전할 수 있습니다.
  `PI미매칭`으로 표시된 FOC 항목은 PI 원문을 직접 확인해 주세요.
- **스캔 PDF** (`OCR필요`): 2차 단계에서 PaddleOCR (Apache-2.0) 적용 예정.
- CI/PL의 FOC 항목 단가는 PI($0.00)와 달리 명목 가격이 기재되므로 단가 비교를 하지 않습니다.

## 필요 패키지

```bash
pip install openpyxl pypdf xlrd
```

또는 가상환경 사용:

```bash
python3 -m venv .venv
.venv/bin/pip install -r requirements.txt
.venv/bin/python3 pi_review_tool.py
```
