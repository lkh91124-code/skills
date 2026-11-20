---
name: construction-doc-review
description: >-
  건설 공사 내역서(BOQ)·원가계산서·도면(PDF/xlsx)을 검토·검증하는 스킬. 내역서 산술(수량×단가=금액),
  소계·합계, 중복·단위·단가 편차, 원가계산서 요율(간접노무비·보험료·일반관리비·이윤·부가세) 재계산과
  내역서 금액 연계, 도면목록↔도면 번호·축척·개정, 도면-내역 규격 대조를 수행하고 근거가 붙은 검토 보고서를 만든다.
  Use when the user asks to review or verify a Korean construction bill of quantities (BOQ),
  cost estimate (원가계산서), or drawing set. Triggers: 내역서, 공사비, 원가계산서, 도면, 도면목록,
  검토, 검증, 수량, BOQ, bill of quantities, cost estimate, drawing review.
license: Complete terms in LICENSE.txt
metadata:
  author: lkh91124-code
  last-updated: '2026-11-20'
  keywords:
  - 내역서
  - 원가계산서
  - 공사비
  - 도면
  - 도면목록
  - 검토
  - 검증
  - 수량
  - BOQ
  - bill of quantities
  - cost estimate
  - drawing review
---

# 건설 내역서·원가계산서·도면 검토

## 목적

사용자가 준 내역서(BOQ), 원가계산서, 도면 PDF 를 스크립트로 추출·자동 검사하고, 에이전트가 오탐을 걸러
**위치·근거·권고가 붙은 검토 보고서**를 만든다. 결과는 검토 보조 자료이며 공인 검토·감리 의견이 아니다.

## 입력 확인

시작 전에 다음을 확인한다(모르면 사용자에게 묻는다).

1. 파일 종류와 위치: 내역서(xlsx/pdf), 원가계산서(xlsx/pdf), 도면(PDF 여러 개 가능), 도면목록 포함 여부
2. 내역서의 대상 시트/페이지 범위 (집계표·일위대가·단가산출서가 섞인 파일이면 `--sheet`/`--pages` 로 내역 부분만)
3. 원가계산 요율 기준(발주처·적용 연도). 모르면 문서 표기 % 로 검사하고 한계에 적는다 → `references/rates-guide.md`
4. `.xls`(구형)은 지원하지 않으므로 xlsx 로 저장 요청. 수식 계산값이 없는 xlsx 는 Excel 에서 열어 저장 요청

## 워크플로우

스크립트는 스킬 폴더 기준 `scripts/` 에 있다. 아래에서 `SKILL_DIR` 은 이 SKILL.md 가 있는 폴더,
`WORK` 는 임시 작업 폴더(예: `mktemp -d`)다. **사용자 원본 파일은 읽기만 하고, 모든 출력은 `WORK` 에 쓴다.**

```bash
SKILL_DIR=<이 스킬 폴더>; WORK=$(mktemp -d)
```

### 1. 내역서 추출·검사

```bash
python3 "$SKILL_DIR/scripts/extract_boq.py" 내역서.xlsx [--sheet 내역서] [--header-row 4] --out "$WORK/boq.json"
python3 "$SKILL_DIR/scripts/extract_boq.py" 내역서.pdf [--pages 3-20] --out "$WORK/boq.json"
python3 "$SKILL_DIR/scripts/check_boq.py" "$WORK/boq.json" [--tol 1] [--outlier-pct 5] --out "$WORK/findings_boq.json"
```

- 추출 후 `stats`(rows, low_confidence_ratio, scanned_pages)와 `warnings` 를 먼저 본다.
  - 헤더를 못 찾으면: 파일을 열어 헤더 행 번호를 확인하고 `--header-row` 지정
  - "수식 계산값이 저장되어 있지 않습니다" → 사용자에게 Excel 에서 열어 저장한 파일을 요청
  - 행 수가 원본과 크게 다르면 표본 몇 행을 원본과 대조
- 2행 헤더(재료비/노무비/경비/합계 × 단가/금액), 병합 셀, 페이지마다 반복되는 헤더는 자동 처리된다.

### 2. 스캔 페이지·저신뢰 행 처리

`scanned_pages` 나 `low_confidence: true` 행이 있으면 해당 페이지를 이미지로 만들어 직접 읽는다.

```bash
pdftoppm -r 150 -f N -l N -png 원본.pdf "$WORK/page"   # → $WORK/page-N.png (자릿수 패딩 가능)
```

이미지를 열어 필요한 값을 판독하고, 그 값으로 낸 지적은 근거에 **`근거: 이미지판독`** 을 붙인다.
판독이 불확실하면 "확인 필요"로 남긴다. 페이지가 많으면 금액이 큰 공종·소계 페이지부터 본다.

### 3. 원가계산서 검사

```bash
python3 "$SKILL_DIR/scripts/check_cost_summary.py" 원가계산서.xlsx [--sheet 원가계산서] \
  --boq "$WORK/boq.json" [--rates "$WORK/rates.json"] [--tol 1] --out "$WORK/findings_cost.json"
```

- `--rates` 는 사용자가 기준 요율을 준 경우에만 만든다(형식: `references/rates-guide.md`). 요율을 추측해 넣지 않는다.
- 결과 JSON 의 `parsed` 로 비목 인식이 맞는지 확인한다. 경고의 "역산 요율"은 요율 정보가 없을 때의 참고값이다.

### 4. 도면 추출·검사·교차 확인

```bash
python3 "$SKILL_DIR/scripts/extract_drawings.py" a.pdf b.pdf [--title-region 0.6,0.75,1,1] --out "$WORK/dwg.json"
python3 "$SKILL_DIR/scripts/check_drawings.py" "$WORK/dwg.json" --out "$WORK/findings_dwg.json"
python3 "$SKILL_DIR/scripts/cross_check.py" "$WORK/boq.json" "$WORK/dwg.json" --out "$WORK/findings_x.json"
```

- `--title-region` 은 표제란 위치(페이지 대비 비율 x0,y0,x1,y1, 기본 우하단 0.6,0.75,1,1). 표제란이 다른 곳이면
  한 페이지를 `pdftoppm` 으로 보고 비율을 정해 다시 실행한다.
- 스캔 도면은 2단계와 같이 이미지로 표제란을 판독한다.

### 5. 보고서 초안 생성

```bash
python3 "$SKILL_DIR/scripts/render_report.py" "$WORK"/findings_*.json --out "$WORK/report.md"
```

### 6. 에이전트 검토 (필수)

자동 결과를 그대로 전달하지 않는다.

1. **오탐 제거·병합**: 각 지적의 위치를 원본(또는 이미지)으로 확인. 추출 오류(열 밀림, 줄바꿈)로 생긴 지적은 삭제,
   연쇄 지적(항목 오류 → 소계 → 합계)은 원인 1건으로 묶는다. 의도된 반복(층별 동일 수량 B04 등)은 근거와 함께 낮춤.
2. **B06 누락 공종 판단**: `references/error-catalog.md` 체크리스트와 내역서 공종 목록을 대조해 누락 의심을 추가.
3. **X01 정리**: 주요 부재(콘크리트 강도, 철근 규격·강종, 강재 단면) 위주로 남기고 잡음은 요약.
4. 심각도는 `references/error-catalog.md` 기준으로 조정. 가정(요율 출처, 소계 범위 추정 등)은 근거에 명시.

## 보고서 형식

```markdown
# <공사명> 내역서·원가계산서·도면 검토 결과

## 요약
- 검토 대상: 파일명(시트/페이지 범위)
- 지적: 높음 N건 / 중간 N건 / 낮음 N건 — 핵심 3~5줄 (금액 영향 큰 순)

## 지적 사항
| 번호 | 코드 | 심각도 | 위치(페이지·행) | 내용 | 근거 | 권고 |
|---|---|---|---|---|---|---|
| 1 | B01 | 높음 | 시트 '내역서' 8행 (안전난간) | 합계 금액 산술 오류 | 50 × 5,000 = 250,000, 기재 260,000 | 금액 재계산 |

## 검토 한계
- 자동 추출 기반이며 스캔·저신뢰 페이지(목록)는 이미지 판독/미검토
- 요율 적정성은 미검증(문서 표기 기준 계산 일치 여부만 확인) / 적용 요율 출처
- 단가 적정성(물가·일위대가)·수량 산출 근거(도면 물량 산출)는 범위 밖
```

## 규칙

- **사용자 파일을 절대 수정·덮어쓰지 않는다.** 출력은 작업 폴더에만 쓴다.
- 가정·추정은 반드시 "가정:"으로 표시(요율 출처, 소계 범위, 단위 해석 등). 요율·법정 기준을 기억으로 단정하지 않는다.
- 지적마다 위치와 숫자 근거를 단다. 근거 없는 지적은 넣지 않는다.
- 결과는 **검토 보조 자료**이며 공인 적산·감리 검토를 대체하지 않는다는 문구를 보고서에 넣는다.
- 스크립트가 치명적 오류(종료 코드 ≠ 0)를 내면 메시지를 읽고 옵션(`--sheet`, `--pages`, `--header-row`)을 조정한다.
  행 단위 문제는 경고로 남고 처리는 계속된다.

## 참고

- `references/error-catalog.md` — 코드(B01~B06, C01~C03, D01~D04, X01)·심각도 기준·누락 공종 체크리스트
- `references/rates-guide.md` — 요율 산정기준, 연도별 변경 주의, `rates.json` 형식
- `scripts/_numparse.py` — 공용 숫자/단위 파서 (△·괄호 음수, 전각, ㎡→m2 등)
