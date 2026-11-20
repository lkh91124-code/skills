#!/usr/bin/env python3
"""원가계산서 검증: 요율 재계산(C01), 내역서 대비 기초금액(C02), 소계·총계 합산(C03).

사용법:
  python3 check_cost_summary.py 원가계산서.xlsx [--sheet 원가계산서] [--boq boq.json]
          [--rates rates.json] [--tol 1] --out findings_cost.json
  python3 check_cost_summary.py 원가계산서.pdf [--pages 1-2] ... --out findings_cost.json

요율은 --rates JSON 이 우선이며, 없으면 문서에 표기된 % (요율 열, 산출근거/비고, 비목명 괄호)를 쓴다.
둘 다 없으면 검사하지 않고 경고에 '역산 요율'만 남긴다.
금액 비교 시 1/10/100/1000원 단위 절사·반올림은 일치로 본다.

※ 아래 BASES(산정 기준)는 국내 공사원가계산에서 흔히 쓰는 관행이다.
  실제 적용 기준·요율은 발주처 지침·계약조건·해당 연도 고시에 따르므로 사용자/발주처 확인이 필요하다.
"""

from __future__ import annotations

import argparse
import os
import re
import sys
import unicodedata

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from _numparse import (fmt_num, make_finding, match_with_rounding, parse_num,  # noqa: E402
                       read_json, write_json)

# ---------------------------------------------------------------------------
# 비목 정의
# ---------------------------------------------------------------------------
# (표준키, 별칭들) — 별칭은 공백 제거·NFKC 후 포함 여부로 비교, 긴 별칭 우선
ALIASES = [
    ("직접재료비", ["직접재료비"]),
    ("간접재료비", ["간접재료비"]),
    ("작업설부산물", ["작업설", "부산물"]),
    ("직접노무비", ["직접노무비"]),
    ("간접노무비", ["간접노무비"]),
    ("기계경비", ["기계경비"]),
    ("산재보험료", ["산재보험", "산업재해보상보험"]),
    ("고용보험료", ["고용보험"]),
    ("노인장기요양보험료", ["노인장기요양", "장기요양보험"]),
    ("국민건강보험료", ["국민건강보험", "건강보험"]),
    ("국민연금보험료", ["국민연금", "연금보험"]),
    ("퇴직급여충당금", ["퇴직급여충당", "퇴직급여"]),
    ("산업안전보건관리비", ["산업안전보건관리비", "안전보건관리비", "안전관리비"]),
    ("환경보전비", ["환경보전비"]),
    ("기타경비", ["기타경비"]),
    ("일반관리비", ["일반관리비"]),
    ("이윤", ["이윤"]),
    ("순공사원가", ["순공사원가", "순공사비"]),
    ("총원가", ["총원가"]),
    ("공급가액", ["공급가액"]),
    ("부가가치세", ["부가가치세", "부가세", "vat"]),
    ("총공사비", ["총공사비", "도급액", "도급금액", "총공사금액", "공사비총계", "총사업비"]),
]
SUBTOTAL_HEADS = {"재료비": "재료비계", "노무비": "노무비계", "경비": "경비계"}
SUBTOTAL_WORDS = ("소계", "계", "합계")

# 비목의 소속 블록
CATEGORY = {
    "직접재료비": "재료", "간접재료비": "재료", "작업설부산물": "재료",
    "직접노무비": "노무", "간접노무비": "노무",
    "기계경비": "경비", "산재보험료": "경비", "고용보험료": "경비", "국민건강보험료": "경비",
    "노인장기요양보험료": "경비", "국민연금보험료": "경비", "퇴직급여충당금": "경비",
    "산업안전보건관리비": "경비", "환경보전비": "경비", "기타경비": "경비",
}
BLOCK_SUBTOTAL = {"재료": "재료비계", "노무": "노무비계", "경비": "경비계"}
ANCHORS = ["순공사원가", "총원가", "공급가액", "총공사비"]

# 요율 산정 기준(base). 값은 아래 집계명(합산 대상)의 튜플.
#   노무비 = 직접노무비 + 간접노무비, 재료비 = 재료비계, 경비 = 경비계
# 흔히 쓰는 관행일 뿐 — 발주처 지침에 따라 다를 수 있으므로 사용자 확인 필수.
BASES = {
    "간접노무비": ("직접노무비",),                    # 간접노무비 = 직접노무비 × r
    "산재보험료": ("노무비",),                        # 산재보험료 = (직접+간접)노무비 × r
    "고용보험료": ("노무비",),                        # 고용보험료 = (직접+간접)노무비 × r
    "국민건강보험료": ("직접노무비",),                # 건강보험료 = 직접노무비 × r
    "국민연금보험료": ("직접노무비",),                # 연금보험료 = 직접노무비 × r
    "노인장기요양보험료": ("국민건강보험료",),        # 노인장기요양 = 건강보험료 × r
    "퇴직급여충당금": ("노무비",),                    # 관행: 노무비 × r (발주처 확인)
    "산업안전보건관리비": ("재료비", "직접노무비"),    # 관행: (재료비+직접노무비) × r (관급자재 등 조정 별도)
    "환경보전비": ("재료비", "직접노무비", "기계경비"),  # 관행: (재료비+직접노무비+기계경비) × r
    "기타경비": ("재료비", "노무비"),                 # 관행: (재료비+노무비) × r
    "일반관리비": ("재료비", "노무비", "경비"),        # 일반관리비 = (재료비+노무비+경비) × r
    "이윤": ("노무비", "경비", "일반관리비"),          # 이윤 = (노무비+경비+일반관리비) × r
    "부가가치세": ("공급가액",),                      # 부가세 = 공급가액 × 10%
}
DEFAULT_RATES_PCT = {"부가가치세": 10.0}  # 법정 세율. 그 외 요율은 기본값 없음

PCT_RE = re.compile(r"(\d+(?:\.\d+)?)\s*%")


def nrm(s) -> str:
    s = unicodedata.normalize("NFKC", "" if s is None else str(s))
    return re.sub(r"[\s\[\]\(\)<>·ㆍ:_\-.,]", "", s).lower()


def canon(label: str):
    """라벨 → 표준 비목키. 소계류는 'SUB:<블록>' 또는 'SUB' 반환. 모르면 None."""
    raw = unicodedata.normalize("NFKC", label or "")
    raw = re.sub(r"^\s*(\d+(\.\d+)*[.)]?|[가-하]\.|\(\d+\))\s*", "", raw)  # 번호 제거
    raw = PCT_RE.sub("", raw)
    n = nrm(raw)
    if not n:
        return None
    for head, key in SUBTOTAL_HEADS.items():
        if n in (head, head + "계", head + "소계", head + "합계"):
            return key
    if n in SUBTOTAL_WORDS:
        return "SUB"
    best = None  # (위치, -길이, 키)
    for key, als in ALIASES:
        for a in als:
            p = n.find(nrm(a))
            if p >= 0:
                cand = (p, -len(a), key)
                if best is None or cand < best:
                    best = cand
    if best:
        return best[2]
    if n.endswith("계") and any(h in n for h in ("순공사", "원가")):
        return "순공사원가"
    return None


# ---------------------------------------------------------------------------
# 파싱
# ---------------------------------------------------------------------------
def rate_from_text(text):
    """문자열에서 마지막 % 값을 (pct, 소수자리수) 로."""
    m = list(PCT_RE.finditer(unicodedata.normalize("NFKC", text or "")))
    if not m:
        return None, None
    s = m[-1].group(1)
    return float(s), (len(s.split(".")[1]) if "." in s else 0)


def parse_xlsx(path, sheet, warnings):
    try:
        import openpyxl
    except ImportError:
        sys.exit("오류: openpyxl 이 설치되어 있지 않습니다.")
    try:
        wb = openpyxl.load_workbook(path, data_only=True)
    except Exception as e:  # noqa: BLE001
        sys.exit(f"오류: 엑셀 파일을 열 수 없습니다: {e}")
    if sheet and sheet not in wb.sheetnames:
        sys.exit(f"오류: 시트 '{sheet}' 가 없습니다. 시트 목록: {', '.join(wb.sheetnames)}")
    best = []
    for ws in ([wb[sheet]] if sheet else wb.worksheets):
        rows = parse_grid(
            [[(c.value, c.number_format) for c in r] for r in ws.iter_rows()],
            lambda ri, t=ws.title: f"시트 '{t}' {ri}행", warnings)
        if len([r for r in rows if r["key"]]) > len([r for r in best if r["key"]]):
            best = rows
    return best


def parse_grid(grid, locf, warnings):
    """셀 격자 [(값, 서식)] → 비목 행 목록."""
    amt_col = rate_col = None
    note_cols = set()
    start = 0
    for i, row in enumerate(grid[:15]):
        labels = [nrm(v) for v, _ in row]
        if any(lab in ("금액", "금액원", "공사금액") or lab.startswith("금액") for lab in labels):
            for j, lab in enumerate(labels):
                if lab.startswith("금액") or lab == "공사금액":
                    if amt_col is None:
                        amt_col = j
                elif ("요율" in lab or lab in ("율", "비율", "적용율", "적용률", "%")) and "구성" not in lab:
                    rate_col = j
                elif any(k in lab for k in ("비고", "산출근거", "근거", "계산식", "산식", "적용기준")):
                    note_cols.add(j)
            start = i + 1
            break
    out = []
    for ri in range(start, len(grid)):
        row = grid[ri]
        cells = [v for v, _ in row]
        try:
            amount = None
            if amt_col is not None and amt_col < len(cells):
                amount, _ = parse_num(cells[amt_col])
            label_cells = [(j, v) for j, v in enumerate(cells)
                           if isinstance(v, str) and v.strip() and j not in note_cols
                           and j != rate_col and (amt_col is None or j < amt_col)
                           and parse_num(v)[0] is None]
            if amt_col is None:  # 헤더가 없으면 1 이상인 가장 큰 숫자를 금액으로
                nums = [parse_num(v)[0] for j, (v, f) in enumerate(row)
                        if "%" not in (f or "") and not (isinstance(v, str) and "%" in v)]
                nums = [x for x in nums if x is not None and abs(x) >= 1]
                amount = max(nums, key=abs) if nums else None
            if not label_cells or amount is None:
                continue
            label = str(label_cells[-1][1])          # 가장 오른쪽(가장 구체적) 라벨
            key = canon(label)
            if key is None:                          # 오른쪽 라벨이 모르는 값이면 다른 셀도 시도
                for _, v in reversed(label_cells[:-1]):
                    if canon(str(v)):
                        key, label = canon(str(v)), str(v)
                        break
            rate, dec, src = None, None, None
            if rate_col is not None and rate_col < len(row):
                rv, rf = row[rate_col]
                if isinstance(rv, (int, float)) and not isinstance(rv, bool):
                    if "%" in (rf or ""):
                        rate, src = rv * 100, "문서(요율 열)"
                    else:
                        rate, src = float(rv), "문서(요율 열, 단위 추정)"
                elif isinstance(rv, str):
                    rate, dec = rate_from_text(rv)
                    if rate is None:
                        rate = parse_num(rv)[0]
                    src = "문서(요율 열)" if rate is not None else None
            if rate is None:
                texts = [str(cells[j]) for j in sorted(note_cols) if j < len(cells) and cells[j]]
                texts.append(label)
                for t in texts:
                    rate, dec = rate_from_text(t)
                    if rate is not None:
                        src = "문서(산출근거/비고)"
                        break
            out.append({"loc": locf(ri + 1), "label": label.strip(), "key": key,
                        "amount": amount, "rate_pct": rate, "rate_dec": dec, "rate_src": src})
        except Exception as e:  # noqa: BLE001 - 행 오류는 경고 후 계속
            warnings.append(f"{locf(ri + 1)} 처리 오류: {e}")
    return out


def parse_pdf(path, pages_spec, warnings):
    try:
        import pdfplumber
    except ImportError:
        sys.exit("오류: pdfplumber 가 설치되어 있지 않습니다.")
    from extract_boq import parse_pages
    out = []
    try:
        pdf = pdfplumber.open(path)
    except Exception as e:  # noqa: BLE001
        sys.exit(f"오류: PDF 를 열 수 없습니다: {e}")
    with pdf:
        for pno in parse_pages(pages_spec, len(pdf.pages)):
            page = pdf.pages[pno - 1]
            text = page.extract_text() or ""
            if len(re.sub(r"\s", "", text)) < 30:
                warnings.append(f"{pno}페이지: 스캔(이미지) 페이지로 보임 — 이미지 판독 필요")
                continue
            tables = page.extract_tables() or []
            got = []
            for t in tables:
                grid = [[(c, "") for c in r] for r in t]
                got += parse_grid(grid, lambda ri, p=pno: f"{p}페이지 표{ri}행", warnings)
            if not [r for r in got if r["key"]]:  # 표가 없으면 텍스트 줄 단위
                got = []
                for li, line in enumerate(text.splitlines(), 1):
                    nums = re.findall(r"[△▲\-]?\(?\d[\d,]*(?:\.\d+)?\)?(?!\s*%)", line)
                    vals = [parse_num(x)[0] for x in nums]
                    vals = [v for v in vals if v is not None and abs(v) >= 1000]
                    if not vals:
                        continue
                    # 라벨 = 첫 금액(1,000 이상) 앞의 텍스트
                    m = re.search(r"[△▲\-]?\(?\d{1,3}(?:,\d{3})+|\d{4,}", line)
                    label = (line[:m.start()] if m else line).strip() or line.strip()
                    key = canon(label)
                    rate, dec = rate_from_text(line)
                    got.append({"loc": f"{pno}페이지 {li}번째 줄", "label": label[:40],
                                "key": key, "amount": vals[-1], "rate_pct": rate, "rate_dec": dec,
                                "rate_src": "문서(본문 %)" if rate is not None else None})
            out += got
    return out


# ---------------------------------------------------------------------------
# 검사
# ---------------------------------------------------------------------------
def assign_blocks(rows, warnings):
    """소계(SUB) 해석과 블록(재료/노무/경비/post) 지정."""
    block = "재료"
    for r in rows:
        k = r["key"]
        if k in CATEGORY:
            block = CATEGORY[k]
        if k == "SUB":
            if block in BLOCK_SUBTOTAL:
                r["key"] = k = BLOCK_SUBTOTAL[block]
            else:
                warnings.append(f"{r['loc']}: '{r['label']}' 소계 대상 판단 불가 — 제외")
                r["key"] = None
                continue
        r["block"] = block if k not in BLOCK_SUBTOTAL.values() and k not in ANCHORS else None
        if k in BLOCK_SUBTOTAL.values():
            block = {"재료비계": "노무", "노무비계": "경비", "경비계": "post"}[k]
        elif k in ANCHORS or k in ("일반관리비", "이윤", "부가가치세"):
            block = "post"
            r["block"] = "post" if k not in ANCHORS else None


def build_values(rows):
    """비목 → 금액. 같은 비목이 여러 번 나오면 첫 값 사용."""
    v = {}
    for r in rows:
        if r["key"] and r["key"] not in v:
            v[r["key"]] = r["amount"]
    blk = lambda b: sum(r["amount"] for r in rows if r.get("block") == b)  # noqa: E731
    agg = dict(v)
    agg["재료비"] = v.get("재료비계", blk("재료") if any(r.get("block") == "재료" for r in rows) else None)
    if "노무비계" in v:
        agg["노무비"] = v["노무비계"]
    elif "직접노무비" in v:
        agg["노무비"] = v["직접노무비"] + (v.get("간접노무비") or 0)
    agg["경비"] = v.get("경비계", blk("경비") if any(r.get("block") == "경비" for r in rows) else None)
    if "공급가액" not in agg:
        agg["공급가액"] = v.get("총원가")
    return v, agg


def check_rates(rows, agg, rates_json, tol, F, warnings):
    for r in rows:
        k = r["key"]
        if k not in BASES or r["amount"] is None:
            continue
        parts = BASES[k]
        if any(agg.get(p) is None for p in parts):
            warnings.append(f"{r['loc']} {k}: 산정기준({'+'.join(parts)}) 값을 찾지 못해 요율 검사 생략")
            continue
        base = sum(agg[p] for p in parts)
        if base == 0:
            continue
        implied = r["amount"] / base * 100
        doc_rate, src = r.get("rate_pct"), r.get("rate_src")
        if doc_rate is not None and src and "추정" in src and doc_rate < 1:
            # 서식 없는 소수(0.035 등): % 값/분수 중 역산 요율에 가까운 쪽 선택
            doc_rate = min((doc_rate, doc_rate * 100), key=lambda x: abs(x - implied))
        j_rate = rates_json.get(k)
        if j_rate is None and k in DEFAULT_RATES_PCT and doc_rate is None:
            j_rate = DEFAULT_RATES_PCT[k]
        if j_rate is not None and doc_rate is not None and abs(j_rate - doc_rate) > 1e-9:
            F.append(make_finding(
                "C01", "중간", f"{r['loc']} ({r['label']})", f"{k} 요율 표기가 기준 요율과 다름",
                f"문서 표기 {doc_rate:g}% vs 기준(rates.json) {j_rate:g}%",
                "적용 요율 근거(발주처 지침·고시 연도) 확인"))
        rate = j_rate if j_rate is not None else doc_rate
        rsrc = "rates.json" if j_rate is not None and k in rates_json else (
            "법정 세율" if j_rate is not None else src)
        if rate is None:
            warnings.append(f"{r['loc']} {k}: 요율 정보 없음 — 역산 요율 {implied:.4f}% "
                            f"(기준 {'+'.join(parts)} {fmt_num(base)}). 사용자 확인 필요")
            continue
        expected = base * rate / 100
        if match_with_rounding(expected, r["amount"], tol):
            continue
        dec = r.get("rate_dec")
        # 문서에 소수 2자리 이상으로 표기된 요율은 표시 반올림(±0.5 최소단위) 범위를 허용
        if rsrc and rsrc.startswith("문서") and dec is not None and dec >= 2 and \
                abs(implied - rate) <= 0.5 * 10 ** (-dec) + 1e-9:
            warnings.append(f"{r['loc']} {k}: 표기 요율 {rate:g}% 의 반올림 범위 내 차이 "
                            f"(역산 {implied:.4f}%) — 실제 적용 요율 확인 권장")
            continue
        F.append(make_finding(
            "C01", "높음", f"{r['loc']} ({r['label']})", f"{k} 요율 계산 불일치",
            f"{'+'.join(parts)} {fmt_num(base)} × {rate:g}% ({rsrc}) = {fmt_num(expected)}, "
            f"기재 {fmt_num(r['amount'])} (차이 {fmt_num(r['amount'] - expected)}, 역산 요율 {implied:.4f}%)",
            "적용 요율·산정기준·단수처리 확인 후 재계산"))


def check_sums(rows, v, tol, F):
    def chk(key, expected, how):
        if key not in v or expected is None:
            return
        if not match_with_rounding(expected, v[key], tol):
            loc = next(f"{r['loc']} ({r['label']})" for r in rows if r["key"] == key)
            F.append(make_finding(
                "C03", "높음", loc, f"{key} 합산 불일치",
                f"{how} = {fmt_num(expected)}, 기재 {fmt_num(v[key])} (차이 {fmt_num(v[key] - expected)})",
                "합산 범위·누락 비목 확인"))

    for b, sk in BLOCK_SUBTOTAL.items():
        items = [r for r in rows if r.get("block") == b]
        if items:
            chk(sk, sum(r["amount"] for r in items), " + ".join(r["label"] for r in items[:6])
                + (" 외" if len(items) > 6 else ""))
    if all(k in v for k in ("재료비계", "노무비계", "경비계")):
        chk("순공사원가", v["재료비계"] + v["노무비계"] + v["경비계"], "재료비계+노무비계+경비계")
    # 앵커 체인: 다음 앵커 = 이전 앵커 + 사이의 비목들
    idx = [(i, r) for i, r in enumerate(rows) if r["key"] in ANCHORS]
    for (i0, a), (i1, b) in zip(idx, idx[1:]):
        mids = [r for r in rows[i0 + 1:i1] if r["key"] not in BLOCK_SUBTOTAL.values()
                and r["key"] not in ANCHORS]
        chk(b["key"], a["amount"] + sum(r["amount"] for r in mids),
            " + ".join([a["key"]] + [r["label"] for r in mids]))


_SUMMARY_SEG_WORDS = ("집계", "총괄", "요약", "갑지", "summary")


def _is_summary_segment(seg, items):
    """집계표/총괄표 구간 판단: 시트명이 집계류이거나, 항목이 전부 '식' 단위 수량 1 인 구간."""
    if any(w in (seg or "").lower() for w in _SUMMARY_SEG_WORDS):
        return True
    return bool(items) and all(
        (r.get("unit") or "").strip() in ("식", "式", "LS", "ls", "L.S")
        and (r.get("qty") in (None, 1, 1.0)) for r in items)


def check_boq(agg, v, boq, tol, F, warnings):
    """C02: 원가계산서 직접재료비·직접노무비·기계경비 vs 내역서 금액.

    내역서에 집계표와 내역 시트가 함께 있으면 전체 합산은 이중 계상이 되므로 구간(시트) 별로
    비교한다. 어느 구간의 항목 합 또는 합계행과 맞으면 통과. 집계 구간을 뺀 구간 전체 합
    (공종별로 시트가 나뉜 내역서)도 후보로 둔다.
    """
    rows = boq.get("rows", [])
    segs = {}
    for r in rows:
        sg = segs.setdefault(r.get("segment", ""), {"items": [], "totals": []})
        if r.get("type") == "item":
            sg["items"].append(r)
        elif r.get("type") == "total":
            sg["totals"].append(r)
    detail = [s for s, d in segs.items() if d["items"] and not _is_summary_segment(s, d["items"])]
    summ = [s for s, d in segs.items() if d["items"] and _is_summary_segment(s, d["items"])]
    if summ:
        warnings.append(f"내역서 집계 구간 {summ} 은 C02 전체 합산에서 제외하고 구간별로만 비교")
    # 대표 구간: 항목 행이 가장 많은 구간 (보고용)
    main_seg = max(segs, key=lambda s: len(segs[s]["items"])) if segs else ""
    pairs = [("직접재료비", "mat_amt", "재료비"), ("직접노무비", "lab_amt", "노무비"),
             ("기계경비", "exp_amt", "경비")]
    for ck, col, lab in pairs:
        key = ck if ck in v else ({"직접재료비": "재료비계"}.get(ck))
        if key not in v or v[key] is None:
            warnings.append(f"원가계산서에 {ck} 항목이 없어 내역서 대비 검사 생략")
            continue
        cands = []  # (이름, 값, 대표여부)
        for sname, d in segs.items():
            lab_s = f"'{sname}'" if sname else "내역"
            if d["items"]:
                cands.append((f"{lab_s} 항목 합", sum(r.get(col) or 0 for r in d["items"]),
                              sname == main_seg))
            if d["totals"]:
                cands.append((f"{lab_s} 합계행", d["totals"][-1].get(col) or 0, sname == main_seg))
        if len(detail) > 1:
            cands.append((f"구간 {detail} 항목 합", sum(r.get(col) or 0 for s in detail
                                                       for r in segs[s]["items"]), False))
        cands = [c for c in cands if c[1] != 0]
        if not cands:
            continue
        if any(match_with_rounding(c, v[key], tol) for _, c, _ in cands):
            continue
        shown = sorted(cands, key=lambda c: not c[2])[:4]
        F.append(make_finding(
            "C02", "높음", f"원가계산서 {key}",
            f"{key}가 내역서 {lab} 금액과 불일치",
            f"원가계산서 {fmt_num(v[key])} vs " + ", ".join(f"{n} {fmt_num(c)}" for n, c, _ in shown)
            + (" (내역서 저신뢰/스캔 페이지 있음)" if (boq.get('stats') or {}).get('scanned_pages') else ""),
            "내역서 집계 금액과 원가계산서 기초금액(관급자재·제외 공종 포함 여부) 확인"))


def main():
    for st in (sys.stdout, sys.stderr):
        try:
            st.reconfigure(encoding="utf-8", errors="replace")
        except Exception:  # noqa: BLE001
            pass
    ap = argparse.ArgumentParser(description="원가계산서 요율·합산·내역서 연계 검사")
    ap.add_argument("file", help="원가계산서 (.xlsx/.pdf)")
    ap.add_argument("--sheet", help="엑셀 시트 이름")
    ap.add_argument("--pages", help="PDF 페이지 범위")
    ap.add_argument("--boq", help="extract_boq.py 출력 JSON (C02 검사용)")
    ap.add_argument("--rates", help="요율 JSON (references/rates-guide.md 참조)")
    ap.add_argument("--tol", type=float, default=1.0, help="금액 허용오차(원), 기본 1")
    ap.add_argument("--out", required=True, help="출력 findings JSON")
    args = ap.parse_args()

    if not os.path.isfile(args.file):
        sys.exit(f"오류: 파일이 없습니다: {args.file}")
    warnings, F = [], []
    rates = {}
    if args.rates:
        try:
            rj = read_json(args.rates)
            rates = {k: float(x) for k, x in (rj.get("rates_pct") or {}).items() if x is not None}
            if rj.get("note"):
                warnings.append(f"요율 출처: {rj['note']}")
        except Exception as e:  # noqa: BLE001
            sys.exit(f"오류: 요율 JSON 을 읽을 수 없습니다: {e}")
    ext = os.path.splitext(args.file)[1].lower()
    if ext in (".xlsx", ".xlsm"):
        rows = parse_xlsx(args.file, args.sheet, warnings)
    elif ext == ".pdf":
        rows = parse_pdf(args.file, args.pages, warnings)
    else:
        sys.exit(f"오류: 지원하지 않는 형식입니다: {ext} (xlsx, pdf 만 지원)")
    found = {r["key"] for r in rows if r["key"] and r["key"] != "SUB"}
    if len(found - set(SUBTOTAL_HEADS.values())) < 3:
        sys.exit("오류: 원가계산서 비목(직접재료비, 직접노무비, 일반관리비 등)을 충분히 찾지 못했습니다. "
                 "원가계산서 파일/시트(--sheet)/페이지(--pages)가 맞는지 확인하세요.")

    unknown = [r for r in rows if not r["key"]]
    for r in unknown:
        warnings.append(f"{r['loc']}: 미분류 비목 '{r['label']}' ({fmt_num(r['amount'])}) — 소속 블록 합산에만 반영")
    assign_blocks(rows, warnings)
    v, agg = build_values(rows)
    for name, fn in (("요율", lambda: check_rates(rows, agg, rates, args.tol, F, warnings)),
                     ("합산", lambda: check_sums(rows, v, args.tol, F))):
        try:
            fn()
        except Exception as e:  # noqa: BLE001
            warnings.append(f"{name} 검사 중 오류: {e!r}")
    if args.boq:
        try:
            check_boq(agg, v, read_json(args.boq), args.tol, F, warnings)
        except FileNotFoundError:
            sys.exit(f"오류: 파일이 없습니다: {args.boq}")
        except Exception as e:  # noqa: BLE001
            warnings.append(f"내역서 대비 검사 중 오류: {e!r}")

    write_json({"source": os.path.abspath(args.file), "findings": F, "warnings": warnings,
                "parsed": [{k: r.get(k) for k in ("loc", "label", "key", "amount", "rate_pct", "rate_src")}
                           for r in rows]}, args.out)
    print(f"원가계산서 검사 완료: 비목 {len(rows)}개, 지적 {len(F)}건, 경고 {len(warnings)}건 → {args.out}")


if __name__ == "__main__":
    main()
