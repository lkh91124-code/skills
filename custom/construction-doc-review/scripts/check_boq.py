#!/usr/bin/env python3
"""내역서(BOQ) 자동 검증: extract_boq.py 결과(boq.json) → findings JSON.

사용법:
  python3 check_boq.py boq.json [--tol 1] [--outlier-pct 5] --out findings_boq.json

검사 코드:
  B01 (높음) 산술 오류: 수량×단가≠금액 (재료/노무/경비/합계), 재+노+경≠합계 (단가·금액)
  B02 (높음) 소계/합계 ≠ 하위 항목 금액 합
  B03 (중간) 같은 품명·규격인데 단위가 다름
  B04 (중간) 중복 항목 (같은 공종 경로 + 품명·규격 + 수량)
  B05 (중간) 같은 품명·규격인데 단가 차이가 기준(%) 초과
금액 비교 시 1/10/100/1000원 단위 절사·반올림은 일치로 본다.
"""

from __future__ import annotations

import argparse
import os
import sys
from collections import defaultdict

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from _numparse import (fmt_num, make_finding, match_with_rounding, norm_key,  # noqa: E402
                       norm_unit, read_json, write_json)

COMP = [("mat", "재료비"), ("lab", "노무비"), ("exp", "경비"), ("tot", "합계")]
AMT_COLS = [("mat_amt", "재료비"), ("lab_amt", "노무비"), ("exp_amt", "경비"), ("tot_amt", "합계")]
UP_COLS = [("mat_up", "재료비단가"), ("lab_up", "노무비단가"), ("exp_up", "경비단가"),
           ("tot_up", "합계단가")]


def loc_str(r) -> str:
    loc = r.get("loc", {})
    if "sheet" in loc:
        s = f"시트 '{loc['sheet']}' {loc.get('row')}행"
    else:
        s = f"{loc.get('page')}페이지 표{loc.get('row')}행"
    label = r.get("name") or r.get("code") or ""
    if r.get("spec"):
        label += f" [{r['spec']}]"
    return f"{s} ({label})" if label else s


def item_label(r) -> str:
    return (r.get("name") or "") + (f" {r['spec']}" if r.get("spec") else "")


def check_b01(rows, tol, F):
    for r in rows:
        if r["type"] not in ("item", "subtotal", "total"):
            continue
        conf = " (저신뢰 추출 행 — 원본 확인 필요)" if r.get("low_confidence") else ""
        # 1) 수량 × 단가 = 금액 (항목 행만)
        if r["type"] == "item" and r.get("qty") is not None:
            for key, label in COMP:
                up, amt = r.get(f"{key}_up"), r.get(f"{key}_amt")
                if up is None or amt is None:
                    continue
                exp = r["qty"] * up
                if not match_with_rounding(exp, amt, tol):
                    F.append(make_finding(
                        "B01", "높음", loc_str(r),
                        f"{label} 금액 산술 오류",
                        f"수량 {fmt_num(r['qty'])} × 단가 {fmt_num(up)} = {fmt_num(exp)}, "
                        f"기재 {fmt_num(amt)} (차이 {fmt_num(amt - exp)}){conf}",
                        "수량·단가·금액 셀을 원본에서 확인하고 금액 재계산"))
        # 2) 재료비 + 노무비 + 경비 = 합계 (단가·금액)
        for suffix, what in (("up", "단가"), ("amt", "금액")):
            if r["type"] != "item" and suffix == "up":
                continue
            tot = r.get(f"tot_{suffix}")
            parts = [r.get(f"{k}_{suffix}") for k in ("mat", "lab", "exp")]
            if tot is None or all(p is None for p in parts):
                continue
            s = sum(p or 0 for p in parts)
            if not match_with_rounding(s, tot, tol):
                F.append(make_finding(
                    "B01", "높음", loc_str(r),
                    f"재료비+노무비+경비 ≠ 합계 ({what})",
                    f"{' + '.join(fmt_num(p) for p in parts)} = {fmt_num(s)}, "
                    f"합계 기재 {fmt_num(tot)} (차이 {fmt_num(tot - s)}){conf}",
                    f"합계{what} 산식 확인"))


def check_b02(rows, tol, F, warnings):
    by_seg = defaultdict(list)
    for i, r in enumerate(rows):
        by_seg[r.get("segment", "")].append((i, r))
    for seg, srows in by_seg.items():
        last_break = -1  # 직전 소계/합계 위치 (closes 가 없는 소계용)
        last_total = -1
        for pos, (i, r) in enumerate(srows):
            if r["type"] not in ("subtotal", "total"):
                continue
            items_before = [x for _, x in srows[:pos] if x["type"] == "item"]
            # 비교 후보 범위들
            cands = []
            if r["type"] == "total":
                cands.append(("구간 내 전체 항목", items_before))
                cands.append(("직전 합계 이후 항목",
                              [x for _, x in srows[last_total + 1:pos] if x["type"] == "item"]))
            else:
                for gid in r.get("scope") or []:
                    cands.append((f"그룹#{gid}", [x for x in items_before if gid in (x.get("gids") or [])]))
                cands.append(("직전 소계 이후 항목",
                              [x for _, x in srows[last_break + 1:pos] if x["type"] == "item"]))
            cands = [(n, c) for n, c in cands if c]
            last_break = pos
            if r["type"] == "total":
                last_total = pos
            if not cands:
                warnings.append(f"{loc_str(r)}: 비교할 하위 항목이 없어 소계 검사 생략")
                continue
            for col, label in AMT_COLS:
                v = r.get(col)
                if v is None:
                    continue
                sums = [(n, sum(x.get(col) or 0 for x in c), len(c)) for n, c in cands]
                if any(match_with_rounding(s, v, tol) for _, s, _ in sums):
                    continue
                n0, s0, k0 = sums[0]
                kind = "합계" if r["type"] == "total" else "소계"
                path = " > ".join(r.get("path") or []) or "(경로 미상)"
                F.append(make_finding(
                    "B02", "높음", loc_str(r),
                    f"{kind} {label} 금액이 하위 항목 합과 불일치 (공종: {path})",
                    f"하위 {k0}개 항목 합 {fmt_num(s0)} vs 기재 {fmt_num(v)} "
                    f"(차이 {fmt_num(v - s0)}; 비교범위: {n0})"
                    + (" — 소계 범위 추정이 불확실할 수 있음" if r.get("closes") is None else ""),
                    f"{kind} 범위와 합산식(SUM 범위 누락/중복) 확인"))


def check_keys(rows, pct, F):
    items = [r for r in rows if r["type"] == "item" and (r.get("name") or "").strip()]
    by_key = defaultdict(list)
    for r in items:
        by_key[norm_key(r.get("name"), r.get("spec"))].append(r)

    for key, grp in by_key.items():
        # B03 단위 불일치
        units = defaultdict(list)
        for r in grp:
            if r.get("unit"):
                units[norm_unit(r["unit"])].append(r)
        if len(units) > 1:
            desc = "; ".join(f"'{u}': {', '.join(loc_str(x) for x in rs[:3])}"
                             + (f" 외 {len(rs) - 3}건" if len(rs) > 3 else "")
                             for u, rs in units.items())
            F.append(make_finding(
                "B03", "중간", loc_str(grp[0]),
                f"동일 품명·규격 단위 불일치: {item_label(grp[0])}",
                f"단위별 위치 — {desc}",
                "단위 통일 여부와 수량 환산(㎡↔㎥, EA↔SET 등) 확인"))

        # B04 중복 (같은 경로 + 키 + 수량)
        dup = defaultdict(list)
        for r in grp:
            if r.get("qty") is None:
                continue
            dup[(tuple(r.get("path") or []), round(r["qty"], 6))].append(r)
        for (path, q), rs in dup.items():
            if len(rs) > 1:
                F.append(make_finding(
                    "B04", "중간", loc_str(rs[1]),
                    f"중복 항목 의심: {item_label(rs[0])} (수량 {fmt_num(q)})",
                    "동일 공종(" + (" > ".join(path) or "최상위") + ") 내 반복: "
                    + ", ".join(loc_str(x) for x in rs),
                    "의도된 반복(구간·층별)인지 확인 후 중복이면 삭제"))

        # B05 단가 편차 (같은 단위끼리만 비교)
        for u, rs in units.items() if units else [("", grp)]:
            if len(rs) < 2:
                continue
            msgs = []
            for col, label in UP_COLS:
                vals = [(r.get(col), r) for r in rs if r.get(col) not in (None, 0)]
                if len(vals) < 2:
                    continue
                lo = min(vals, key=lambda t: t[0])
                hi = max(vals, key=lambda t: t[0])
                if lo[0] <= 0:
                    continue
                dev = (hi[0] - lo[0]) / lo[0] * 100
                if dev > pct:
                    msgs.append(f"{label} {fmt_num(lo[0])}({loc_str(lo[1])}) ~ "
                                f"{fmt_num(hi[0])}({loc_str(hi[1])}), 편차 {dev:.1f}%")
            if msgs:
                F.append(make_finding(
                    "B05", "중간", loc_str(rs[0]),
                    f"동일 품명·규격 단가 상이: {item_label(rs[0])} ({u or '단위없음'})",
                    "; ".join(msgs),
                    "적용 단가(일위대가·물가자료) 및 할증 적용 여부 확인"))


def main():
    ap = argparse.ArgumentParser(description="내역서 산술·소계·중복·단위·단가 자동 검사")
    ap.add_argument("boq", help="extract_boq.py 출력 JSON")
    ap.add_argument("--tol", type=float, default=1.0, help="금액 허용오차(원), 기본 1")
    ap.add_argument("--outlier-pct", type=float, default=5.0, help="B05 단가 편차 기준(%%), 기본 5")
    ap.add_argument("--out", required=True, help="출력 findings JSON")
    args = ap.parse_args()

    try:
        boq = read_json(args.boq)
    except FileNotFoundError:
        sys.exit(f"오류: 파일이 없습니다: {args.boq}")
    except Exception as e:  # noqa: BLE001
        sys.exit(f"오류: JSON 을 읽을 수 없습니다: {e}")
    rows = boq.get("rows")
    if not isinstance(rows, list):
        sys.exit("오류: boq.json 형식이 아닙니다 (rows 없음). extract_boq.py 출력을 넣어 주세요.")

    F, warnings = [], list(boq.get("warnings") or [])
    for fn in (lambda: check_b01(rows, args.tol, F),
               lambda: check_b02(rows, args.tol, F, warnings),
               lambda: check_keys(rows, args.outlier_pct, F)):
        try:
            fn()
        except Exception as e:  # noqa: BLE001 - 한 검사가 실패해도 나머지 진행
            warnings.append(f"검사 중 오류: {e!r}")
    st = boq.get("stats") or {}
    if st.get("scanned_pages"):
        warnings.append(f"스캔 페이지 {st['scanned_pages']} 는 자동 검사에서 제외됨 — 이미지 판독 필요")
    if (st.get("low_confidence_ratio") or 0) > 0.2:
        warnings.append(f"저신뢰 행 비율 {st['low_confidence_ratio']:.0%} — 원본 대조 필요")

    write_json({"source": boq.get("source", args.boq), "findings": F, "warnings": warnings}, args.out)
    cnt = defaultdict(int)
    for f in F:
        cnt[f["code"]] += 1
    print(f"내역서 검사 완료: 지적 {len(F)}건 {dict(sorted(cnt.items()))}, 경고 {len(warnings)}건 → {args.out}")


if __name__ == "__main__":
    main()
