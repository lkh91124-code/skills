#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
내역서(boq.json) ↔ 도면(dwg.json) 규격 교차 확인 (X01).

사용법:
    python3 cross_check.py boq.json dwg.json --out findings_x.json

방법:
- 도면 페이지별 spec_tokens 와, 내역서 항목(name+spec)에서 같은 규칙(extract_drawings.extract_spec_tokens)
  으로 뽑은 토큰을 카테고리별로 정규화하여 비교한다.
  (콘크리트 강도 fck, 철근 호칭 D, 철근 강종 SD/HD, 강재 단면 H-, 두께 THK, 관경 Φ, KS 규격, 부재 치수)
- 한쪽에만 나타나는 토큰을 "확인 필요"로 보고한다. 이는 오류 확정이 아니라 검토 단서이다.
  (도면 주기·일반사항의 표준 문구, 내역서의 일위대가식 표기 차이 등으로 오탐이 흔함)
- 잡음을 줄이기 위해 토큰 단위로 묶어 페이지/행을 나열하고, 잡음이 많은 카테고리(size 등)는
  일정 개수를 넘으면 한 건으로 요약한다. 배근 간격(@)은 내역서에 거의 나타나지 않으므로 비교하지 않는다.
"""
import argparse
import os
import sys
from collections import OrderedDict, defaultdict

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from _numparse import make_finding, read_json, write_json  # noqa: E402
from extract_drawings import extract_spec_tokens  # noqa: E402

CAT_LABEL = OrderedDict([
    ("concrete", "콘크리트 강도"),
    ("rebar_grade", "철근 강종"),
    ("rebar_dia", "철근 호칭경"),
    ("steel", "강재 단면"),
    ("thk", "두께"),
    ("pipe", "관경"),
    ("ks", "KS"),
    ("size", "부재 치수"),
])
# 한쪽에만 있을 때의 심각도 (도면에만 / 내역서에만)
SEV = {
    "concrete": ("중간", "중간"),
    "rebar_grade": ("중간", "중간"),
    "rebar_dia": ("중간", "낮음"),
    "steel": ("중간", "중간"),
    "thk": ("낮음", "낮음"),
    "pipe": ("낮음", "낮음"),
    "ks": ("낮음", "낮음"),
    "size": ("낮음", "낮음"),
}
# 카테고리별 개별 보고 상한(초과분은 1건으로 요약)
CAP = {"size": 10, "thk": 15, "pipe": 15, "ks": 10}
MAX_LOCS = 8  # 근거에 나열할 위치 최대 개수


def boq_loc(row):
    loc = row.get("loc") or {}
    if "sheet" in loc:
        return "%s 시트 %s행" % (loc.get("sheet"), loc.get("row"))
    if "page" in loc:
        return "p.%s 행%s" % (loc.get("page"), loc.get("row"))
    return "행 ?"


def fmt_locs(locs):
    locs = list(dict.fromkeys(locs))
    s = ", ".join(locs[:MAX_LOCS])
    if len(locs) > MAX_LOCS:
        s += " 외 %d곳" % (len(locs) - MAX_LOCS)
    return s


def main():
    for s in (sys.stdout, sys.stderr):
        try:
            s.reconfigure(encoding="utf-8", errors="replace")
        except Exception:  # noqa: BLE001
            pass
    ap = argparse.ArgumentParser(description="내역서-도면 규격 토큰 교차 확인(X01)")
    ap.add_argument("boq_json", help="extract_boq.py 출력 JSON")
    ap.add_argument("dwg_json", help="extract_drawings.py 출력 JSON")
    ap.add_argument("--out", required=True, help="결과 findings JSON 경로")
    args = ap.parse_args()

    for p in (args.boq_json, args.dwg_json):
        if not os.path.isfile(p):
            sys.stderr.write("오류: 파일을 찾을 수 없습니다: %s\n" % p)
            sys.exit(2)
    try:
        boq = read_json(args.boq_json)
        dwg = read_json(args.dwg_json)
        rows = boq["rows"]
        pages = dwg["pages"]
    except Exception as e:
        sys.stderr.write("오류: 입력 JSON을 읽을 수 없습니다: %s\n" % e)
        sys.exit(2)

    warnings = []
    # 위치 표기는 파일명만 사용(전체 경로는 보고서를 길게 하고 사용자 경로를 노출함)
    boq_src = os.path.basename(str(boq.get("source") or args.boq_json).replace("\\", "/"))

    # 도면 토큰 수집: (cat, token) -> [위치], 원문 예시
    d_tok, d_raw = defaultdict(list), {}
    n_scanned = 0
    for p in pages:
        if p.get("scanned"):
            n_scanned += 1
            continue
        for t in p.get("spec_tokens", []):
            if t["cat"] not in CAT_LABEL:
                continue
            key = (t["cat"], t["token"])
            d_tok[key].append("%s p.%d" % (os.path.basename(str(p["file"]).replace("\\", "/")), p["page"]))
            d_raw.setdefault(key, t.get("raw"))
    if n_scanned:
        warnings.append("X01: 스캔 도면 %d페이지는 규격 토큰을 추출하지 못해 교차 확인에서 제외됨 — 이미지로 확인 필요" % n_scanned)

    # 내역서 토큰 수집 (항목 행 위주, 소계/합계 제외)
    b_tok, b_raw = defaultdict(list), {}
    for r in rows:
        if r.get("type") in ("subtotal", "total"):
            continue
        text = " ".join(str(r.get(k) or "") for k in ("name", "spec"))
        for t in extract_spec_tokens(text):
            if t["cat"] not in CAT_LABEL:
                continue
            key = (t["cat"], t["token"])
            b_tok[key].append("%s %s" % (boq_src, boq_loc(r)))
            b_raw.setdefault(key, "%s / %s" % (r.get("name") or "", r.get("spec") or ""))

    if not d_tok:
        warnings.append("X01: 도면에서 규격 토큰을 찾지 못해 교차 확인 결과가 제한적임")
    if not b_tok:
        warnings.append("X01: 내역서에서 규격 토큰을 찾지 못해 교차 확인 결과가 제한적임")

    findings = []
    try:
        compare(b_tok, b_raw, d_tok, d_raw, findings)
    except Exception as e:  # noqa: BLE001 - 검사 실패는 숨기지 않고 E00 으로 보고
        findings.append(make_finding(
            "E00", "높음", "cross_check.py",
            "검사 실패: 도면-내역 규격 교차 확인(X01)을 끝내지 못함 — 이 검사 결과는 신뢰할 수 없음",
            "%s: %s" % (type(e).__name__, e),
            "입력 JSON(boq.json, dwg.json)을 확인하고 재실행. 해결되지 않으면 주요 규격을 수동 대조"))

    out = {"source": "%s ↔ %s" % (os.path.basename(args.boq_json), os.path.basename(args.dwg_json)),
           "findings": findings, "warnings": warnings}
    write_json(out, args.out)
    print("교차 확인 완료: 확인 필요 %d건, 경고 %d건 → %s" % (len(findings), len(warnings), args.out))


def compare(b_tok, b_raw, d_tok, d_raw, findings):
    """카테고리별로 한쪽에만 있는 토큰을 X01 로 보고."""
    for cat, label in CAT_LABEL.items():
        b_cat_tokens = sorted(t for (c, t) in b_tok if c == cat)
        d_cat_tokens = sorted(t for (c, t) in d_tok if c == cat)
        # 한쪽 카테고리가 통째로 비어 있으면 표기 방식 차이일 가능성이 커 요약 1건만
        only_d = [t for t in d_cat_tokens if (cat, t) not in b_tok]
        only_b = [t for t in b_cat_tokens if (cat, t) not in d_tok]

        for direction, toks, src_map, raw_map, other_has in (
            ("dwg", only_d, d_tok, d_raw, bool(b_cat_tokens)),
            ("boq", only_b, b_tok, b_raw, bool(d_cat_tokens)),
        ):
            if not toks:
                continue
            sev = SEV[cat][0 if direction == "dwg" else 1]
            cap = CAP.get(cat, 999)
            if direction == "dwg":
                where, other = "도면", "내역서"
                rec = "해당 규격 자재가 내역서에 누락되었는지, 다른 표기(예: 강종 HD↔SD, 25-24-150↔24MPa)로 기재되었는지 확인"
            else:
                where, other = "내역서", "도면"
                rec = "내역서 규격이 도면 근거가 있는지(설계변경·구도면 기준 여부) 확인"
            if not other_has:
                # 다른 쪽에 이 카테고리 토큰이 전혀 없음 → 개별 보고 대신 요약 1건(낮음)
                findings.append(make_finding(
                    "X01", "낮음", "%s 전반" % where,
                    "확인 필요: %s 규격(%s)이 %s에만 있고 %s에는 해당 종류 표기가 전혀 없음"
                    % (label, ", ".join(toks[:15]) + (" 외" if len(toks) > 15 else ""), where, other),
                    "예: %s (%s)" % (raw_map.get((cat, toks[0])), fmt_locs(src_map[(cat, toks[0])])),
                    "%s의 규격 표기 방식 확인(텍스트 미표기·이미지·별도 일위대가 등). %s" % (other, rec)))
                continue
            if len(toks) > cap:
                findings.append(make_finding(
                    "X01", "낮음", "%s 전반" % where,
                    "확인 필요: %s 규격 %d종이 %s에만 있음(다수라 요약)" % (label, len(toks), where),
                    ", ".join(toks[:30]) + (" 외" if len(toks) > 30 else ""),
                    rec))
                continue
            for t in toks:
                findings.append(make_finding(
                    "X01", sev, fmt_locs(src_map[(cat, t)]),
                    "확인 필요: %s '%s'이(가) %s에는 있으나 %s에서 찾지 못함" % (label, t, where, other),
                    "원문: %s / %s 내 동일 종류 규격: %s"
                    % (raw_map.get((cat, t)), other,
                       ", ".join(d_cat_tokens if direction == "boq" else b_cat_tokens)[:200]),
                    rec))


if __name__ == "__main__":
    main()
