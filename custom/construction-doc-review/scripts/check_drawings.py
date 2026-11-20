#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
도면 세트 자체 정합성 검사 (extract_drawings.py 결과 사용).

사용법:
    python3 check_drawings.py dwg.json --out findings_dwg.json

검사 코드:
    D01  도면목록 ↔ 실제 도면 불일치
         - 도면에 있으나 목록에 없음(중간), 목록에 있으나 도면 없음(높음), 도면명 불일치(중간)
    D02  축척 누락 또는 목록과 축척 상이(낮음)
    D03  개정번호 목록과 상이(중간)
    D04  동일 도면번호가 여러 페이지에 중복(중간)
스캔 페이지는 판정 불가 → 경고로 남겨 이미지 확인을 요청한다.
"""
import argparse
import os
import re
import sys
from collections import defaultdict

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from _numparse import make_finding, read_json, write_json  # noqa: E402


def norm_title(t):
    """도면명 비교용 정규화: 공백·구분기호 제거, 대문자."""
    if not t:
        return ""
    return re.sub(r"[\s\-_·.,()\[\]]", "", t).upper()


def loc_of(p):
    return "%s p.%d" % (p["file"], p["page"])


def loc_list(e):
    s = "%s p.%d 도면목록" % (e.get("file", "?"), e.get("page", 0))
    if e.get("no"):
        s += " 번호 %s" % e["no"]
    return s


def main():
    ap = argparse.ArgumentParser(description="도면 세트 정합성 검사(D01~D04)")
    ap.add_argument("dwg_json", help="extract_drawings.py 출력 JSON")
    ap.add_argument("--out", required=True, help="결과 findings JSON 경로")
    args = ap.parse_args()

    if not os.path.isfile(args.dwg_json):
        sys.stderr.write("오류: 파일을 찾을 수 없습니다: %s\n" % args.dwg_json)
        sys.exit(2)
    try:
        dwg = read_json(args.dwg_json)
        pages = dwg["pages"]
    except Exception as e:
        sys.stderr.write("오류: 도면 JSON을 읽을 수 없습니다(%s): %s\n" % (args.dwg_json, e))
        sys.exit(2)

    dlist = dwg.get("drawing_list", [])
    findings = []
    warnings = list(dwg.get("warnings", []))  # 추출 단계 경고를 그대로 전달

    # 판정 대상: 스캔 아님 + 도면번호 있음
    scanned = [p for p in pages if p.get("scanned")]
    drawings = [p for p in pages if not p.get("scanned") and p.get("dwg_no")]
    for p in scanned:
        warnings.append("%s: 스캔 페이지라 D01~D04 자동 검사 불가 — 이미지로 표제란 확인 필요" % loc_of(p))

    # --- D04 중복 도면번호 -------------------------------------------------
    by_no = defaultdict(list)
    for p in drawings:
        by_no[p["dwg_no"]].append(p)
    for no, ps in sorted(by_no.items()):
        if len(ps) > 1:
            titles = ", ".join("%s(%s)" % (loc_of(p), p.get("title") or "도면명 미상") for p in ps)
            findings.append(make_finding(
                "D04", "중간", ", ".join(loc_of(p) for p in ps),
                "도면번호 %s 가 %d개 페이지에 중복 사용됨" % (no, len(ps)),
                titles,
                "중복 페이지가 동일 도면의 사본인지, 번호 오기인지 확인 후 도면번호 정정"))

    # --- 목록 대조 ---------------------------------------------------------
    list_map = {}
    for e in dlist:
        if e["dwg_no"] in list_map:
            findings.append(make_finding(
                "D04", "중간", loc_list(e),
                "도면목록에 도면번호 %s 가 중복 기재됨" % e["dwg_no"],
                "%s / %s" % (list_map[e["dwg_no"]].get("raw"), e.get("raw")),
                "도면목록 정정"))
            continue
        list_map[e["dwg_no"]] = e

    if not dlist:
        warnings.append("도면목록이 없어 D01(목록 대조)·D03 및 D02의 목록 비교를 생략함 — 축척 누락만 검사")

    for p in drawings:
        if p.get("is_list"):
            continue  # 목록 페이지 자신은 목록 대조 제외
        loc = loc_of(p)
        e = list_map.get(p["dwg_no"]) if dlist else None
        if dlist and e is None:
            findings.append(make_finding(
                "D01", "중간", loc,
                "도면 %s(%s)가 도면목록에 없음" % (p["dwg_no"], p.get("title") or "도면명 미상"),
                "표제란: 도면번호 %s / 도면명 %s" % (p["dwg_no"], p.get("title")),
                "도면목록에 추가하거나 도면번호 오기 여부 확인"))
        if e is not None:
            if e.get("title") and p.get("title") and norm_title(e["title"]) != norm_title(p["title"]):
                findings.append(make_finding(
                    "D01", "중간", loc,
                    "도면명 불일치: %s" % p["dwg_no"],
                    "표제란 '%s' ↔ 도면목록 '%s'(%s)" % (p["title"], e["title"], loc_list(e)),
                    "도면목록 또는 표제란 도면명을 통일"))
            if e.get("rev") is not None and p.get("rev") is not None and str(e["rev"]) != str(p["rev"]):
                findings.append(make_finding(
                    "D03", "중간", loc,
                    "개정번호 불일치: %s" % p["dwg_no"],
                    "표제란 REV %s ↔ 도면목록 REV %s(%s)" % (p["rev"], e["rev"], loc_list(e)),
                    "최신 개정본 여부 확인 후 목록/표제란 개정번호 일치시킴"))
            elif e.get("rev") is not None and p.get("rev") is None:
                warnings.append("%s: 표제란 개정번호를 인식하지 못해 목록(REV %s)과 비교 불가" % (loc, e["rev"]))
        # D02 축척
        if not p.get("scale_norm"):
            findings.append(make_finding(
                "D02", "낮음", loc,
                "축척 표기 누락(또는 인식 불가): %s" % p["dwg_no"],
                "표제란 텍스트에서 축척(1/100, S=1/50, NONE 등) 미검출" +
                (" / 목록 축척 %s" % e["scale"] if e is not None and e.get("scale") else ""),
                "표제란 축척 기입 여부 확인(상세도·NONE 포함)"))
        elif e is not None and e.get("scale_norm") and e["scale_norm"] != p["scale_norm"]:
            findings.append(make_finding(
                "D02", "낮음", loc,
                "축척 불일치: %s" % p["dwg_no"],
                "표제란 %s ↔ 도면목록 %s" % (p.get("scale"), e.get("scale")),
                "실제 도면 축척 확인 후 목록/표제란 정정"))

    # 목록에는 있으나 도면이 없는 경우 (높음)
    present = set(by_no)
    for no, e in list_map.items():
        if no not in present:
            sev = "높음"
            note = ""
            if scanned:
                sev = "중간"
                note = " (스캔 페이지 %d개가 있어 그중 하나일 수 있음)" % len(scanned)
            findings.append(make_finding(
                "D01", sev, loc_list(e),
                "도면목록의 %s(%s) 도면이 PDF에 없음%s" % (no, e.get("title"), note),
                "목록 행: %s" % e.get("raw"),
                "누락 도면 수령 요청 또는 목록 정정"))

    out = {"source": args.dwg_json, "findings": findings, "warnings": warnings}
    write_json(out, args.out)
    print("도면 검사 완료: 지적 %d건, 경고 %d건 → %s" % (len(findings), len(warnings), args.out))


if __name__ == "__main__":
    main()
