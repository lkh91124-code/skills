#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
여러 findings JSON 을 병합하여 한국어 마크다운 검토 보고서를 만든다.

사용법:
    python3 render_report.py findings_boq.json findings_dwg.json findings_x.json \
        [--title "OO공사 내역서·도면 검토"] --out report.md

- 정렬: 심각도(높음 > 중간 > 낮음) → 위치(파일·페이지·행, 숫자는 자연 정렬)
- 표 셀의 '|' 와 줄바꿈은 이스케이프한다.
- 모든 경고는 '경고 / 검토 한계' 섹션에 그대로 나열한다.
"""
import argparse
import datetime
import os
import re
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from _numparse import read_json  # noqa: E402

SEV_ORDER = {"높음": 0, "중간": 1, "낮음": 2}


def esc(s):
    """마크다운 표 셀 이스케이프: 파이프, 줄바꿈."""
    s = "" if s is None else str(s)
    s = s.replace("\\", "\\\\").replace("|", "\\|")
    s = re.sub(r"\r\n|\r|\n", "<br>", s)
    return s.strip()


def natural_key(s):
    return [int(t) if t.isdigit() else t for t in re.split(r"(\d+)", str(s or ""))]


def main():
    ap = argparse.ArgumentParser(description="findings JSON → 한국어 마크다운 보고서")
    ap.add_argument("findings", nargs="+", help="findings JSON 파일들")
    ap.add_argument("--title", default="내역서·도면 검토 보고서", help="보고서 제목")
    ap.add_argument("--out", required=True, help="출력 마크다운 경로")
    args = ap.parse_args()

    all_f, all_w, sources = [], [], []
    for p in args.findings:
        if not os.path.isfile(p):
            sys.stderr.write("오류: 파일을 찾을 수 없습니다: %s\n" % p)
            sys.exit(2)
        try:
            d = read_json(p)
        except Exception as e:
            sys.stderr.write("오류: JSON을 읽을 수 없습니다(%s): %s\n" % (p, e))
            sys.exit(2)
        src = d.get("source") or os.path.basename(p)
        sources.append((os.path.basename(p), src))
        for f in d.get("findings", []):
            if f.get("severity") not in SEV_ORDER:
                all_w.append("%s: 알 수 없는 심각도 '%s' → '낮음'으로 처리 (%s)" % (os.path.basename(p), f.get("severity"), f.get("code")))
                f = dict(f, severity="낮음")
            all_f.append(f)
        for w in d.get("warnings", []):
            all_w.append(w if isinstance(w, str) else str(w))

    # 동일 경고 중복 제거(순서 유지): check_drawings 가 추출 경고를 전달하는 경우 등
    all_w = list(dict.fromkeys(all_w))
    all_f.sort(key=lambda f: (SEV_ORDER[f["severity"]], natural_key(f.get("location")), f.get("code", "")))

    cnt = {k: sum(1 for f in all_f if f["severity"] == k) for k in SEV_ORDER}
    L = []
    L.append("# %s" % args.title)
    L.append("")
    L.append("작성일시: %s  " % datetime.datetime.now().strftime("%Y-%m-%d %H:%M"))
    L.append("본 보고서는 자동 추출·대조 결과이며, 각 지적사항은 원본 문서로 확인 후 확정해야 합니다.")
    L.append("")
    L.append("## 요약")
    L.append("")
    L.append("- 검토 입력:")
    for fn, src in sources:
        L.append("  - `%s` (원본: %s)" % (fn, src))
    L.append("- 지적사항: 총 %d건 (높음 %d · 중간 %d · 낮음 %d)" % (len(all_f), cnt["높음"], cnt["중간"], cnt["낮음"]))
    codes = {}
    for f in all_f:
        codes[f.get("code", "?")] = codes.get(f.get("code", "?"), 0) + 1
    if codes:
        L.append("- 코드별: " + ", ".join("%s %d건" % (k, v) for k, v in sorted(codes.items())))
    L.append("- 경고(검토 한계): %d건" % len(all_w))
    L.append("")
    L.append("## 지적사항")
    L.append("")
    if not all_f:
        L.append("자동 검사에서 지적된 사항이 없습니다. 다만 아래 경고/검토 한계를 확인하십시오.")
    else:
        L.append("| 심각도 | 코드 | 위치(페이지·행) | 내용 | 근거 | 권고 |")
        L.append("|---|---|---|---|---|---|")
        for f in all_f:
            L.append("| %s | %s | %s | %s | %s | %s |" % tuple(esc(f.get(k)) for k in (
                "severity", "code", "location", "content", "evidence", "recommendation")))
    L.append("")
    L.append("## 경고 / 검토 한계")
    L.append("")
    if all_w:
        for w in all_w:
            L.append("- %s" % w.replace("\n", " "))
    else:
        L.append("- 없음")
    L.append("")

    with open(args.out, "w", encoding="utf-8") as fh:
        fh.write("\n".join(L))
    print("보고서 작성 완료: 지적 %d건(높음 %d/중간 %d/낮음 %d), 경고 %d건 → %s"
          % (len(all_f), cnt["높음"], cnt["중간"], cnt["낮음"], len(all_w), args.out))


if __name__ == "__main__":
    main()
