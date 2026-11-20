#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
여러 findings JSON 을 병합하여 한국어 마크다운 검토 보고서를 만든다.

사용법:
    python3 render_report.py findings_boq.json findings_dwg.json findings_x.json \
        [--title "OO공사 내역서·도면 검토"] --out report.md
    python3 render_report.py "$WORK"            # 폴더 → 그 안의 findings*.json
    python3 render_report.py "$WORK/findings_*.json"   # 글로브 패턴(따옴표 가능, PowerShell 용)

- 인자는 파일·폴더·글로브 패턴 모두 가능(스크립트가 직접 펼치므로 셸 글로브가 없는 PowerShell 에서도 동작).
- 보고서 구성: 요약 → 지적 사항(표) → 검토 한계 및 경고.
- 정렬: 심각도(높음 > 중간 > 낮음) → 위치(파일·페이지·행, 숫자는 자연 정렬). E00(검사 실패)은 맨 앞.
- E00 이 있으면 요약에 "검사 실패 N건 — 결과 신뢰 불가" 를 강조한다.
- 표 셀의 '|' 와 줄바꿈은 이스케이프한다. 내용의 "연쇄 가능" 표기는 그대로 둔다.
"""
import argparse
import datetime
import glob
import os
import re
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from _numparse import read_json  # noqa: E402

SEV_ORDER = {"높음": 0, "중간": 1, "낮음": 2}
DISCLAIMER = "본 보고서는 검토 보조 자료이며 공인 검토를 대체하지 않습니다."


def esc(s):
    """마크다운 표 셀 이스케이프: 파이프, 줄바꿈."""
    s = "" if s is None else str(s)
    s = s.replace("\\", "\\\\").replace("|", "\\|")
    s = re.sub(r"\r\n|\r|\n", "<br>", s)
    return s.strip()


def natural_key(s):
    return [int(t) if t.isdigit() else t for t in re.split(r"(\d+)", str(s or ""))]


def expand_inputs(args, out_path=None):
    """파일·폴더·글로브 패턴 인자를 findings JSON 파일 목록으로 펼친다(순서 유지, 중복 제거)."""
    files, missing = [], []
    out_abs = os.path.abspath(out_path) if out_path else None
    for a in args:
        if os.path.isdir(a):
            hits = sorted(glob.glob(os.path.join(a, "findings*.json")))
            if not hits:
                missing.append("%s (폴더에 findings*.json 없음)" % a)
        elif glob.has_magic(a):
            hits = sorted(glob.glob(a))
            if not hits:
                missing.append("%s (패턴과 일치하는 파일 없음)" % a)
        elif os.path.isfile(a):
            hits = [a]
        else:
            hits = []
            missing.append(a)
        for h in hits:
            if os.path.isfile(h) and os.path.abspath(h) != out_abs and os.path.abspath(h) not in \
                    [os.path.abspath(f) for f in files]:
                files.append(h)
    return files, missing


def main():
    for s in (sys.stdout, sys.stderr):
        try:
            s.reconfigure(encoding="utf-8", errors="replace")
        except Exception:  # noqa: BLE001
            pass
    ap = argparse.ArgumentParser(description="findings JSON → 한국어 마크다운 보고서")
    ap.add_argument("findings", nargs="+", help="findings JSON 파일·폴더·글로브 패턴")
    ap.add_argument("--title", default="내역서·도면 검토 보고서", help="보고서 제목")
    ap.add_argument("--out", required=True, help="출력 마크다운 경로")
    args = ap.parse_args()

    files, missing = expand_inputs(args.findings, args.out)
    if missing:
        sys.stderr.write("오류: 입력을 찾을 수 없습니다: %s\n" % "; ".join(missing))
        sys.exit(2)
    if not files:
        sys.stderr.write("오류: 읽을 findings JSON 이 없습니다.\n")
        sys.exit(2)

    all_f, all_w, sources = [], [], []
    for p in files:
        try:
            d = read_json(p)
        except Exception as e:  # noqa: BLE001
            sys.stderr.write("오류: JSON을 읽을 수 없습니다(%s): %s\n" % (p, e))
            sys.exit(2)
        if not isinstance(d, dict) or "findings" not in d:
            all_w.append("%s: findings JSON 형식이 아니어서 건너뜀" % os.path.basename(p))
            continue
        src = str(d.get("source") or os.path.basename(p))
        if os.sep in src or "/" in src:
            src = os.path.basename(src.replace("\\", "/"))
        sources.append((os.path.basename(p), src))
        for f in d.get("findings", []):
            if f.get("severity") not in SEV_ORDER:
                all_w.append("%s: 알 수 없는 심각도 '%s' → '낮음'으로 처리 (%s)"
                             % (os.path.basename(p), f.get("severity"), f.get("code")))
                f = dict(f, severity="낮음")
            all_f.append(f)
        for w in d.get("warnings", []):
            all_w.append(w if isinstance(w, str) else str(w))

    # 동일 경고 중복 제거(순서 유지): check_drawings 가 추출 경고를 전달하는 경우 등
    all_w = list(dict.fromkeys(all_w))
    all_f.sort(key=lambda f: (0 if f.get("code") == "E00" else 1, SEV_ORDER[f["severity"]],
                              natural_key(f.get("location")), f.get("code", "")))

    cnt = {k: sum(1 for f in all_f if f["severity"] == k) for k in SEV_ORDER}
    n_fail = sum(1 for f in all_f if f.get("code") == "E00")
    L = []
    L.append("# %s" % args.title)
    L.append("")
    L.append("작성일시: %s  " % datetime.datetime.now().strftime("%Y-%m-%d %H:%M"))
    L.append(DISCLAIMER + "  ")
    L.append("각 지적 사항은 자동 추출·대조 결과이므로 원본 문서로 확인 후 확정해야 합니다.")
    L.append("")
    L.append("## 요약")
    L.append("")
    if n_fail:
        L.append("> **검사 실패 %d건 — 결과 신뢰 불가.** 실패한 검사(E00)의 범위는 지적이 없더라도 "
                 "오류가 없다는 뜻이 아닙니다. 원인을 해결하고 다시 실행하십시오." % n_fail)
        L.append("")
    L.append("- 검토 입력:")
    for fn, src in sources:
        L.append("  - `%s` (원본: %s)" % (fn, src))
    L.append("- 지적 사항: 총 %d건 (높음 %d · 중간 %d · 낮음 %d)" % (len(all_f), cnt["높음"], cnt["중간"], cnt["낮음"]))
    codes = {}
    for f in all_f:
        codes[f.get("code", "?")] = codes.get(f.get("code", "?"), 0) + 1
    if codes:
        L.append("- 코드별: " + ", ".join("%s %d건" % (k, v) for k, v in sorted(codes.items())))
    n_chain = sum(1 for f in all_f if "연쇄 가능" in " ".join(str(f.get(k) or "") for k in
                                                          ("content", "evidence", "recommendation")))
    if n_chain:
        L.append("- '연쇄 가능' 표시 %d건: 다른 지적의 결과일 수 있으므로 원인 지적과 묶어 검토" % n_chain)
    L.append("- 검토 한계 및 경고: %d건" % len(all_w))
    L.append("")
    L.append("## 지적 사항")
    L.append("")
    if not all_f:
        L.append("자동 검사에서 지적된 사항이 없습니다. 다만 아래 검토 한계 및 경고를 확인하십시오.")
    else:
        L.append("| 번호 | 심각도 | 코드 | 위치(페이지·행) | 내용 | 근거 | 권고 |")
        L.append("|---|---|---|---|---|---|---|")
        for i, f in enumerate(all_f, start=1):
            L.append("| %d | %s | %s | %s | %s | %s | %s |" % ((i,) + tuple(esc(f.get(k)) for k in (
                "severity", "code", "location", "content", "evidence", "recommendation"))))
    L.append("")
    L.append("## 검토 한계 및 경고")
    L.append("")
    if all_w:
        for w in all_w:
            L.append("- %s" % w.replace("\n", " "))
    else:
        L.append("- 스크립트 경고 없음")
    L.append("")

    with open(args.out, "w", encoding="utf-8") as fh:
        fh.write("\n".join(L))
    print("보고서 작성 완료: 입력 %d개, 지적 %d건(높음 %d/중간 %d/낮음 %d, 검사 실패 %d), 경고 %d건 → %s"
          % (len(files), len(all_f), cnt["높음"], cnt["중간"], cnt["낮음"], n_fail, len(all_w), args.out))


if __name__ == "__main__":
    main()
