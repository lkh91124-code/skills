#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
도면 PDF에서 표제란(도면번호/도면명/축척/개정), 도면목록, 규격 토큰을 추출한다.

사용법:
    python3 extract_drawings.py A.pdf [B.pdf ...] [--title-region 0.6,0.75,1,1] --out dwg.json

출력(JSON):
    {"source_files": [...],
     "pages": [{file, page, dwg_no, title, scale, scale_norm, rev, is_list, scanned,
                spec_tokens:[{cat, token, raw}], title_text_raw}],
     "drawing_list": [{file, page, no, dwg_no, title, scale, scale_norm, rev, raw}],
     "warnings": [...]}

주의:
- 표제란은 기본적으로 페이지 우하단(가로 60~100%, 세로 75~100%)으로 가정한다.
  도면 양식이 다르면 --title-region 으로 조정한다(좌상단 원점, 비율 값).
- 텍스트가 30자 미만인 페이지는 스캔(이미지) 도면으로 보고 scanned:true 로 표시한다.
  이런 페이지는 이미지를 직접 확인해야 한다(pdftoppm 으로 렌더링).
- 이 모듈의 extract_spec_tokens() 는 cross_check.py 에서도 재사용한다.
"""
import argparse
import os
import re
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

try:
    import pdfplumber
except ImportError:  # pragma: no cover
    sys.stderr.write("오류: pdfplumber 가 설치되어 있지 않습니다 (pip install pdfplumber).\n")
    sys.exit(2)

from _numparse import write_json  # noqa: E402  (공용 헬퍼)

# ---------------------------------------------------------------------------
# 정규식 정의
# ---------------------------------------------------------------------------

# 도면번호: A-101, S-201, AR-001, M-01-001, 기-01, E-1001 등
# 앞뒤가 영숫자이면 제외(예: H-300x300 의 'H-300' 은 뒤에 'x'가 붙어 제외됨)
DWG_NO_RE = re.compile(
    r"(?<![0-9A-Za-z가-힣\-])((?:[A-Z]{1,3}|[가-힣]{1,2})\s?-\s?\d{1,4}(?:-\d{1,4})?[A-Z]?)(?![0-9A-Za-z\-])"
)

# 축척: 1/100, 1:100, S=1/50, A1:1/100, NONE, N.S, NTS
SCALE_RE = re.compile(
    r"(?:(A[0-4])\s*[:=]?\s*)?(?:S\s*=\s*)?1\s*[/:]\s*(\d{1,5})(?!\d)", re.I
)
SCALE_NONE_RE = re.compile(r"(?<![A-Za-z])(NONE|N\.\s?S\.?|NTS|N\.T\.S\.?|없음)(?![A-Za-z])", re.I)

# 개정: REV 2, REV.A, 개정 1, REVISION: 3
REV_LABEL_RE = re.compile(r"(?:REV(?:ISION)?|개정)\s*(?:NO)?\s*[.:#]?\s*([0-9]{1,2}|[A-Z])(?![0-9A-Za-z])", re.I)
# 개정 이력표 행: "2  2024.03.15  설계변경 ..."
REV_ROW_RE = re.compile(r"^\s*([0-9]{1,2}|[A-Z])\s+(?:\d{2,4})[.\-/]\d{1,2}[.\-/]\d{1,2}")

# 표제란 라벨 (긴 것부터 검사해야 "DRAWING NO" 와 "DRAWING TITLE" 이 구분됨)
LABELS = {
    "dwg_no": [r"도\s*면\s*번\s*호", r"DRAWING\s*NO\.?", r"DWG\.?\s*NO\.?", r"SHEET\s*NO\.?"],
    "title": [r"도\s*면\s*명", r"DRAWING\s*TITLE", r"(?<![A-Z])TITLE"],
    "scale": [r"축\s*척", r"SCALE"],
    "rev": [r"개\s*정(?!\s*이력)", r"REV(?:ISION)?\.?(?:\s*NO\.?)?"],
}
ALL_LABEL_RE = re.compile("|".join(p for ps in LABELS.values() for p in ps), re.I)

LIST_KEYWORDS_RE = re.compile(r"도\s*면\s*목\s*록|DRAWING\s*LIST|DRAWING\s*INDEX|LIST\s*OF\s*DRAWINGS", re.I)

# ---------------------------------------------------------------------------
# 규격 토큰
# ---------------------------------------------------------------------------
REBAR_DIAS = "10|13|16|19|22|25|29|32|35|38|41|51"
SPEC_PATTERNS = [
    # 강재 단면: H-300x300x10x15 (먼저 추출 후 텍스트에서 제거하여 치수와 중복 방지)
    ("steel", re.compile(r"(?<![A-Za-z])([HI]|□|C|L)\s*-\s*(\d{2,4}(?:\s*[xX×*]\s*\d{1,4}(?:\.\d)?){1,3})")),
    # 콘크리트 강도: 24MPa, fck=24, 레미콘 25-24-150
    ("concrete", re.compile(r"(\d{2,3})\s*MPa", re.I)),
    ("concrete", re.compile(r"f\s*ck\s*[=:]?\s*(\d{2,3})", re.I)),
    ("concrete", re.compile(r"(?<![\d\-])(?:19|20|25|40)\s*-\s*(\d{2})\s*-\s*(?:\d{2,3})(?![\d\-])")),
    # 철근 강종: SD400, SD500W, SHD, HD, UHD
    ("rebar_grade", re.compile(r"(?<![A-Za-z])(SD\s*\d{3}W?|SHD|UHD|HD)(?=\s*\d|\b|[^A-Za-z])")),
    # 철근 호칭: D10~D51 (HD13, SHD16, SD400 D13 등)
    ("rebar_dia", re.compile(r"(?<![A-Za-z])(?:SHD|UHD|HD|SD|D)\s?(" + REBAR_DIAS + r")(?!\d)")),
    # 두께: THK=200, THK200, T=200, 두께 200
    ("thk", re.compile(r"(?<![A-Za-z])(?:THK\s*[=.:]?\s*|T\s*=\s*|두께\s*[=:]?\s*)(\d{1,4}(?:\.\d)?)", re.I)),
    # 배근 간격: @200
    ("spacing", re.compile(r"@\s*(\d{2,4})")),
    # 관경: Φ100, ø50, DN100
    ("pipe", re.compile(r"(?:[ΦφøØ]|DN)\s*(\d{1,4})")),
    # KS 규격: KS D 3504, KS F2405
    ("ks", re.compile(r"KS\s*([A-Z])\s*-?\s*(\d{3,4})")),
    # 단면/부재 치수: 300x600, 300X600, 300*600
    ("size", re.compile(r"(?<![\d.\-])(\d{2,4})\s*[xX×*]\s*(\d{2,4})(?:\s*[xX×*]\s*(\d{2,4}))?(?![\d.])")),
]


def _norm_spec(cat, m):
    """정규식 매치를 비교 가능한 정규화 토큰 문자열로 변환."""
    if cat == "steel":
        dims = re.split(r"\s*[xX×*]\s*", m.group(2))
        return "%s-%s" % (m.group(1).upper(), "x".join(d.strip() for d in dims))
    if cat == "concrete":
        return "fck%s" % int(m.group(1))
    if cat == "rebar_grade":
        return re.sub(r"\s+", "", m.group(1)).upper()
    if cat == "rebar_dia":
        return "D%s" % m.group(1)
    if cat == "thk":
        v = m.group(1)
        return "THK%s" % (v[:-2] if v.endswith(".0") else v)
    if cat == "spacing":
        return "@%s" % int(m.group(1))
    if cat == "pipe":
        return "Φ%s" % int(m.group(1))
    if cat == "ks":
        return "KS %s %s" % (m.group(1).upper(), m.group(2))
    if cat == "size":
        parts = [g for g in m.groups() if g]
        return "x".join(str(int(p)) for p in parts)
    return m.group(0)


def extract_spec_tokens(text):
    """텍스트에서 규격 토큰 목록 [{cat, token, raw}] 을 추출(중복 제거)."""
    if not text:
        return []
    out, seen = [], set()
    work = text
    for cat, rx in SPEC_PATTERNS:
        for m in rx.finditer(work):
            tok = _norm_spec(cat, m)
            key = (cat, tok)
            if key not in seen:
                seen.add(key)
                out.append({"cat": cat, "token": tok, "raw": m.group(0).strip()})
        if cat == "steel":
            # 강재 단면 문자열은 이후 'size' 등에 중복 매치되지 않도록 공백으로 치환
            work = rx.sub(lambda mm: " " * len(mm.group(0)), work)
    return out


# ---------------------------------------------------------------------------
# 표제란 파싱
# ---------------------------------------------------------------------------

def norm_dwg_no(s):
    """도면번호 정규화: 공백 제거, 대시 통일, 대문자."""
    if not s:
        return None
    s = re.sub(r"[‐‑‒–—―－]", "-", s)
    s = re.sub(r"\s+", "", s).upper()
    return s or None


def parse_scale(s):
    """축척 문자열 -> (raw, norm). norm 예: '1/100', 'A1=1/100;A3=1/200', 'NONE'."""
    if not s:
        return None, None
    found = []
    for m in SCALE_RE.finditer(s):
        paper = (m.group(1) or "").upper()
        found.append(((paper + "=") if paper else "") + "1/" + str(int(m.group(2))))
    if found:
        # 중복 제거(순서 유지)
        uniq = list(dict.fromkeys(found))
        raw = s.strip()
        return raw, ";".join(uniq)
    m = SCALE_NONE_RE.search(s)
    if m:
        return m.group(0), "NONE"
    return None, None


def norm_rev(r):
    if r is None:
        return None
    r = str(r).strip().upper()
    if r.isdigit():
        return str(int(r))
    return r or None


def _segments_by_label(lines):
    """각 줄에서 라벨 위치로 분할 → [(field, value_on_same_line, line_index)]."""
    segs = []
    for i, line in enumerate(lines):
        hits = []
        for field, pats in LABELS.items():
            for p in pats:
                for m in re.finditer(p, line, re.I):
                    hits.append((m.start(), m.end(), field))
        if not hits:
            continue
        # 겹치는 라벨은 시작이 빠르고 긴 것 우선
        hits.sort(key=lambda h: (h[0], -(h[1] - h[0])))
        clean, last_end = [], -1
        for h in hits:
            if h[0] >= last_end:
                clean.append(h)
                last_end = h[1]
        for k, (s, e, field) in enumerate(clean):
            nxt = clean[k + 1][0] if k + 1 < len(clean) else len(line)
            val = line[e:nxt].strip(" :：.|/-\t")
            segs.append((field, val, i))
    return segs


def parse_title_block(text):
    """표제란 텍스트에서 도면번호/도면명/축척/개정을 추출."""
    res = {"dwg_no": None, "title": None, "scale": None, "scale_norm": None, "rev": None}
    if not text:
        return res
    lines = [l for l in text.splitlines() if l.strip()]
    segs = _segments_by_label(lines)

    def next_line_value(i):
        # 라벨 다음 줄이 다른 라벨 없이 값만 있으면 그 값 사용
        if i + 1 < len(lines) and not ALL_LABEL_RE.search(lines[i + 1]):
            return lines[i + 1].strip()
        return ""

    for field, val, i in segs:
        if not val:
            val = next_line_value(i)
        if not val:
            continue
        if field == "dwg_no" and not res["dwg_no"]:
            m = DWG_NO_RE.search(val)
            if m:
                res["dwg_no"] = norm_dwg_no(m.group(1))
        elif field == "title" and not res["title"]:
            res["title"] = re.sub(r"\s+", " ", val).strip()
        elif field == "scale" and not res["scale"]:
            raw, nrm = parse_scale(val)
            if nrm:
                res["scale"], res["scale_norm"] = raw, nrm
        elif field == "rev" and res["rev"] is None:
            m = re.match(r"\s*([0-9]{1,2}|[A-Z])(?![0-9A-Za-z])", val)
            if m:
                res["rev"] = norm_rev(m.group(1))

    # 라벨 기반으로 못 찾은 경우의 대체 규칙
    if not res["dwg_no"]:
        m = DWG_NO_RE.search(text)
        if m:
            res["dwg_no"] = norm_dwg_no(m.group(1))
            res["_dwg_no_guess"] = True
    if not res["scale_norm"]:
        raw, nrm = parse_scale(text)
        if nrm:
            res["scale"], res["scale_norm"] = raw, nrm
    if res["rev"] is None:
        m = REV_LABEL_RE.search(text)
        if m:
            res["rev"] = norm_rev(m.group(1))
        else:
            # 개정 이력표: 가장 큰 개정번호
            revs = [norm_rev(mm.group(1)) for l in lines for mm in [REV_ROW_RE.match(l)] if mm]
            if revs:
                nums = [r for r in revs if r.isdigit()]
                res["rev"] = str(max(int(r) for r in nums)) if nums else sorted(revs)[-1]
    return res


# ---------------------------------------------------------------------------
# 도면목록 파싱
# ---------------------------------------------------------------------------
HEADER_MAP = [
    ("no", re.compile(r"^(번호|NO\.?|순번|S/N)$", re.I)),
    ("dwg_no", re.compile(r"도면\s*번호|DWG\.?\s*NO|DRAWING\s*NO|SHEET\s*NO", re.I)),
    ("title", re.compile(r"도\s*면\s*명|TITLE|DESCRIPTION", re.I)),
    ("scale", re.compile(r"축\s*척|SCALE", re.I)),
    ("rev", re.compile(r"개\s*정|REV", re.I)),
]


def _map_header(row):
    cols = {}
    for idx, cell in enumerate(row):
        c = (cell or "").replace("\n", " ").strip()
        for key, rx in HEADER_MAP:
            if key not in cols and rx.search(c):
                cols[key] = idx
                break
    return cols


def parse_list_tables(page):
    """pdfplumber extract_tables 로 도면목록 행을 추출."""
    entries = []
    try:
        tables = page.extract_tables()
    except Exception:
        tables = []
    for tb in tables or []:
        cols = None
        for row in tb:
            if not row:
                continue
            if cols is None:
                h = _map_header(row)
                if "dwg_no" in h:
                    cols = h
                continue
            cell = lambda k: ((row[cols[k]] or "").replace("\n", " ").strip() if k in cols and cols[k] < len(row) else "")
            m = DWG_NO_RE.search(cell("dwg_no"))
            if not m:
                continue
            sraw, snorm = parse_scale(cell("scale"))
            rv = cell("rev")
            entries.append({
                "no": cell("no") or None,
                "dwg_no": norm_dwg_no(m.group(1)),
                "title": re.sub(r"\s+", " ", cell("title")) or None,
                "scale": sraw, "scale_norm": snorm,
                "rev": norm_rev(rv) if rv else None,
                "raw": " | ".join((c or "").replace("\n", " ") for c in row),
                "method": "table",
            })
    return entries


def parse_list_text(text):
    """표 인식 실패 시 텍스트 줄 단위로 도면목록 추정: [번호] 도면번호 도면명 [축척] [개정]."""
    entries = []
    for line in (text or "").splitlines():
        m = DWG_NO_RE.search(line)
        if not m:
            continue
        pre = line[:m.start()].strip()
        rest = line[m.end():].strip()
        sraw, snorm = None, None
        sm = SCALE_RE.search(rest) or SCALE_NONE_RE.search(rest)
        rev = None
        title = rest
        if sm:
            title = rest[:sm.start()].strip()
            after = rest[sm.end():].strip()
            # 축척 뒤 추가 축척(A3:1/200 등)은 함께 포함
            sraw, snorm = parse_scale(rest[sm.start():])
            # raw 는 축척 표기 부분만 남김(뒤의 개정번호 제외)
            spans = [mm.end() for mm in SCALE_RE.finditer(rest[sm.start():])]
            sraw = rest[sm.start():sm.start() + (spans[-1] if spans else sm.end() - sm.start())].strip()
            tail = SCALE_RE.sub("", after).strip()
            rm = re.match(r"^([0-9]{1,2}|[A-Z])$", tail.split()[-1]) if tail.split() else None
            if rm:
                rev = norm_rev(rm.group(1))
        else:
            parts = rest.split()
            if parts and re.fullmatch(r"[0-9]{1,2}", parts[-1]):
                rev = norm_rev(parts[-1])
                title = " ".join(parts[:-1])
        entries.append({
            "no": pre if re.fullmatch(r"\d{1,4}", pre) else None,
            "dwg_no": norm_dwg_no(m.group(1)),
            "title": re.sub(r"\s+", " ", title) or None,
            "scale": sraw, "scale_norm": snorm, "rev": rev,
            "raw": line.strip(), "method": "text",
        })
    return entries


# ---------------------------------------------------------------------------
# 메인 처리
# ---------------------------------------------------------------------------

def crop_text(page, region):
    x0, y0, x1, y1 = region
    W, H = page.width, page.height
    bbox = (W * x0, H * y0, W * x1, H * y1)
    try:
        return page.within_bbox(bbox).extract_text() or ""
    except Exception:
        return ""


def process_pdf(path, region, warnings):
    pages_out, list_out = [], []
    fname = os.path.basename(path)
    with pdfplumber.open(path) as pdf:
        for pno, page in enumerate(pdf.pages, start=1):
            loc = "%s p.%d" % (fname, pno)
            try:
                full = page.extract_text() or ""
            except Exception as e:
                warnings.append("%s: 텍스트 추출 실패(%s) — 이미지로 확인 필요" % (loc, e))
                full = ""
            scanned = len(re.sub(r"\s", "", full)) < 30
            rec = {"file": fname, "page": pno, "dwg_no": None, "title": None, "scale": None,
                   "scale_norm": None, "rev": None, "is_list": False, "scanned": scanned,
                   "spec_tokens": [], "title_text_raw": ""}
            if scanned:
                warnings.append("%s: 텍스트가 거의 없음(스캔 도면 추정) — pdftoppm 으로 이미지 렌더링 후 육안 확인 필요" % loc)
                pages_out.append(rec)
                continue

            ttext = crop_text(page, region)
            rec["title_text_raw"] = ttext
            tb = parse_title_block(ttext)
            if not tb.get("dwg_no"):
                # 표제란 영역에서 실패 → 전체 텍스트의 라벨 근처에서 재시도
                tb2 = parse_title_block(full)
                for k in ("dwg_no", "title", "scale", "scale_norm", "rev"):
                    if not tb.get(k) and tb2.get(k) and not (k == "dwg_no" and tb2.get("_dwg_no_guess")):
                        tb[k] = tb2[k]
                if tb.get("dwg_no"):
                    warnings.append("%s: 표제란 영역에서 도면번호를 못 찾아 페이지 전체 텍스트에서 추정함(%s)" % (loc, tb["dwg_no"]))
            elif tb.get("_dwg_no_guess"):
                warnings.append("%s: '도면번호' 라벨 없이 패턴으로 추정한 도면번호(%s) — 확인 권장" % (loc, tb["dwg_no"]))
            for k in ("dwg_no", "title", "scale", "scale_norm", "rev"):
                rec[k] = tb.get(k)
            if not rec["dwg_no"]:
                warnings.append("%s: 도면번호 미검출 — 표제란 위치(--title-region) 확인 필요" % loc)

            # 도면목록 페이지 판정 및 행 파싱
            if LIST_KEYWORDS_RE.search(full):
                ents = parse_list_tables(page)
                if not ents:
                    # 표제란 텍스트 줄은 제외하고 텍스트 파싱
                    body = full.replace(ttext, "") if ttext else full
                    ents = parse_list_text(body)
                    if ents:
                        warnings.append("%s: 도면목록 표 인식 실패 → 텍스트 줄 기반으로 %d행 추정(정확도 낮음)" % (loc, len(ents)))
                # 표제란의 자기 도면번호만 잡힌 경우 목록으로 보지 않음
                ents = [e for e in ents if e["dwg_no"] != rec["dwg_no"] or len(ents) > 1]
                if len(ents) >= 2:
                    rec["is_list"] = True
                    for e in ents:
                        e.update({"file": fname, "page": pno})
                    list_out.extend(ents)
                else:
                    warnings.append("%s: 도면목록 키워드는 있으나 목록 행을 인식하지 못함" % loc)

            # 규격 토큰(도면목록 페이지는 제외 — 도면명 등에서 오탐 방지)
            if not rec["is_list"]:
                rec["spec_tokens"] = extract_spec_tokens(full)
            pages_out.append(rec)
    return pages_out, list_out


def parse_region(s):
    try:
        vals = [float(v) for v in s.split(",")]
        assert len(vals) == 4 and all(0 <= v <= 1 for v in vals) and vals[0] < vals[2] and vals[1] < vals[3]
        return vals
    except Exception:
        sys.stderr.write("오류: --title-region 형식이 잘못되었습니다. 예) 0.6,0.75,1,1 (0~1 비율, x0<x1, y0<y1)\n")
        sys.exit(2)


def main():
    ap = argparse.ArgumentParser(description="도면 PDF 표제란·도면목록·규격 토큰 추출")
    ap.add_argument("pdfs", nargs="+", help="도면 PDF 파일(여러 개 가능)")
    ap.add_argument("--title-region", default="0.6,0.75,1,1",
                    help="표제란 영역 x0,y0,x1,y1 (페이지 대비 비율, 좌상단 원점). 기본 0.6,0.75,1,1")
    ap.add_argument("--out", required=True, help="출력 JSON 경로 (예: dwg.json)")
    args = ap.parse_args()

    region = parse_region(args.title_region)
    for p in args.pdfs:
        if not os.path.isfile(p):
            sys.stderr.write("오류: 파일을 찾을 수 없습니다: %s\n" % p)
            sys.exit(2)

    warnings, pages, dlist = [], [], []
    for p in args.pdfs:
        try:
            po, lo = process_pdf(p, region, warnings)
        except Exception as e:
            sys.stderr.write("오류: PDF를 읽을 수 없습니다(%s): %s\n" % (p, e))
            sys.exit(2)
        pages.extend(po)
        dlist.extend(lo)

    if not dlist:
        warnings.append("도면목록 페이지를 찾지 못함 — 목록 대조(D01~D03)는 생략됨")
    out = {"source_files": [os.path.basename(p) for p in args.pdfs], "pages": pages,
           "drawing_list": dlist, "warnings": warnings}
    write_json(out, args.out)
    n_scan = sum(1 for p in pages if p["scanned"])
    print("도면 추출 완료: %d페이지(스캔 %d), 도면목록 %d행, 경고 %d건 → %s"
          % (len(pages), n_scan, len(dlist), len(warnings), args.out))


if __name__ == "__main__":
    main()
