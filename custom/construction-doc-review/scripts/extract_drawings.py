#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
도면 PDF에서 표제란(도면번호/도면명/축척/개정), 도면목록, 규격 토큰을 추출한다.

사용법:
    python3 extract_drawings.py A.pdf [B.pdf ...] [--title-region 0.6,0.75,1,1] --out dwg.json

출력(JSON):
    {"source_files": [...],
     "pages": [{file, page, dwg_no, title, scale, scale_norm, rev, is_list, scanned,
                rev_suffix?, title_region, spec_tokens:[{cat, token, raw}], title_text_raw}],
     "drawing_list": [{file, page, no, dwg_no, title, scale, scale_norm, rev, raw}],
     "warnings": [...]}

주의:
- 표제란은 기본적으로 페이지 우하단(가로 60~100%, 세로 75~100%)으로 가정한다(--title-region 으로 변경).
  이 영역에서 '도면번호' 라벨을 못 찾으면 우상단·좌하단·좌상단·우측 띠·하단 띠·좌측 띠·상단 띠,
  마지막으로 페이지 전체 순으로 재시도하고, 어느 영역을 썼는지 경고에 남긴다(pages[].title_region).
- 도면번호는 라벨 옆 값·도면목록 '도면번호' 열에서 우선 읽고, 도면목록 번호를 어휘로 사용해 표제란과 맞춘다.
  라벨이 없을 때만 본문에서 엄격 패턴(D10·SD400·GL-1500 등 규격 표기 제외)으로 추정한다.
  S-201R1 처럼 개정 접미가 붙은 번호는 'S-201' 로 정규화하고 접미를 개정번호(rev_suffix)로 쓴다.
- 도면목록에서 해석하지 못한 비어 있지 않은 행/칸은 버리지 않고 경고로 남긴다.
- 텍스트가 30자 미만인 페이지는 스캔(이미지) 도면으로 보고 scanned:true 로 표시한다.
  이런 페이지는 이미지를 직접 확인해야 한다(pdftoppm 으로 렌더링).
- 이 모듈의 extract_spec_tokens() 는 cross_check.py 에서도 재사용한다.
"""
import argparse
import contextlib
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

# 도면번호: A-101, A101, E01, S-201R1, ARCH-101, A-101-01-2, 기-01, 기계설비-01 등
#   접두: 영문 대문자 1~6자 또는 한글 1~6자, 구분: 대시(선택, 한글 접두는 필수),
#   숫자 구간: 1~4자리 여러 개(대시 구분), 접미: 영문 1자(A-101A) 또는 개정 접미(R1, REV2).
# 앞뒤가 영숫자·한글·대시이면 제외(예: H-300x300 은 뒤에 'x' 가 붙어 제외, 24MPa 는 숫자로 시작해 제외).
_DWG_CORE = (
    r"(?:[A-Z]{1,6}(?:\s?-\s?|(?=\d))|[가-힣]{1,6}\s?-\s?)"   # 접두 + 구분
    r"\d{1,4}(?:\s?-\s?\d{1,4})*"                              # 숫자 구간
    r"(?:-?(?:R|REV\.?)\d{1,2}|[A-Z])?"                         # 접미
)
_DWG_LB = (r"(?<![0-9A-Za-z가-힣\-_./])"
           # 대시 없는 규격 표기(D10, HD13, SD400, DN100, M16, SS400, A1 용지 등)는 도면번호로 보지 않음
           r"(?!(?:SHD|UHD|HD|SD|DN|THK|STK|SS|SM|SN|PL|D|M|T|R|W|C|L|H|I|F)\d)(?![AB][0-5](?!\d))")
_DWG_LA = r"(?![0-9A-Za-z가-힣\-_]|\.\d|\s*[xX×*]\s*\d|/\d)"
DWG_NO_RE = re.compile(_DWG_LB + "(" + _DWG_CORE + ")" + _DWG_LA)
# 개정 접미 분리용
REV_SUFFIX_RE = re.compile(r"^(.*\d)-?(?:R|REV\.?)(\d{1,2})$")
# 라벨·목록 열 문맥이 없는 본문 텍스트에서는 규격 표기와 혼동되는 접두를 제외한다.
# (D10, HD13, SD400, DN100, KS D 3504, A1 용지, GL-1500 레벨, M16 볼트, SS400 강종 등)
FREE_TEXT_BAD_PREFIX = {
    "D", "HD", "SD", "SHD", "UHD", "DN", "KS", "THK", "T", "M", "H", "I", "C", "L", "F", "W", "R",
    "REV", "NO", "GL", "FL", "EL", "SL", "BL", "CL", "SS", "SM", "SN", "STK", "SPP", "SGP", "STS",
    "PE", "PVC", "PP", "CV", "VV", "SQ", "MM", "CM", "KG", "TON", "EA", "MPA", "FCK", "PL", "LV",
    "HFIX", "CD", "SEC", "QTY", "TEL", "FAX", "B",
}
PAPER_RE = re.compile(r"^[AB][0-5]$")


def _dwg_prefix(tok):
    m = re.match(r"([A-Z]+|[가-힣]+)", tok)
    return m.group(1) if m else ""


def find_dwg_no(text, vocab=None, strict=False):
    """텍스트에서 도면번호 1개를 찾는다 → (정규화 번호, 개정 접미 또는 None) / (None, None).

    vocab: 도면목록 번호 집합(정규화). 텍스트에 목록 번호가 있으면 그것을 우선한다.
    strict: 라벨·목록 열 문맥이 없는 본문 텍스트일 때 True (규격 표기 접두 제외).
    """
    if not text:
        return None, None
    if vocab:
        best = None
        for m in DWG_NO_RE.finditer(text):
            n, rv = split_rev_suffix(norm_dwg_no(m.group(1)))
            if n in vocab:
                return n, rv
            if best is None:
                best = (n, rv)
        # 공백·대시 표기 차이 등으로 정규식에 안 걸리는 목록 번호 직접 탐색
        compact = re.sub(r"\s+", "", text).upper()
        for v in sorted(vocab, key=len, reverse=True):
            i = compact.find(v)
            while i >= 0:
                before = compact[i - 1] if i > 0 else ""
                after = compact[i + len(v)] if i + len(v) < len(compact) else ""
                if not re.match(r"[0-9A-Z가-힣\-]", before) and not re.match(r"[0-9A-Z가-힣\-]", after):
                    return v, None
                i = compact.find(v, i + 1)
    for m in DWG_NO_RE.finditer(text):
        tok = norm_dwg_no(m.group(1))
        if strict:
            pre = _dwg_prefix(tok)
            if pre in FREE_TEXT_BAD_PREFIX or PAPER_RE.match(tok):
                continue
        return split_rev_suffix(tok)
    return None, None


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


def split_rev_suffix(no):
    """'S-201R1' → ('S-201', '1'). 개정 접미가 없으면 (no, None)."""
    if not no:
        return no, None
    m = REV_SUFFIX_RE.match(no)
    if m:
        return m.group(1).rstrip("-"), norm_rev(m.group(2))
    return no, None


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


def parse_title_block(text, vocab=None):
    """표제란 텍스트에서 도면번호/도면명/축척/개정을 추출.

    반환 dict 의 보조 키: _labeled(라벨 옆 값에서 도면번호를 찾음), _dwg_no_guess(라벨 없이 추정),
    _dwg_no_unparsed(라벨 옆 값이 있으나 도면번호로 해석 못 함), rev_suffix(S-201R1 의 '1').
    """
    res = {"dwg_no": None, "title": None, "scale": None, "scale_norm": None, "rev": None,
           "rev_suffix": None, "_labeled": False}
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
            no, rsuf = find_dwg_no(val, vocab, strict=False)
            if no:
                res["dwg_no"], res["rev_suffix"], res["_labeled"] = no, rsuf, True
            elif not res.get("_dwg_no_unparsed"):
                res["_dwg_no_unparsed"] = val
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

    # 라벨 기반으로 못 찾은 경우의 대체 규칙(본문 혼입을 고려해 엄격 모드)
    if not res["dwg_no"]:
        no, rsuf = find_dwg_no(text, vocab, strict=True)
        if no:
            res["dwg_no"], res["rev_suffix"] = no, rsuf
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
    if res["rev"] is None and res["rev_suffix"] is not None:
        res["rev"] = res["rev_suffix"]
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
# 정규식에 맞지 않아도 도면번호로 볼 만한 식별자(숫자 포함, 공백 없음)
LOOSE_ID_RE = re.compile(r"[0-9A-Za-z가-힣\-_.]{2,24}")


def _map_header(row):
    cols = {}
    for idx, cell in enumerate(row):
        c = (cell or "").replace("\n", " ").strip()
        for key, rx in HEADER_MAP:
            if key not in cols and rx.search(c):
                cols[key] = idx
                break
    return cols


def _list_dwg_no(cell, loc, raw, warnings):
    """목록 '도면번호' 칸 → (번호, 개정 접미). 실패하면 경고 후 (None, None)."""
    no, rsuf = find_dwg_no(cell, strict=False)
    if no:
        return no, rsuf
    compact = re.sub(r"\s+", "", cell)
    if LOOSE_ID_RE.fullmatch(compact) and re.search(r"\d", compact):
        warnings.append("%s: 도면목록 도면번호 '%s' 가 표준 형식이 아니어서 그대로 사용 — 확인 권장 (행: %s)"
                        % (loc, cell, raw))
        return split_rev_suffix(norm_dwg_no(compact))
    warnings.append("%s: 도면목록 행의 도면번호 '%s' 를 인식하지 못해 제외함 — 원본 확인 필요 (행: %s)"
                    % (loc, cell, raw))
    return None, None


def parse_list_tables(page, loc="", warnings=None):
    """pdfplumber extract_tables 로 도면목록 행을 추출. 해석 못 한 비어 있지 않은 행은 경고로 남긴다."""
    if warnings is None:
        warnings = []
    entries = []
    try:
        tables = page.extract_tables()
    except Exception as e:  # noqa: BLE001
        warnings.append("%s: 도면목록 표 추출 실패(%s: %s) — 텍스트 기반으로 재시도" % (loc, type(e).__name__, e))
        tables = []
    for tb in tables or []:
        cols = None
        for row in tb:
            if not row or not any((c or "").strip() for c in row):
                continue
            h = _map_header(row)
            if "dwg_no" in h and len(h) >= 2:
                cols = h  # 헤더(페이지·표마다 반복되는 헤더 포함)
                continue
            if cols is None:
                continue

            def cell(k):
                return ((row[cols[k]] or "").replace("\n", " ").strip()
                        if k in cols and cols[k] < len(row) else "")
            raw = " | ".join((c or "").replace("\n", " ") for c in row)
            dcell = cell("dwg_no")
            if not dcell:
                warnings.append("%s: 도면목록 행에 도면번호 칸이 비어 있어 제외함 — 원본 확인 필요 (행: %s)" % (loc, raw))
                continue
            no, rsuf = _list_dwg_no(dcell, loc, raw, warnings)
            if not no:
                continue
            scell = cell("scale")
            sraw, snorm = parse_scale(scell)
            if scell and not snorm:
                warnings.append("%s: 도면목록 %s 의 축척 '%s' 를 인식하지 못함 — 축척 비교 생략" % (loc, no, scell))
            rv = cell("rev")
            rev = norm_rev(rv) if rv else rsuf
            if rv and not re.fullmatch(r"[0-9]{1,2}|[A-Za-z]", rv.strip()):
                warnings.append("%s: 도면목록 %s 의 개정 '%s' 가 일반 형식이 아님 — 확인 권장" % (loc, no, rv))
            entries.append({
                "no": cell("no") or None,
                "dwg_no": no,
                "title": re.sub(r"\s+", " ", cell("title")) or None,
                "scale": sraw, "scale_norm": snorm,
                "rev": rev,
                "raw": raw,
                "method": "table",
            })
    return entries


ROW_LIKE_RE = re.compile(r"^\s*\d{1,4}[\s.)]+\S")


def parse_list_text(text, loc="", warnings=None):
    """표 인식 실패 시 텍스트 줄 단위로 도면목록 추정: [번호] 도면번호 도면명 [축척] [개정]."""
    if warnings is None:
        warnings = []
    entries = []
    for line in (text or "").splitlines():
        m = DWG_NO_RE.search(line)
        if not m:
            # 번호로 시작하는 행처럼 보이는데 도면번호를 못 찾으면 조용히 버리지 않는다
            if ROW_LIKE_RE.match(line) and not ALL_LABEL_RE.search(line) and not LIST_KEYWORDS_RE.search(line):
                warnings.append("%s: 도면목록 텍스트 행에서 도면번호를 인식하지 못해 제외함 — 원본 확인 필요 (행: %s)"
                                % (loc, line.strip()))
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
        no, rsuf = split_rev_suffix(norm_dwg_no(m.group(1)))
        entries.append({
            "no": pre if re.fullmatch(r"\d{1,4}", pre) else None,
            "dwg_no": no,
            "title": re.sub(r"\s+", " ", title) or None,
            "scale": sraw, "scale_norm": snorm, "rev": rev if rev is not None else rsuf,
            "raw": line.strip(), "method": "text",
        })
    return entries


# ---------------------------------------------------------------------------
# 메인 처리
# ---------------------------------------------------------------------------
# 기본 영역에서 표제란 라벨을 못 찾으면 차례로 시도하는 후보 영역 (x0, y0, x1, y1 비율, 좌상단 원점)
CANDIDATE_REGIONS = [
    ("우상단", (0.6, 0.0, 1.0, 0.25)),
    ("좌하단", (0.0, 0.75, 0.4, 1.0)),
    ("좌상단", (0.0, 0.0, 0.4, 0.25)),
    ("우측 띠", (0.7, 0.0, 1.0, 1.0)),
    ("하단 띠", (0.0, 0.8, 1.0, 1.0)),
    ("좌측 띠", (0.0, 0.0, 0.3, 1.0)),
    ("상단 띠", (0.0, 0.0, 1.0, 0.2)),
]


def crop_text(page, region, warnings=None, loc=""):
    x0, y0, x1, y1 = region
    bx0, btop, bx1, bbot = page.bbox
    W, H = bx1 - bx0, bbot - btop
    bbox = (bx0 + W * x0, btop + H * y0, bx0 + W * x1, btop + H * y1)
    try:
        return page.within_bbox(bbox).extract_text() or ""
    except Exception as e:  # noqa: BLE001
        if warnings is not None:
            warnings.append("%s: 영역 %s 텍스트 추출 실패(%s: %s)" % (loc, ",".join("%g" % v for v in region),
                                                              type(e).__name__, e))
        return ""


def _fmt_region(r):
    return ",".join("%g" % v for v in r)


def locate_title_block(page, full, region, vocab, warnings, loc):
    """표제란 위치 탐색: 지정(기본) 영역 → 후보 영역(모서리·띠) → 페이지 전체.

    '도면번호' 라벨 옆에서 번호를 찾은 첫 영역을 쓴다. 반환 (tb, ttext, region_name).
    """
    tried = [("기본", region)] + [(n, r) for n, r in CANDIDATE_REGIONS if list(r) != list(region)]
    first_tb, first_text = None, ""
    for name, reg in tried:
        txt = crop_text(page, reg, warnings, loc)
        tb = parse_title_block(txt, vocab)
        if first_tb is None:
            first_tb, first_text = tb, txt
        if tb.get("_labeled"):
            if name != "기본":
                warnings.append("%s: 기본 표제란 영역(%s)에서 라벨을 못 찾아 '%s' 영역(%s)의 표제란을 사용함"
                                % (loc, _fmt_region(region), name, _fmt_region(reg)))
            return tb, txt, name
    tb = parse_title_block(full, vocab)
    if tb.get("_labeled"):
        warnings.append("%s: 모든 후보 영역에서 표제란 라벨을 못 찾아 페이지 전체 텍스트의 라벨에서 도면번호를 찾음(%s)"
                        " — --title-region 지정 권장" % (loc, tb["dwg_no"]))
        return tb, full, "페이지 전체"
    # 라벨이 어디에도 없음 → 기본 영역의 추정값, 없으면 전체 텍스트 추정값
    if first_tb.get("dwg_no"):
        return first_tb, first_text, "기본"
    if tb.get("dwg_no"):
        warnings.append("%s: 표제란 라벨이 없어 페이지 전체 텍스트에서 도면번호를 추정함(%s) — 확인 필요"
                        % (loc, tb["dwg_no"]))
        return tb, full, "페이지 전체"
    return first_tb, first_text, "기본"


def _read_page(page, loc, warnings):
    try:
        return page.extract_text() or ""
    except Exception as e:  # noqa: BLE001
        warnings.append("%s: 텍스트 추출 실패(%s) — 이미지로 확인 필요" % (loc, e))
        return ""


def process_pdfs(paths, region, warnings):
    """모든 PDF 를 2단계로 처리: (1) 도면목록 수집 → (2) 목록 번호를 어휘로 표제란 해석."""
    pages_out, list_out = [], []
    with contextlib.ExitStack() as stack:
        docs = []
        for path in paths:
            try:
                pdf = stack.enter_context(pdfplumber.open(path))
            except Exception as e:  # noqa: BLE001
                sys.stderr.write("오류: PDF를 읽을 수 없습니다(%s): %s\n" % (path, e))
                sys.exit(2)
            docs.append((os.path.basename(path), pdf))

        # 1단계: 텍스트·스캔 판정·도면목록 후보
        info = []
        for fname, pdf in docs:
            for pno, page in enumerate(pdf.pages, start=1):
                loc = "%s p.%d" % (fname, pno)
                full = _read_page(page, loc, warnings)
                scanned = len(re.sub(r"\s", "", full)) < 30
                ents = []
                if not scanned and LIST_KEYWORDS_RE.search(full):
                    ents = parse_list_tables(page, loc, warnings)
                    if not ents:
                        dtext = crop_text(page, region, warnings, loc)
                        body = full.replace(dtext, "") if dtext else full
                        ents = parse_list_text(body, loc, warnings)
                        if ents:
                            warnings.append("%s: 도면목록 표 인식 실패 → 텍스트 줄 기반으로 %d행 추정(정확도 낮음)"
                                            % (loc, len(ents)))
                info.append((fname, pno, page, loc, full, scanned, ents))
        vocab = {e["dwg_no"] for it in info for e in it[6] if e.get("dwg_no")}

        # 2단계: 표제란
        for fname, pno, page, loc, full, scanned, ents in info:
            rec = {"file": fname, "page": pno, "dwg_no": None, "title": None, "scale": None,
                   "scale_norm": None, "rev": None, "is_list": False, "scanned": scanned,
                   "spec_tokens": [], "title_text_raw": "", "title_region": None}
            if scanned:
                warnings.append("%s: 텍스트가 거의 없음(스캔 도면 추정) — pdftoppm 으로 이미지 렌더링 후 육안 확인 필요" % loc)
                pages_out.append(rec)
                continue
            tb, ttext, rname = locate_title_block(page, full, region, vocab, warnings, loc)
            rec["title_text_raw"] = ttext
            rec["title_region"] = rname
            if tb.get("_dwg_no_guess"):
                warnings.append("%s: '도면번호' 라벨 없이 패턴으로 추정한 도면번호(%s) — 확인 권장" % (loc, tb["dwg_no"]))
            if tb.get("_dwg_no_unparsed") and not tb.get("_labeled"):
                warnings.append("%s: 도면번호 라벨 옆 값 '%s' 를 도면번호로 해석하지 못함 — 원본 확인 필요"
                                % (loc, tb["_dwg_no_unparsed"]))
            if tb.get("rev_suffix") is not None:
                rec["rev_suffix"] = tb["rev_suffix"]
            for k in ("dwg_no", "title", "scale", "scale_norm", "rev"):
                rec[k] = tb.get(k)
            if not rec["dwg_no"]:
                warnings.append("%s: 도면번호 미검출 — 표제란 위치(--title-region) 확인 필요" % loc)

            if LIST_KEYWORDS_RE.search(full):
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
    except Exception:  # noqa: BLE001
        sys.stderr.write("오류: --title-region 형식이 잘못되었습니다. 예) 0.6,0.75,1,1 (0~1 비율, x0<x1, y0<y1)\n")
        sys.exit(2)


def main():
    for s in (sys.stdout, sys.stderr):
        try:
            s.reconfigure(encoding="utf-8", errors="replace")
        except Exception:  # noqa: BLE001
            pass
    ap = argparse.ArgumentParser(description="도면 PDF 표제란·도면목록·규격 토큰 추출")
    ap.add_argument("pdfs", nargs="+", help="도면 PDF 파일(여러 개 가능)")
    ap.add_argument("--title-region", default="0.6,0.75,1,1",
                    help="표제란 영역 x0,y0,x1,y1 (페이지 대비 비율, 좌상단 원점). 기본 0.6,0.75,1,1. "
                         "이 영역에서 라벨을 못 찾으면 모서리·띠 영역, 페이지 전체 순으로 자동 재시도")
    ap.add_argument("--out", required=True, help="출력 JSON 경로 (예: dwg.json)")
    args = ap.parse_args()

    region = parse_region(args.title_region)
    for p in args.pdfs:
        if not os.path.isfile(p):
            sys.stderr.write("오류: 파일을 찾을 수 없습니다: %s\n" % p)
            sys.exit(2)

    warnings = []
    pages, dlist = process_pdfs(args.pdfs, region, warnings)

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
