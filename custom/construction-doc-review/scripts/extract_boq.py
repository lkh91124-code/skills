#!/usr/bin/env python3
"""내역서(BOQ) 추출기: xlsx / PDF → 표준 JSON.

사용법:
  python3 extract_boq.py 내역서.xlsx [--sheet 내역서] [--header-row 4] --out boq.json
  python3 extract_boq.py 내역서.pdf [--pages 3-20] --out boq.json

출력 형식:
  {"source", "rows": [{loc, segment, type, level, path, gids, closes, scope,
    name, spec, unit, qty, mat_up, mat_amt, lab_up, lab_amt, exp_up, exp_amt,
    tot_up, tot_amt, raw, low_confidence}], "warnings": [], "stats": {...}}

행 유형(type): item(세부 항목) / group(공종 헤더) / subtotal(소계·계) / total(합계·총계)
 - gids   : 해당 행이 속한 공종 그룹 id 목록(상위→하위)
 - closes : subtotal 이 닫는 그룹 id (추정)
 - scope  : 소계 비교 후보 그룹 id 목록 (가장 하위부터). total 은 ["*"]
원본 파일은 읽기만 하며 절대 수정하지 않는다.
"""

from __future__ import annotations

import argparse
import os
import re
import sys
import unicodedata

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from _numparse import parse_num, write_json  # noqa: E402

NUM_FIELDS = ["qty", "mat_up", "mat_amt", "lab_up", "lab_amt",
              "exp_up", "exp_amt", "tot_up", "tot_amt"]
MONEY_FIELDS = NUM_FIELDS[1:]
HEADER_SCAN_ROWS = 15
SCANNED_TEXT_MIN = 30  # 페이지 텍스트가 이보다 적으면 스캔(이미지) 페이지로 간주


# ---------------------------------------------------------------------------
# 텍스트/헤더 유틸
# ---------------------------------------------------------------------------
def cell_text(v) -> str:
    """셀 값을 한 줄 문자열로 (줄바꿈은 공백으로)."""
    if v is None:
        return ""
    if isinstance(v, float) and v.is_integer():
        v = int(v)
    s = str(v).replace("\r", "\n")
    return re.sub(r"[ \t]*\n[ \t]*", " ", s).strip()


def hnorm(v) -> str:
    """헤더/라벨 비교용 정규화: NFKC, 공백·괄호·구분기호 제거, 소문자."""
    s = unicodedata.normalize("NFKC", cell_text(v))
    s = re.sub(r"[\s\[\]\(\)<>〈〉【】「」『』·ㆍ/\\:_\-.,]", "", s)
    return s.lower()


def header_field(label: str):
    """정규화된 헤더 라벨 → 필드명. 해당 없으면 None."""
    if not label:
        return None
    if label in ("no", "번호", "순번", "항번", "코드", "호표", "일련번호") or label.startswith("코드"):
        return "code"
    if any(k in label for k in ("품명", "공종", "명칭", "품목", "공사명", "항목")):
        return "name"
    if "규격" in label or "사양" in label:
        return "spec"
    if label in ("단위",):
        return "unit"
    if label in ("수량", "물량"):
        return "qty"
    if label.startswith("비고"):
        return "note"
    grp = None
    if "재료" in label:
        grp = "mat"
    elif "노무" in label:
        grp = "lab"
    elif "경비" in label:
        grp = "exp"
    elif any(k in label for k in ("합계", "총액", "합")) or label in ("계", "계단가", "계금액"):
        grp = "tot"
    sub = None
    if "단가" in label:
        sub = "up"
    elif "금액" in label:
        sub = "amt"
    if grp and sub:
        return f"{grp}_{sub}"
    if sub:  # 단일 단가/금액 열
        return f"tot_{sub}"
    if grp:  # "재료비" 단독 열은 금액으로 간주
        return f"{grp}_amt"
    return None


def build_colmap(top, bottom=None):
    """헤더 행(1행 또는 2행)으로 열 매핑을 만든다. 반환: {field: col_index}."""
    n = max(len(top), len(bottom or []))
    labels = []
    last_top = ""
    for i in range(n):
        t = hnorm(top[i]) if i < len(top) else ""
        b = hnorm(bottom[i]) if bottom and i < len(bottom) else ""
        # 2행 헤더에서 상단 병합(재료비 등)이 빈칸으로 오면 왼쪽 값 이어받기
        if not t and b in ("단가", "금액") and last_top:
            t = last_top
        if t:
            last_top = t
        labels.append(t + (b if b and b != t else ""))
    cmap = {}
    for i, lab in enumerate(labels):
        f = header_field(lab)
        if f and f not in cmap:
            cmap[f] = i
    return cmap


def header_score(row) -> int:
    """행이 헤더처럼 보이는 정도 (name/unit/qty/금액류 키워드 수)."""
    fields = {header_field(hnorm(c)) for c in row}
    fields.discard(None)
    score = sum(1 for f in ("name", "unit", "qty", "spec") if f in fields)
    if any(f.endswith("_up") or f.endswith("_amt") for f in fields):
        score += 1
    return score


def is_header_row(row) -> bool:
    fields = {header_field(hnorm(c)) for c in row}
    return "name" in fields and ("unit" in fields or "qty" in fields)


def is_subheader_row(row) -> bool:
    return sum(1 for c in row if hnorm(c) in ("단가", "금액")) >= 2


def detect_header(rows, forced=None):
    """첫 15행 안에서 헤더를 찾는다. 반환: (cmap, 데이터 시작 인덱스) 또는 (None, 0)."""
    cand = [forced] if forced is not None else range(min(HEADER_SCAN_ROWS, len(rows)))
    for i in cand:
        if i is None or i >= len(rows):
            continue
        if forced is None and not is_header_row(rows[i]):
            continue
        if i + 1 < len(rows) and is_subheader_row(rows[i + 1]):
            return build_colmap(rows[i], rows[i + 1]), i + 2
        return build_colmap(rows[i]), i + 1
    return None, 0


# ---------------------------------------------------------------------------
# 행 분류 / 계층
# ---------------------------------------------------------------------------
_TOTAL_KEYS = ("합계", "총계", "총공사비", "총합", "공사비계", "총액")
_HANGUL_ENUM = "가나다라마바사아자차카타파하"


def row_kind_label(label: str):
    """라벨로 소계/합계 판단. 반환 'total' / 'subtotal' / None."""
    n = hnorm(label)
    if not n:
        return None
    if any(k in n for k in _TOTAL_KEYS):
        return "total"
    if n in ("계", "소계", "소소계", "중계") or "소계" in n:
        return "subtotal"
    raw = cell_text(label)
    if re.search(r"\s계\s*\]?$", raw) or re.match(r"^\[.*계\s*\]$", raw):
        return "subtotal"  # "가설공사 계", "[가설공사 계]"
    return None


def numbering_style(raw_name: str, indent: int = 0) -> str:
    """번호 체계로 계층 스타일 키를 만든다 (NFKC 전에 원문으로 판단)."""
    s = raw_name.lstrip()
    m = re.match(r"^(\d+(?:\.\d+)+)\.?\s*", s)
    if m:
        return f"multi{len(m.group(1).split('.'))}"
    if re.match(r"^\d+\.\s*\D", s):
        return "n."
    if re.match(r"^\d+\)", s):
        return "n)"
    if re.match(r"^\(\d+\)", s):
        return "(n)"
    if re.match(rf"^[{_HANGUL_ENUM}]\.\s*", s):
        return "가."
    if re.match(rf"^\([{_HANGUL_ENUM}]\)", s):
        return "(가)"
    if re.match(r"^[①-⑳]", s):
        return "circ"
    if re.match(r"^[ⅠⅡⅢⅣⅤⅥⅦⅧⅨⅩ]|^[IVX]+\.\s", s):
        return "roman"
    if re.match(r"^[A-Z]\.\s", s):
        return "A."
    if re.match(r"^\[.+\]$", s):
        return "bracket"
    return f"plain{indent}"


def bare_name(name: str) -> str:
    """번호·괄호를 뗀 비교용 공종명."""
    s = re.sub(r"^(\d+(?:\.\d+)*[.)]?|\(\d+\)|[①-⑳]|[ⅠⅡⅢⅣⅤⅥⅦⅧⅨⅩ]\.?|[가-하]\.|[A-Z]\.)\s*", "",
               cell_text(name))
    return hnorm(s)


class Hierarchy:
    """공종 그룹 스택. 번호 체계 스타일이 이미 스택에 있으면 그 수준으로 되돌아간다."""

    def __init__(self):
        self.stack = []  # [(style, gid, name)]
        self.next_gid = 1

    def reset(self):
        self.stack = []

    def push_group(self, name, style):
        styles = [s for s, _, _ in self.stack]
        if style in styles:
            del self.stack[styles.index(style):]
        gid = self.next_gid
        self.next_gid += 1
        self.stack.append((style, gid, cell_text(name)))
        return gid

    def close_for_subtotal(self, label):
        """소계가 닫는 그룹을 추정하고 스택에서 제거. 반환 (closes, scope, 닫힌 그룹의 경로)."""
        if not self.stack:
            return None, [], []
        lab = hnorm(label)
        idx = len(self.stack) - 1
        for i in range(len(self.stack) - 1, -1, -1):
            b = bare_name(self.stack[i][2])
            if b and b in lab:
                idx = i
                break
        scope = [g for _, g, _ in reversed(self.stack[: idx + 1])]
        closes = self.stack[idx][1]
        closed_path = [n for _, _, n in self.stack[: idx + 1]]
        del self.stack[idx:]
        return closes, scope, closed_path

    @property
    def path(self):
        return [n for _, _, n in self.stack]

    @property
    def gids(self):
        return [g for _, g, _ in self.stack]


# ---------------------------------------------------------------------------
# 공통 행 처리기
# ---------------------------------------------------------------------------
class RowBuilder:
    def __init__(self, warnings):
        self.rows = []
        self.warnings = warnings
        self.h = Hierarchy()

    def new_segment(self):
        self.h.reset()

    def add(self, cells, cmap, loc, segment, indent=0, header_labels=None,
            text_strategy=False):
        def get(f):
            i = cmap.get(f)
            return cells[i] if i is not None and i < len(cells) else None

        raw_name = "" if get("name") is None else str(get("name"))
        name = cell_text(raw_name)
        spec = cell_text(get("spec"))
        unit = cell_text(get("unit"))
        code = cell_text(get("code"))
        low = False
        vals = {}
        for f in NUM_FIELDS:
            v = get(f)
            num, rawv = parse_num(v if not isinstance(v, str) else v.replace("\n", " "))
            if num is None and cell_text(v) and cell_text(v) not in ("-", "—", "※"):
                low = True  # 숫자 열인데 해석 불가
            if isinstance(v, str) and "\n" in v.strip() and num is None:
                low = True
            vals[f] = num
        raw = {}
        for f, i in cmap.items():
            if i < len(cells) and cells[i] not in (None, ""):
                raw[f] = cell_text(cells[i])
        if not raw:
            return  # 빈 행

        has_num = any(vals[f] is not None for f in NUM_FIELDS)
        label = name or code or spec
        kind = row_kind_label(name) or (row_kind_label(code) if not name else None)

        # 줄바꿈(랩) 이어쓰기: 규격만 있는 행 → 앞 행 규격에 붙임
        prev = self.rows[-1] if self.rows else None
        if (not kind and not name and spec and not unit and not has_num and prev
                and prev["type"] == "item"):
            prev["spec"] = (prev["spec"] + " " + spec).strip()
            prev["raw"]["spec"] = prev["spec"]
            return
        # text 전략 PDF: 번호 없는 짧은 품명만 있는 행은 앞 항목 품명의 이어쓰기로 간주
        if (text_strategy and not kind and name and not spec and not unit and not has_num
                and prev and prev["type"] == "item"
                and numbering_style(raw_name).startswith("plain") and len(name) <= 12):
            prev["name"] = (prev["name"] + " " + name).strip()
            prev["low_confidence"] = True
            return

        row = {
            "loc": loc, "segment": segment, "type": None, "level": 0, "path": [],
            "gids": [], "closes": None, "scope": [],
            "code": code, "name": name, "spec": spec, "unit": unit,
            **vals, "raw": raw, "low_confidence": low,
        }
        if kind == "total":
            row["type"] = "total"
            row["scope"] = ["*"]
            row["level"] = 0
            self.h.reset()
        elif kind == "subtotal":
            closes, scope, closed_path = self.h.close_for_subtotal(label)
            # path = 닫힌 공종의 경로. closes 가 None 이면 '직전 소계 이후 항목'과 비교
            row.update(type="subtotal", closes=closes, scope=scope,
                       level=max(len(closed_path), 1), path=closed_path)
        elif name and not unit and vals["qty"] is None and not any(
                vals[f] is not None for f in ("mat_up", "lab_up", "exp_up", "tot_up")):
            style = numbering_style(raw_name, indent)
            gid = self.h.push_group(name, style)
            row.update(type="group", level=len(self.h.stack), path=self.h.path[:-1],
                       gids=self.h.gids)
            row["gid"] = gid
        elif name or has_num or unit:
            row.update(type="item", level=len(self.h.stack) + 1, path=self.h.path,
                       gids=self.h.gids)
            if not name:
                row["low_confidence"] = True
        else:
            return
        self.rows.append(row)


# ---------------------------------------------------------------------------
# XLSX
# ---------------------------------------------------------------------------
def extract_xlsx(path, sheet, header_row, warnings, builder):
    try:
        import openpyxl
    except ImportError:
        sys.exit("오류: openpyxl 이 설치되어 있지 않습니다 (pip install openpyxl).")
    try:
        wb = openpyxl.load_workbook(path, data_only=True)
    except Exception as e:  # noqa: BLE001
        sys.exit(f"오류: 엑셀 파일을 열 수 없습니다: {e}")

    if sheet:
        if sheet not in wb.sheetnames:
            sys.exit(f"오류: 시트 '{sheet}' 가 없습니다. 시트 목록: {', '.join(wb.sheetnames)}")
        sheets = [wb[sheet]]
    else:
        sheets = wb.worksheets

    found_any = False
    for ws in sheets:
        if ws.max_row < 2:
            continue
        grid = [[c.value for c in r] for r in ws.iter_rows()]
        if not grid:
            continue
        ncols = max(len(r) for r in grid)
        grid = [r + [None] * (ncols - len(r)) for r in grid]
        merges = list(ws.merged_cells.ranges)

        # 헤더 탐지용: 병합 셀을 전부 채운 사본
        filled = [r[:] for r in grid]
        for m in merges:
            v = grid[m.min_row - 1][m.min_col - 1]
            for rr in range(m.min_row - 1, min(m.max_row, len(filled))):
                for cc in range(m.min_col - 1, min(m.max_col, ncols)):
                    filled[rr][cc] = v

        forced = header_row - 1 if header_row else None
        cmap, start = detect_header(filled, forced)
        if not cmap or "name" not in cmap:
            if sheet:
                warnings.append(f"시트 '{ws.title}': 헤더(공종/품명·단위·수량)를 찾지 못함")
            continue
        found_any = True
        builder.new_segment()

        # 데이터 영역: 품명/규격 열의 '세로' 병합만 아래로 채움 (금액 열은 채우지 않음)
        data = [r[:] for r in grid]
        text_cols = {cmap.get("name"), cmap.get("spec"), cmap.get("code")} - {None}
        for m in merges:
            if m.min_row - 1 < start or m.min_col != m.max_col or m.max_row == m.min_row:
                continue
            c = m.min_col - 1
            if c in text_cols:
                v = grid[m.min_row - 1][c]
                for rr in range(m.min_row, min(m.max_row, len(data))):
                    data[rr][c] = v

        name_col = cmap["name"]
        for ri in range(start, len(data)):
            cells = data[ri]
            if all(v in (None, "") for v in cells):
                continue
            if header_score(cells) >= 3 and is_header_row(cells):
                continue  # 인쇄용 반복 헤더
            indent = 0
            try:
                al = ws.cell(row=ri + 1, column=name_col + 1).alignment
                indent = int(al.indent or 0)
            except Exception:  # noqa: BLE001
                pass
            raw_name = cells[name_col]
            if isinstance(raw_name, str):
                indent += len(raw_name) - len(raw_name.lstrip(" "))
            try:
                builder.add(cells, cmap, {"sheet": ws.title, "row": ri + 1},
                            segment=ws.title, indent=indent)
            except Exception as e:  # noqa: BLE001 - 행 단위 오류는 경고 후 계속
                warnings.append(f"시트 '{ws.title}' {ri + 1}행 처리 오류: {e}")

        # 캐시 값 없음(수식만 있고 계산값 미저장) 검사
        seg_items = [r for r in builder.rows if r["segment"] == ws.title and r["type"] == "item"]
        amt_fields = [f for f in MONEY_FIELDS if f.endswith("_amt") and f in cmap]
        if seg_items and amt_fields and all(r[f] is None for r in seg_items for f in amt_fields):
            has_formula = False
            try:
                wbf = openpyxl.load_workbook(path, data_only=False, read_only=True)
                for row in wbf[ws.title].iter_rows(min_row=start + 1, values_only=True):
                    if any(isinstance(v, str) and v.startswith("=") for v in row):
                        has_formula = True
                        break
            except Exception:  # noqa: BLE001
                pass
            if has_formula:
                warnings.append(
                    f"시트 '{ws.title}': 금액 열의 수식 계산값이 저장되어 있지 않습니다. "
                    "Excel(또는 LibreOffice)에서 파일을 열어 저장한 뒤 다시 실행하세요.")
            else:
                warnings.append(f"시트 '{ws.title}': 금액 열 값이 모두 비어 있습니다.")

    if not found_any:
        sys.exit("오류: 내역서 헤더(공종/품명, 단위, 수량)를 찾지 못했습니다. "
                 "--sheet 또는 --header-row 로 지정하세요.")
    return []


# ---------------------------------------------------------------------------
# PDF
# ---------------------------------------------------------------------------
def parse_pages(spec, n):
    if not spec:
        return list(range(1, n + 1))
    pages = set()
    for part in spec.split(","):
        part = part.strip()
        if not part:
            continue
        if "-" in part:
            a, b = part.split("-", 1)
            a = int(a) if a else 1
            b = int(b) if b else n
            pages.update(range(a, b + 1))
        else:
            pages.add(int(part))
    return sorted(p for p in pages if 1 <= p <= n)


def extract_pdf(path, pages_spec, header_row, warnings, builder):
    try:
        import pdfplumber
    except ImportError:
        sys.exit("오류: pdfplumber 가 설치되어 있지 않습니다 (pip install pdfplumber).")
    scanned = []
    try:
        pdf = pdfplumber.open(path)
    except Exception as e:  # noqa: BLE001
        sys.exit(f"오류: PDF 를 열 수 없습니다: {e}")

    cmap, ncols = None, None
    builder.new_segment()
    with pdf:
        try:
            pages = parse_pages(pages_spec, len(pdf.pages))
        except ValueError:
            sys.exit(f"오류: --pages 형식이 잘못되었습니다: {pages_spec} (예: 3-20 또는 1,3,5-7)")
        for pno in pages:
            page = pdf.pages[pno - 1]
            text = page.extract_text() or ""
            if len(re.sub(r"\s", "", text)) < SCANNED_TEXT_MIN:
                scanned.append(pno)
                continue
            strategy = "lines"
            tables = page.extract_tables({"vertical_strategy": "lines",
                                          "horizontal_strategy": "lines"}) or []
            tables = [t for t in tables if len(t) >= 2 and max(len(r) for r in t) >= 4]
            if not tables:
                strategy = "text"
                tables = page.extract_tables({"vertical_strategy": "text",
                                              "horizontal_strategy": "text"}) or []
                tables = [t for t in tables if len(t) >= 2]
                if tables:
                    warnings.append(f"{pno}페이지: 표 선이 없어 텍스트 기반으로 추출 (정확도 낮음)")
            if not tables:
                warnings.append(f"{pno}페이지: 표를 찾지 못함")
                continue
            for t in tables:
                width = max(len(r) for r in t)
                t = [list(r) + [None] * (width - len(r)) for r in t]
                forced = (header_row - 1) if (header_row and cmap is None) else None
                new_map, start = detect_header(t, forced)
                if new_map and "name" in new_map:
                    cmap, ncols, start_i = new_map, width, start
                elif cmap is not None:
                    start_i = 0
                    if width != ncols:
                        warnings.append(f"{pno}페이지: 열 수({width})가 헤더({ncols})와 달라 "
                                        "위치 기준으로 매핑 (저신뢰)")
                else:
                    warnings.append(f"{pno}페이지: 헤더 없는 표를 건너뜀")
                    continue
                for ri in range(start_i, len(t)):
                    cells = t[ri]
                    if is_header_row(cells) or is_subheader_row(cells):
                        continue  # 페이지마다 반복되는 헤더
                    before = len(builder.rows)
                    try:
                        builder.add(cells, cmap, {"page": pno, "row": ri + 1},
                                    segment="pdf", text_strategy=(strategy == "text"))
                    except Exception as e:  # noqa: BLE001
                        warnings.append(f"{pno}페이지 {ri + 1}행 처리 오류: {e}")
                        continue
                    if (strategy == "text" or width != ncols) and len(builder.rows) > before:
                        builder.rows[-1]["low_confidence"] = True
    if cmap is None and not scanned:
        sys.exit("오류: PDF 에서 내역서 헤더(공종/품명, 단위, 수량)를 찾지 못했습니다. "
                 "--pages 로 내역 페이지를 지정하세요.")
    if scanned:
        warnings.append(f"스캔(이미지) 페이지로 보이는 쪽: {scanned} — pdftoppm 으로 이미지화하여 육안 판독 필요")
    return scanned


# ---------------------------------------------------------------------------
def main():
    ap = argparse.ArgumentParser(description="내역서(BOQ) xlsx/PDF 를 표준 JSON 으로 추출")
    ap.add_argument("file", help="내역서 파일 (.xlsx/.xlsm/.pdf)")
    ap.add_argument("--sheet", help="엑셀 시트 이름 (미지정 시 헤더가 있는 모든 시트)")
    ap.add_argument("--pages", help="PDF 페이지 범위 (예: 3-20, 1,3,5-7)")
    ap.add_argument("--header-row", type=int, help="헤더 행 번호(1부터). 미지정 시 자동 탐지")
    ap.add_argument("--out", required=True, help="출력 JSON 경로")
    args = ap.parse_args()

    if not os.path.isfile(args.file):
        sys.exit(f"오류: 파일이 없습니다: {args.file}")
    ext = os.path.splitext(args.file)[1].lower()
    warnings = []
    builder = RowBuilder(warnings)
    if ext in (".xlsx", ".xlsm"):
        scanned = extract_xlsx(args.file, args.sheet, args.header_row, warnings, builder)
    elif ext == ".pdf":
        scanned = extract_pdf(args.file, args.pages, args.header_row, warnings, builder)
    elif ext == ".xls":
        sys.exit("오류: .xls(구형 엑셀)는 지원하지 않습니다. Excel 에서 .xlsx 로 다시 저장해 주세요.")
    else:
        sys.exit(f"오류: 지원하지 않는 형식입니다: {ext} (xlsx, pdf 만 지원)")

    rows = builder.rows
    n = len(rows)
    low = sum(1 for r in rows if r["low_confidence"])
    out = {
        "source": os.path.abspath(args.file),
        "rows": rows,
        "warnings": warnings,
        "stats": {
            "rows": n,
            "items": sum(1 for r in rows if r["type"] == "item"),
            "low_confidence_ratio": round(low / n, 3) if n else 0.0,
            "scanned_pages": scanned,
        },
    }
    write_json(out, args.out)
    print(f"추출 완료: {n}행 (항목 {out['stats']['items']}), 저신뢰 비율 "
          f"{out['stats']['low_confidence_ratio']:.1%}, 경고 {len(warnings)}건 → {args.out}")
    for w in warnings:
        print(f"  [경고] {w}", file=sys.stderr)


if __name__ == "__main__":
    main()
