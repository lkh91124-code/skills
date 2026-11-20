"""construction-doc-review 공용 헬퍼.

내역서/원가계산서/도면 검토 스크립트가 공통으로 쓰는 숫자 파싱, 단위·품명 정규화,
근사 비교, 지적사항(finding) 생성, JSON 입출력 함수를 모아 둔다.
다른 스크립트가 아래 이름을 그대로 import 하므로 이름/시그니처를 바꾸지 말 것.
"""

from __future__ import annotations

import json
import os
import re
import unicodedata
from typing import Any, Optional, Tuple

__all__ = [
    "parse_num",
    "norm_unit",
    "norm_key",
    "approx_equal",
    "make_finding",
    "write_json",
    "read_json",
    "SEVERITIES",
    "fmt_num",
    "match_with_rounding",
    "ROUNDING_UNITS_COST",
    "ROUNDING_UNITS_ITEM",
]

SEVERITIES = ("높음", "중간", "낮음")

# 값이 "없음"을 뜻하는 표기 (빈칸, 대시류, 참조표시 등)
_EMPTY_TOKENS = {"", "-", "—", "–", "―", "－", "‐", "※", "/", "n/a", "N/A", "없음"}

# 음수 표기: 괄호 (1,234), 삼각형 △1,234 / ▲1,234, 앞의 마이너스
_NEG_PREFIXES = ("△", "▲", "▵", "-", "−", "－")

_NUM_RE = re.compile(r"^[+]?\d*\.?\d+(?:[eE][+-]?\d+)?$")


def parse_num(s: Any) -> Tuple[Optional[float], str]:
    """문자열을 숫자로 변환한다.

    반환값: (float 또는 None, 원본 문자열)
    - "1,234" / "1,234.5" / "１，２３４" (전각) / "1,234원" / "12.5%" / " 1 234 " 처리
    - "(1,234)", "△1,234", "▲1,234", "-1,234" → 음수
    - None, "", "-", "—", "※" 등 → None
    - int/float 입력은 그대로 통과 (bool 은 None)
    """
    if s is None:
        return None, ""
    if isinstance(s, bool):
        return None, str(s)
    if isinstance(s, (int, float)):
        if s != s:  # NaN
            return None, str(s)
        return float(s), str(s)

    raw = str(s)
    t = unicodedata.normalize("NFKC", raw).strip()
    if t in _EMPTY_TOKENS:
        return None, raw

    # 공백, 콤마, 통화/단위 접미사 제거
    t = re.sub(r"\s+", "", t)
    t = t.replace(",", "").replace("₩", "").replace("\\", "")
    for suffix in ("원", "%", "₩"):
        if t.endswith(suffix):
            t = t[: -len(suffix)]

    neg = False
    if t.startswith("(") and t.endswith(")"):
        neg = True
        t = t[1:-1]
    while t and t[0] in _NEG_PREFIXES:
        neg = not neg if t[0] in ("-", "−", "－") else True
        t = t[1:]
    if t.endswith("-"):  # 일부 회계 양식의 후행 마이너스 "1234-"
        neg = True
        t = t[:-1]

    if not t or not _NUM_RE.match(t):
        return None, raw
    try:
        v = float(t)
    except ValueError:
        return None, raw
    return (-v if neg else v), raw


# 단위 정규화 표: NFKC 이후 소문자 기준 키 → 표준 단위
_UNIT_MAP = {
    "m2": "m2", "m²": "m2", "㎡": "m2", "sqm": "m2", "평방미터": "m2",
    "m3": "m3", "m³": "m3", "㎥": "m3", "cum": "m3", "입방미터": "m3",
    "m": "m", "미터": "m",
    "cm": "cm", "㎝": "cm",
    "mm": "mm", "㎜": "mm",
    "km": "km", "㎞": "km",
    "kg": "kg", "㎏": "kg",
    "g": "g",
    "ton": "ton", "t": "ton", "톤": "ton", "tonne": "ton",
    "ea": "EA", "개": "EA", "개소": "개소", "본": "본", "매": "매", "대": "대",
    "set": "SET", "조": "조",
    "식": "식", "式": "식", "ls": "식", "l.s": "식", "l.s.": "식",
    "l": "L", "ℓ": "L", "리터": "L",
    "hr": "hr", "시간": "hr", "인": "인", "일": "일", "월": "월",
    "인일": "인일", "공": "인",
}


def norm_unit(u: Any) -> str:
    """단위 표기를 표준형으로 정규화한다. (㎡/M2/m² → m2, EA/개 → EA, M → m 등)

    알 수 없는 단위는 NFKC + 공백제거 + 소문자로 돌려준다.
    """
    if u is None:
        return ""
    raw = str(u).strip()
    if not raw:
        return ""
    # NFKC 전에 특수 기호를 먼저 매핑 (NFKC 는 ㎡ → m2 로 바꿔주지만 안전하게 이중 처리)
    if raw in _UNIT_MAP:
        return _UNIT_MAP[raw]
    t = unicodedata.normalize("NFKC", raw)
    t = re.sub(r"\s+", "", t)
    t = t.replace("²", "2").replace("³", "3")
    low = t.lower()
    if low in _UNIT_MAP:
        return _UNIT_MAP[low]
    return low


_MULT_RE = re.compile(r"[×xX*✕✖]")
_DIA_RE = re.compile(r"[ΦφøØϕ⌀]")


def norm_key(name: Any, spec: Any = "") -> str:
    """품명+규격을 비교용 키로 정규화한다.

    NFKC → 공백 제거 → 곱셈기호(×,X,x,*) 통일 → 직경기호(Φ,φ,ø) → D → 소문자.
    결과 형식: "품명|규격"
    """

    def _n(v: Any) -> str:
        if v is None:
            return ""
        t = unicodedata.normalize("NFKC", str(v))
        t = re.sub(r"\s+", "", t)
        t = _DIA_RE.sub("D", t)
        t = _MULT_RE.sub("x", t)
        return t.lower()

    return f"{_n(name)}|{_n(spec)}"


def approx_equal(a: Optional[float], b: Optional[float], tol: float = 1.0) -> bool:
    """절대오차 tol 이내 또는 상대오차 1e-9 이내면 True. 둘 중 하나가 None 이면 False."""
    if a is None or b is None:
        return False
    diff = abs(float(a) - float(b))
    if diff <= tol:
        return True
    scale = max(abs(float(a)), abs(float(b)))
    return scale > 0 and diff / scale <= 1e-9


def make_finding(code: str, severity: str, location: Any, content: str,
                 evidence: str, recommendation: str) -> dict:
    """표준 지적사항 dict 를 만든다. severity 는 높음/중간/낮음 중 하나."""
    if severity not in SEVERITIES:
        severity = "중간"
    return {
        "code": code,
        "severity": severity,
        "location": location,
        "content": content,
        "evidence": evidence,
        "recommendation": recommendation,
    }


def write_json(obj: Any, path: str) -> None:
    """UTF-8, ensure_ascii=False 로 JSON 저장 (상위 폴더 자동 생성)."""
    d = os.path.dirname(os.path.abspath(path))
    os.makedirs(d, exist_ok=True)
    with open(path, "w", encoding="utf-8") as f:
        json.dump(obj, f, ensure_ascii=False, indent=2)


def read_json(path: str) -> Any:
    """UTF-8 JSON 읽기."""
    with open(path, "r", encoding="utf-8") as f:
        return json.load(f)


def fmt_num(v: Optional[float]) -> str:
    """보고용 숫자 포맷 (정수면 콤마, 소수면 소수 3자리까지)."""
    if v is None:
        return "-"
    if abs(v - round(v)) < 1e-9:
        return f"{int(round(v)):,}"
    return f"{v:,.3f}".rstrip("0").rstrip(".")


ROUNDING_UNITS_COST = (1, 10, 100, 1000)  # 원가계산서 요율·합계 비목: 10/100/1000원 단수정리 관행 허용
ROUNDING_UNITS_ITEM = (1,)                 # 내역서 항목 산술(수량×단가, 재+노+경): 1원 단위만 허용


def match_with_rounding(expected: Optional[float], actual: Optional[float],
                        tol: float = 1.0, units=ROUNDING_UNITS_COST) -> bool:
    """계산값 expected 를 units 의 각 단위(원)로 절사(floor/trunc)·반올림한 값 중 하나가
    actual 과 tol 이내로 일치하면 True.

    - units=(1, 10, 100, 1000) (기본): 원가계산서 요율 비목·총원가 등 단수정리 관행 허용.
    - units=(1,): 내역서 항목 산술용. 1원 미만 절사·반올림만 허용하므로
      10~999원 단위의 실제 오류를 숨기지 않는다.
    """
    if expected is None or actual is None:
        return False
    if approx_equal(expected, actual, tol):
        return True
    import math
    for u in units:
        q = expected / u
        cands = [math.floor(q) * u, round(q) * u, math.trunc(q) * u]
        if u == 1:  # 절상은 1원 단위만 허용 (그 이상은 과도한 허용)
            cands.append(math.ceil(q))
        for v in cands:
            if approx_equal(v, actual, tol):
                return True
    return False
