"""주소 문자열 정규화·파싱 유틸리티.

Juso API나 엑셀에 의존하지 않는 순수 텍스트 처리만 담는다.
"""

from __future__ import annotations

import json
import re
from pathlib import Path
from typing import Any, Iterable

from ..regions import SIDO_RE as REGIONS_SIDO_RE

SPECIAL_CHARS = re.compile(r"[%,=><\[\]]+")


def normalize_spaces(value: Any) -> str:
    if value is None:
        return ""
    return re.sub(r"\s+", " ", str(value).replace("\u3000", " ")).strip()


def norm(value: Any) -> str:
    s = normalize_spaces(value)
    s = SPECIAL_CHARS.sub(" ", s)
    s = re.sub(r"[,，]+", " ", s)
    return re.sub(r"\s+", " ", s).strip()


def juso_keyword(value: Any, preserve_commas: bool = False) -> str:
    if preserve_commas:
        return normalize_spaces(SPECIAL_CHARS.sub(" ", normalize_spaces(value)))
    return norm(value)


def typo_fix_first_pass(value: Any) -> str:
    s = normalize_spaces(value)
    replacements = [
        ("서울특별시 시 ", "서울특별시 "),
        ("인천광역시 시 ", "인천광역시 "),
        ("경기도 도 ", "경기도 "),
        ("틀벽시", "특별시"),
    ]
    for a, b in replacements:
        s = s.replace(a, b)
    return normalize_spaces(s)


# 데이터에서 실제로 발견된 오타 교정 규칙. CLI의 --typo-rules JSON으로
# 코드 수정 없이 규칙을 추가할 수 있다.
BASE_TYPO_REPLACEMENTS: list[tuple[str, str]] = [
    ("서울틀벽시", "서울특별시"),
    ("서울특벽시", "서울특별시"),
    ("서욽특별시", "서울특별시"),
    ("서울시", "서울특별시"),
    ("인천시", "인천광역시"),
    ("인천 광역시", "인천광역시"),
    ("인천광역시 시 ", "인천광역시 "),
    ("경기도 도 ", "경기도 "),
    ("논현도", "논현동"),
    ("프루지오", "푸르지오"),
    ("게양대로", "계양대로"),
]

_extra_typo_replacements: list[tuple[str, str]] = []


def set_extra_typo_rules(rules: Iterable[tuple[str, str]] | None) -> None:
    global _extra_typo_replacements
    _extra_typo_replacements = [(str(a), str(b)) for a, b in (rules or [])]


def load_typo_rules(path: Path) -> list[tuple[str, str]]:
    """[["프루지오", "푸르지오"], ...] 또는 {"replacements": [...]} 형식의 JSON을 읽는다."""
    data = json.loads(path.read_text(encoding="utf-8"))
    if isinstance(data, dict):
        data = data.get("replacements", [])
    rules: list[tuple[str, str]] = []
    for item in data:
        if not (isinstance(item, (list, tuple)) and len(item) == 2):
            raise ValueError(
                f'오타 규칙 형식 오류: {item!r} (예: ["프루지오", "푸르지오"])'
            )
        rules.append((str(item[0]), str(item[1])))
    return rules


def typo_fix(value: Any) -> str:
    s = norm(value)
    s = re.sub(r"^\d{5}\s+", "", s)
    for a, b in BASE_TYPO_REPLACEMENTS + _extra_typo_replacements:
        s = s.replace(a, b)
    # 시/도 약칭은 문자열 시작에서만 확장한다. 중간 치환은 '시민로' 같은 도로명이나
    # '서울 빌라' 같은 건물명까지 훼손한다.
    s = re.sub(r"^서울\s+", "서울특별시 ", s)
    s = re.sub(r"^인천\s+", "인천광역시 ", s)
    s = re.sub(r"^경기\s+", "경기도 ", s)
    s = re.sub(r"인천광역시\s+남구\b", "인천광역시 미추홀구", s)
    # 산 지번은 '산12-3' 표기로 통일해 지번 파싱과 Juso 지번주소 대조를 일관되게 한다.
    s = re.sub(r"((?:동|가|리)\s+)산\s+(\d)", r"\1산\2", s)
    return norm(s)


def clean_raw(value: Any) -> str:
    s = typo_fix(value)
    # '3충'·'지하1충' 처럼 층(層)을 충으로 적은 오타 교정. '충청/충정로' 등은
    # 충 뒤 글자로 걸러 건드리지 않는다.
    s = re.sub(r"(\d)\s*충(?=\s|$|호|동|층|\d)", r"\1층", s)
    s = re.sub(r"\s*외\s*\d+\s*필지", "", s)
    s = re.sub(r"제\s*(\d+)\s*동", r"\1동", s)
    s = re.sub(r"제\s*([가-힣A-Za-z])\s*동", r"\1동", s)
    s = re.sub(
        r"제\s*(\d+)\s*\([^)]*\)\s*층", r"\1층", s
    )  # 제1(상층하층)층 같은 복층 표기
    s = re.sub(r"제\s*(\d+)\s*층", r"\1층", s)
    s = re.sub(r"제\s*([가-힣A-Za-z]?\d{1,4})\s*호", r"\1호", s)
    s = re.sub(r"제\s*([가-힣A-Za-z])\s*(\d{3,4})\s*호", r"\1동 \2호", s)
    s = norm(s)
    # 원문에 동·호가 통째로 두 번 적힌 경우('101동 504호 101동 504호') 한 번으로 접는다.
    s = re.sub(r"((?:\d+|[A-Za-z가-힣])동\s*\d+호)(?:\s+\1)+", r"\1", s)
    # 'N동NNN호' 분리 + 호수 앞 외톨이 '동' 제거(2차 검색어도 동일 기준으로 정리).
    s = normalize_unit_dong(s)
    return norm(s)


BARE_TRAILING_NUMBER = re.compile(r"\s+\d{1,4}[A-Za-z]?\s*$")
ADDR_NUMBER = re.compile(
    r"(?:동|가|리)\s+(?:산\s*)?\d|(?:로|길)\d*(?:번길|길|로)?\s*\d"
)

DETAIL_PATTERNS = [
    re.compile(r"\s+\d{1,4}\s*호\s*$"),
    re.compile(r"\s+[가-힣A-Za-z]?\d{1,4}\s*호\s*$"),
    re.compile(r"\s+\d{1,4}\s*층\s*$"),
    re.compile(r"\s+[A-Za-z가-힣]?\s*\d{1,3}\s*동\s+\d{1,4}\s*호\s*$"),
    BARE_TRAILING_NUMBER,
]


def strip_detail(value: Any) -> str:
    s = typo_fix_first_pass(value)
    changed = True
    while changed:
        changed = False
        for pat in DETAIL_PATTERNS:
            ns = normalize_spaces(pat.sub("", s))
            if ns == s:
                continue
            # 말미의 순수 숫자는 호수 표기 없이 적힌 호실로 보고 떼지만,
            # 그 숫자가 지번/건물번호 자체라면('역삼동 123') 주소가 깨지므로 유지한다.
            if pat is BARE_TRAILING_NUMBER and not ADDR_NUMBER.search(ns):
                continue
            s = ns
            changed = True
            break
    return s


def strip_unit(value: Any) -> str:
    s = norm(value)
    s = re.sub(r"\s+[가-힣A-Za-z]?\d{1,4}\s*호\s*$", "", s)
    s = re.sub(r"\s+\d{1,4}\s*층\s*$", "", s)
    s = re.sub(r"\s+[가-힣A-Za-z]?\d{1,3}\s*동\s*$", "", s)
    return norm(s)


SIDO_RE = REGIONS_SIDO_RE

LOT_RE = re.compile(
    rf"(?:(?P<sido>{SIDO_RE})\s+)?"
    r"(?:(?P<city>[가-힣]+시)\s+)?"
    r"(?:(?P<sigungu>[가-힣]+구|[가-힣]+군)\s+)?"
    r"(?:(?P<eupmyeon>[가-힣]+(?:읍|면))\s+)?"
    r"(?P<dong>[가-힣0-9]+(?:동|가|리))\s+"
    r"(?P<lot>(?:산\s*)?\d+(?:-\d+)?)(?!\d)(?!\s*호)"
)


def find_lot(value: str) -> re.Match[str] | None:
    # 시/도가 생략된 원문('수원시 팔달구 인계동 123-4')도 지번주소로 인정하되,
    # 행정구역이 전혀 없는 매치('마장동 801')는 전국에 동명이 많아 주소로 보지 않는다.
    for m in LOT_RE.finditer(value):
        if m.group("sido") or m.group("city") or m.group("sigungu"):
            return m
    return None


def strip_building_tail_after_lot(value: Any) -> str:
    s = norm(value)
    m = find_lot(s)
    return norm(m.group(0)) if m else ""


def parse_lot_addr(*texts: Any) -> dict[str, str]:
    for text in texts:
        s = typo_fix(text)
        m = find_lot(s)
        if m:
            d = {k: (v or "") for k, v in m.groupdict().items()}
            d["addr"] = norm(
                " ".join(
                    x
                    for x in [
                        d["sido"],
                        d["city"],
                        d["sigungu"],
                        d["eupmyeon"],
                        d["dong"],
                        d["lot"],
                    ]
                    if x
                )
            )
            return d
    return {
        "sido": "",
        "city": "",
        "sigungu": "",
        "eupmyeon": "",
        "dong": "",
        "lot": "",
        "addr": "",
    }


def road_no_key(value: Any) -> str:
    s = norm(value)
    m = re.search(r"([가-힣0-9]+(?:로|길)\d*(?:번길|길|로)?\s+\d+(?:-\d+)?)", s)
    return norm(m.group(1)) if m else ""


def lot_key(value: Any) -> str:
    s = norm(value)
    m = re.search(r"([가-힣0-9]+(?:동|가|리))\s+(산\s*)?(\d+(?:-\d+)?)", s)
    if not m:
        return ""
    # Juso 지번주소는 산 지번을 '산101-1'처럼 붙여 쓰므로 같은 형태로 맞춘다.
    return norm(f"{m.group(1)} {'산' if m.group(2) else ''}{m.group(3)}")


def district_key(value: Any) -> str:
    s = norm(value)
    m = re.search(
        rf"((?:{SIDO_RE})\s+(?:[가-힣]+시(?:\s+(?:[가-힣]+구|[가-힣]+군))?|[가-힣]+구|[가-힣]+군))",
        s,
    )
    return norm(m.group(1)) if m else ""


def dong_key(value: Any) -> str:
    s = norm(value)
    m = re.search(r"([가-힣0-9]+(?:동|가|리))\s+\d", s)
    return m.group(1) if m else ""


def lot_variants(value: Any) -> list[str]:
    s = norm(value)
    out: list[str] = []
    for m in re.finditer(
        rf"((?:{SIDO_RE})\s+.+?\s+[가-힣0-9]+(?:동|가|리)\s+)(\d{{3,5}})(\b)", s
    ):
        n = m.group(2)
        splits: list[int] = []
        if len(n) == 4:
            splits = [3, 2]
        elif len(n) == 5:
            splits = [4, 3]
        elif len(n) == 3:
            splits = [2]
        for k in splits:
            out.append(norm(s[: m.start(2)] + n[:k] + "-" + n[k:] + s[m.end(2) :]))
    return out


STOP_BUILDING = {
    "제",
    "층",
    "호",
    "동",
    "외",
    "필지",
    "인천광역시",
    "서울특별시",
    "경기도",
}


def building_tokens(*texts: Any) -> list[str]:
    joined = " ".join(norm(t) for t in texts if t)
    joined = re.sub(
        r"(서울특별시|인천광역시|경기도|[가-힣]+시|[가-힣]+구|[가-힣]+군|[가-힣0-9]+동|[가-힣0-9]+가|[가-힣0-9]+리)",
        " ",
        joined,
    )
    joined = re.sub(
        r"[가-힣0-9]+(?:로|길)\d*(?:번길|길|로)?\s*\d*(?:-\d+)?", " ", joined
    )
    joined = re.sub(
        r"\b산\d+(?:-\d+)?\b|\b\d+(?:-\d+)?\b|\b\d{1,4}호\b|\b\d{1,4}층\b|\b\d{1,3}동\b",
        " ",
        joined,
    )
    seen: set[str] = set()
    out: list[str] = []
    for token in re.findall(r"[가-힣A-Za-z][가-힣A-Za-z0-9\-]{1,}", joined):
        token = token.lower()
        if token not in STOP_BUILDING and token not in seen:
            seen.add(token)
            out.append(token)
    return out[:8]


KOR_DONG_MAP = {
    "에이": "A",
    "비": "B",
    "씨": "C",
    "시": "C",
    "디": "D",
    "이": "E",
    "에프": "F",
    "지": "G",
    "에취": "H",
    "에이치": "H",
    "아이": "I",
    "제이": "J",
    "케이": "K",
}


def normalize_bld_dong(value: Any) -> str:
    x = norm(value).strip(" .-_")
    return KOR_DONG_MAP.get(x, x)


def tail_after_lot(raw: Any, lot_addr: str) -> str:
    s = typo_fix(raw)
    if lot_addr and lot_addr in s:
        return norm(s.split(lot_addr, 1)[1])
    m = find_lot(s)
    if m:
        return norm(s[m.end() :])
    return s


def is_probable_building_dong(token: Any) -> bool:
    x = norm(token).replace("동", "")
    if not x:
        return False
    if re.fullmatch(r"\d{1,4}", x):
        return True
    if re.fullmatch(r"[A-Za-z]", x):
        return True
    if re.fullmatch(r"[가-힣]", x):
        return x in set("가나다라마바사아자차카타파하")
    return x in KOR_DONG_MAP


def normalize_unit_dong(text: Any) -> str:
    """호수 앞 '동' 표기를 정리한다.

    - 'N동NNN호'(붙은 건물동+호수)는 'N동 NNN호'로 띄워 동/호를 각각 잡게 한다.
      뒤가 '\\d+호'일 때만 띄워 '성수동1가' 같은 법정동은 건드리지 않는다.
    - 식별자(숫자·문자) 없이 호수 앞에 붙은 외톨이 '동'('동401호', '-동 201호')은
      잡음이므로 삭제한다. 건물동('101동','A동','가동')·법정동('약대동')은 동 앞에
      숫자·영문·한글이 있어 룩비하인드로 보존된다.
    """
    s = re.sub(r"동(?=\d{1,4}호)", "동 ", str(text))
    s = re.sub(r"(?<![가-힣A-Za-z0-9])동(?=\s*\d{1,4}\s*호)", "", s)
    return normalize_spaces(s)


def unit_extract(raw: Any, final: Any, lot_addr: str) -> dict[str, str]:
    # If the original text is road-address only, tail_after_lot() can include legal dongs
    # such as 경서동/부평동.  Also, text like '만수동 601호' must not be treated as
    # lot 601 + trailing '호'.  Only cut after the lot when the selected lot address
    # actually appears in the raw text; otherwise use the whole raw/final text.
    search = tail_after_actual_lot(raw, lot_addr) or (typo_fix(raw) + " " + norm(final))
    search = re.sub(r"호\s*호\b", "호", search)
    search = normalize_unit_dong(search)
    bld_dong = ""
    ho = ""
    floor = ""
    m = re.search(r"제\s*([가-힣A-Za-z])\s*(\d{3,4})\s*호", search)
    if m:
        bld_dong = normalize_bld_dong(m.group(1))
        ho = m.group(2)
    if not bld_dong:
        for m in re.finditer(r"(?:제\s*)?([가-힣A-Za-z]+|\d{1,4})\s*동\b", search):
            cand = normalize_bld_dong(m.group(1))
            if is_probable_building_dong(cand):
                bld_dong = cand
                break
    hos = re.findall(r"(?:제\s*)?([가-힣A-Za-z]?\d{1,4})\s*호\b", search)
    if hos:
        ho = hos[-1]
    floors = re.findall(r"(?:제\s*)?(\d{1,3})\s*층\b", search)
    if floors:
        floor = floors[-1]
    return {"bld_dong": bld_dong, "floor": floor, "ho": ho}


def unit_out_of_range(bld_dong: Any, ho: Any) -> str:
    """동/호 추출값이 주소로 보기 어려운 범위면 사유 문자열을 돌려준다(정상은 "").

    unit_extract는 '호' 접미사가 붙은 표기만 잡으므로 흔치는 않지만, 지번·연도 등이
    호로 잘못 잡힌 경우(예: 0호, 12345호)를 검토 사유로 드러내 자동 통과를 막는다.
    실제 데이터를 잘못 내리지 않도록 명백히 비정상인 값만 표시한다.
    """
    ho_nums = re.findall(r"\d+", str(ho))
    if ho_nums:
        n = int(ho_nums[-1])
        if n == 0 or n > 9999:
            return f"호수 범위 비정상({ho})"
    dong_nums = re.findall(r"\d+", str(bld_dong))
    if dong_nums and int(dong_nums[-1]) > 9999:
        return f"동 표기 비정상({bld_dong})"
    return ""


def tail_after_actual_lot(raw: Any, lot_addr: str) -> str:
    s = typo_fix(raw)
    if lot_addr and lot_addr in s:
        return norm(s.split(lot_addr, 1)[1])
    return ""


def original_is_under_specified(raw: Any) -> bool:
    """True when the source address is only legal-dong + unit, e.g. 서울 성동구 마장동 801호.

    Such rows are not safely resolvable even if Juso returns a plausible road/lot address,
    because the query lacks a lot number, road building number, or building name.
    """
    s = typo_fix(raw)
    s = re.sub(r"^\d{5}\s+", "", s)
    has_unit = bool(re.search(r"[가-힣A-Za-z]?\d{1,4}\s*호", s))
    if not has_unit:
        return False
    # A lot number must not merely be the unit number in '마장동 801호'.
    has_lot = bool(
        re.search(
            r"[가-힣0-9]+(?:동|가|리)\s*(?:산\s*)?\d+(?:-\d+)?(?:번지)?(?!\d)(?!\s*호)",
            s,
        )
    )
    # Some source rows omit the 법정리 after 읍/면 but still provide a usable lot number.
    has_lot = has_lot or bool(
        re.search(r"[가-힣]+(?:읍|면)\s+\d+(?:-\d+)?(?!\d)(?!\s*호)", s)
    )
    has_road_no = bool(
        re.search(r"[가-힣0-9]+(?:로|길)\d*(?:번길|길|로)?\s*\d+(?:-\d+)?", s)
    )
    if has_lot or has_road_no:
        return False

    stripped = s
    stripped = re.sub(rf"(?:{SIDO_RE})", " ", stripped)
    stripped = re.sub(r"\b[가-힣]+(?:시|구|군|읍|면)\b", " ", stripped)
    # Remove standalone legal-dong/ri tokens only; do not strip building names like 가람빌리지.
    stripped = re.sub(r"(?:^|\s)[가-힣0-9]+(?:동|가|리)(?=\s|$)", " ", stripped)
    stripped = re.sub(
        r"(?:제\s*)?[가-힣A-Za-z]?\d{1,4}\s*호|(?:제\s*)?\d{1,3}\s*층|외\s*\d+\s*필지",
        " ",
        stripped,
    )
    stripped = re.sub(r"[() ,]+", " ", stripped)
    stripped = re.sub(r"\d+", " ", stripped)
    meaningful = [
        t
        for t in re.findall(r"[가-힣A-Za-z]{2,}", stripped)
        if t not in {"번지", "지상"}
    ]
    return not meaningful


def _strip_unit_detail(text: Any) -> str:
    """건물명 후보에서 동·층·호 같은 상세부를 떼어 이름 부분만 남긴다."""
    s = normalize_unit_dong(str(text))
    s = re.sub(r"(?:제\s*)?[가-힣A-Za-z0-9]+\s*동\b", " ", s)
    s = re.sub(r"(?:제\s*)?\d{1,3}\s*층\b", " ", s)
    s = re.sub(r"(?:제\s*)?[가-힣A-Za-z]?\d{1,4}\s*호\b", " ", s)
    return norm(s)


def building_name(raw: Any, lot_addr: str, juso_road: Any = "", final: Any = "") -> str:
    names: list[str] = []
    for text in [raw, final, juso_road]:
        # Keep commas here.  Parentheses often look like (경서동, 아시아드빌),
        # and replacing commas too early turns that into one bad building name.
        for par in re.findall(r"\(([^)]*)\)", normalize_spaces(text)):
            for p in [p.strip() for p in re.split(r"[,/]", par) if p.strip()]:
                # Parentheses often contain legal dong first: (경서동, 아시아드빌)
                if re.search(r"(동|가|리)$", p):
                    continue
                # ...or pure unit detail like (동 102호); 상세부를 떼고 이름만 남긴다.
                p = _strip_unit_detail(p)
                if p:
                    names.append(p)
    # Only use free text after a real lot-address match.  For road-address originals,
    # the whole address tail contains legal dong/road terms and should not become building name.
    t = tail_after_actual_lot(raw, lot_addr)
    t = re.sub(r"외\s*\d+\s*필지", " ", t)
    # 'N동NNN호'를 띄우고 외톨이 '동'을 떼어, 건물동 숫자가 건물명으로 새는 것을 막는다.
    t = normalize_unit_dong(t)
    t = re.sub(r"(?:제\s*)?[가-힣A-Za-z0-9]+\s*동\b", " ", t)
    t = re.sub(r"(?:제\s*)?\d{1,3}\s*층\b", " ", t)
    t = re.sub(r"(?:제\s*)?[가-힣A-Za-z]?\d{1,4}\s*호\b", " ", t)
    t = re.sub(r"^\d{5}\s+", " ", t)
    t = re.sub(r"[()]", " ", t)
    t = norm(t)
    if t and not re.search(r"(광역시|특별시|\d+번길|\d+로|\d+-\d+)", t):
        names.insert(0, t)
    seen: list[str] = []
    for n in names:
        n = norm(re.sub(r"[()]", " ", str(n)))
        # 한글·영문이 없는 후보(순수 숫자 '1', 기호 '-')는 건물명이 아니므로 버린다.
        if n and re.search(r"[가-힣A-Za-z]", n) and n not in seen:
            seen.append(n)
    return seen[0] if seen else ""


def suffix_dong(value: Any) -> str:
    x = norm(value)
    return x if not x or x.endswith("동") else f"{x}동"


def suffix_ho(value: Any) -> str:
    x = norm(value)
    return x if not x or x.endswith("호") else f"{x}호"


EXTRA_PARCELS = re.compile(r"외\s*\d+\s*필지")


def has_extra_parcels(value: Any) -> bool:
    return bool(EXTRA_PARCELS.search(normalize_spaces(value)))
