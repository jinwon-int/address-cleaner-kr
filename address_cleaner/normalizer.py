from __future__ import annotations

from dataclasses import dataclass
import re
from typing import Any

from .regions import ALL_SIDO_NAMES
from .typo import apply_typo_replacements


PAREN_CONTENT_RE = re.compile(r"\([^)]*\)")
ZIPCODE_RE = re.compile(r"^\s*\d{5}\s+")
ET_AL_RE = re.compile(r"외\s*\d+\s*(필지|건|목록|세대)")
FLOOR_DETAIL_RE = re.compile(
    # '제1(상층하층)층' 같은 복층 표기는 괄호가 끼어 있어도 층 정보로 보고 제거한다.
    r"(?<![가-힣0-9])(?:제\s*)?(?:지하\s*)?\d+\s*\([^)]*\)\s*층(?![가-힣0-9])"
    r"|(?<![가-힣0-9])(?:제\s*)?(?:지하\s*)?\d+\s*층(?![가-힣0-9])"
    r"|(?<![가-힣0-9])(?:제\s*)?지하\s*층(?![가-힣0-9])"
    r"|(?<![가-힣0-9])(?:지층|반지하)(?![가-힣0-9])"
)
# 등기·판결문에서 딸려 온 면적/도면 법적 기술('194.49㎡ 중 도면표시 ㄱ,ㄴ,…').
# 첫 표지부터는 주소가 아니므로 잘라내되, 안에 있던 호수는 상세부로 살린다.
LEGAL_DESC_START_RE = re.compile(r"\d+(?:\.\d+)?\s*㎡|도면\s*(?:표\s*시|범위)")
LEGAL_DESC_UNIT_RE = re.compile(r"(?:제\s*)?\d{1,4}\s*호(?![가-힣0-9])")
# juso 표준 표기에서 복사된 '(법정동[, 건물명])' 괄호 주석. 동명 반복은 버리고
# 괄호에만 있던 건물명 같은 새 정보는 평문으로 남겨 검색어에 살린다.
ANNOTATION_PAREN_RE = re.compile(
    r"\(\s*(?P<dong>[가-힣0-9]+(?:동|리|읍|면|가))\s*[,\s]?\s*(?P<extra>[^)]*?)\s*\)"
)
ANNOTATION_ECHO_RE = re.compile(r"\d*가|[가-힣0-9]*(?:동|리|읍|면)")
ROAD_HINT_PAREN_RE = re.compile(r"\(\s*도로명\s*[::][^)]*\)")
# 도시개발/택지지구의 블록·로트 표기는 juso가 모르는 잡음이라 통째로 제거한다.
# 실제 건물명('동탄2신도시…33단지')과 헷갈리지 않게 숫자를 품은 토큰만 지운다.
DEV_DISTRICT_TOKEN_RE = re.compile(
    r"(?<![가-힣])(?:[가-힣0-9-]*\d[가-힣0-9-]*(?:지구|구역|블록|블럭|로트|공구)"
    r"|도시개발(?:사업)?(?:구역|지구)|택지개발(?:예정)?지구)(?=\s|$)"
)
DEV_DISTRICT_PREFIX_RE = re.compile(
    r"(?<![가-힣])[가-힣0-9-]*\d[가-힣0-9-]*?(?:블록|블럭|로트|공구)(?=[가-힣])"
)
DETAIL_START_RE = re.compile(
    r"\s+(?:"
    r"(?:제)?\d+동|(?:제)?[A-Za-z가-힣]동|(?:제)?\d+층|지하\s*\d*층?|"
    r"(?:제)?\d+호|[가-힣A-Za-z0-9_-]+(?:아파트|빌라|펠리스|타운|빌|맨션|하우스|오피스텔|주택|연립|다세대|상가)"
    r")\b"
)
ROAD_SUFFIXES = "번길|대로|로|길"
# 도로명은 시군구가 생략된 원주소('테헤란로 152 ...')도 검색어로 살린다.
# 단일 여부는 API 검증(2건이상검색)으로 가린다. 지번(LOT)은 동명이 전국에
# 많아 행정구역 없는 매치를 주소로 보지 않는 기존 정책을 유지한다.
ROAD_QUERY_RE = re.compile(
    rf"^(?P<prefix>(?:.+?\s)?[가-힣0-9·.\-]+(?:{ROAD_SUFFIXES}))\s*"
    r"(?P<num>\d+(?:-\d+)?)\b"
)
LOT_QUERY_RE = re.compile(
    r"^(?P<prefix>.+?\s[가-힣0-9]+(?:동|읍|면|리|가))\s*"
    r"(?P<num>산?\s*\d+(?:-\d+)?)\b"
)
# 지번이 없는 원주소용 행정구역 접두: 마지막 동/읍/면/리/가 토큰까지 잡는다.
# '비동'(단독 식별동)이나 '204동'(건물동)이 법정동으로 오인되지 않게, 연장
# 토큰은 한글로 시작하고 두 글자 이상인 것만 허용한다.
BUILDING_PREFIX_RE = re.compile(
    r"^(?P<prefix>.+?\s(?!제\s*\d)[가-힣][가-힣0-9]*(?:동|읍|면|리|가)"
    r"(?:\s(?!제\s*\d)[가-힣][가-힣0-9]+(?:동|읍|면|리|가)(?=\s|$))*)(?=\s|$)"
)
# 동/호/층 같은 상세 토큰. 건물명 후보에서 걸러낸다.
BUILDING_UNIT_TOKEN_RE = re.compile(
    r"^(?:제\s*)?(?:\d+동|[A-Za-z가-힣]동|[가-힣A-Za-z]*\d{1,4}호|\d+(?:-\d+)?)$"
)
BUILDING_NAME_TOKEN_RE = re.compile(r"^[가-힣A-Za-z0-9·.\-()]{2,}$")
HANGUL_RE = re.compile(r"[가-힣]")
ADDRESS_TOKEN_RE = re.compile(
    r"(?:특별시|광역시|특별자치시|특별자치도|도|시|군|구|읍|면|동|리|가|대로|번길|로|길)"
)
ADDRESS_NUMBER_RE = re.compile(r"(?:산\s*)?\d+(?:-\d+)?")
SQL_FILTER_RE = re.compile(r"[%=><\[\]]")
SQL_WORD_RE = re.compile(
    r"\b(OR|SELECT|INSERT|DELETE|UPDATE|CREATE|DROP|EXEC|UNION|FETCH|DECLARE|TRUNCATE)\b",
    re.IGNORECASE,
)
COMMON_INVALID_MARKERS = {
    "주소없음",
    "미상",
    "없음",
    "미정",
    "확인중",
    "해당없음",
    "불명",
    "없다",
}

ORPHAN_UNIT_DONG_RE = re.compile(
    r"(?<![가-힣A-Za-z0-9])동(?=\s*(?:제\s*)?\d{1,4}\s*호\b)"
)


@dataclass(frozen=True)
class NormalizedAddress:
    original: str
    query: str
    kind: str
    status: str
    detail: str = ""

    @property
    def searchable(self) -> bool:
        return self.status == "ok" and bool(self.query)


def to_addr_str(raw_addr: Any) -> str:
    if raw_addr is None:
        return ""
    if isinstance(raw_addr, float) and raw_addr != raw_addr:  # NaN (pandas 셀 빈값)
        return ""
    text = str(raw_addr).strip()
    return "" if text.lower() == "nan" else text


def normalize_spaces(text: str) -> str:
    return re.sub(r"\s+", " ", text).strip()


def normalize_unit_dong(text: Any) -> str:
    """Normalize building-dong notation immediately before a numbered unit.

    A bare ``동`` has no building identifier, so remove it from forms such as
    ``동401호``, ``동 401호``, and ``동 제401호``.  Identified building dongs
    (``101동``, ``A동``) and legal-dong names (``송내동``) are preserved.
    Duplicated unit phrases (``제비동동402호``, ``907호 동907호``) collapse to
    a single notation.
    """
    # '제비동동402호'처럼 식별동 뒤에 중복된 동 표기는 하나로 합친다.
    value = re.sub(
        r"([가-힣A-Za-z0-9]동)\s*동(?=\s*(?:제\s*)?\d{1,4}\s*호\b)", r"\1 ", str(text)
    )
    value = re.sub(r"동(?=\d{1,4}\s*호\b)", "동 ", value)
    value = ORPHAN_UNIT_DONG_RE.sub("", value)
    value = normalize_spaces(value)
    # 같은 호 표기가 연달아 반복되면('404호 404호') 첫 표기만 남긴다.
    value = re.sub(r"(\S*\d{1,4}호)(?:\s+\1)+(?=\s|$)", r"\1", value)
    # '제비동 402호 … 제비동 402호'처럼 떨어져 반복된 동·호 구절도 정리한다.
    value = re.sub(
        r"(?P<pair>\S+\s\S*\d{1,4}호)(?P<mid>(?:\s\S+)*?)\s(?P=pair)(?=\s|$)",
        r"\g<pair>\g<mid>",
        value,
    )
    # '604호 1동 604호'처럼 같은 호가 동 식별과 함께 반복되면 식별 표기만 남긴다.
    value = re.sub(r"(\S*\d{1,4}호)\s+((?:제\s*)?\S*동)\s+\1(?=\s|$)", r"\2 \1", value)
    return normalize_spaces(value)


def _strip_legal_description(text: str) -> str:
    """면적/도면 법적 기술을 잘라내고, 그 안에 있던 첫 호수만 상세부로 살린다."""
    match = LEGAL_DESC_START_RE.search(text)
    if not match:
        return text
    head, tail = text[: match.start()], text[match.start() :]
    unit = LEGAL_DESC_UNIT_RE.search(tail)
    if unit:
        head = f"{head} {unit.group(0)}"
    return normalize_spaces(head)


def _fold_annotation_parens(text: str) -> str:
    """'(법정동[, 건물명])' 괄호 주석을 정리한다.

    동명 반복('(주안동)', '(영등포동2가)')은 지우고, 괄호에만 있던 건물명
    ('(화곡동, 타운캐슬)')은 평문으로 남겨 검색어에서 살아남게 한다.
    """
    text = ROAD_HINT_PAREN_RE.sub(" ", text)

    def fold(match: re.Match[str]) -> str:
        extra = normalize_spaces(match.group("extra"))
        if not extra or ANNOTATION_ECHO_RE.fullmatch(extra):
            return " "
        rest = text[: match.start()] + text[match.end() :]
        # 괄호 밖에 이미 있는 건물명이면 중복 주입하지 않는다.
        return " " if extra in rest else f" {extra} "

    return normalize_spaces(ANNOTATION_PAREN_RE.sub(fold, text))


def preprocess_raw_address(raw_addr: Any) -> str:
    """Clean noisy Excel/HUG address strings before query extraction."""
    text = to_addr_str(raw_addr)
    if not text:
        return ""

    text = text.replace("\t", " ")
    text = re.sub(r"[\x00-\x1f]", " ", text)
    text = ZIPCODE_RE.sub("", text).strip()
    # 수확(--corrections-out)→검토→승격된 오타 규칙을 등기 모드와 같은 지점(원문
    # 정리 단계)에서 적용해, ROAD/LOT 골격 매칭 전에 오타가 교정되게 한다.
    text = apply_typo_replacements(text)

    # 시/도 명칭은 regions.py가 단일 출처다 (구명칭 포함).
    sido_keywords = ALL_SIDO_NAMES
    for sido in sido_keywords:
        if text.count(sido) >= 2:
            text = text[text.rfind(sido) :]
            break

    chars: list[str] = []
    depth = 0
    for ch in text:
        if ch == "(":
            depth += 1
        elif ch == ")":
            depth = max(0, depth - 1)
        chars.append(" " if ch == "," and depth == 0 else ch)
    text = "".join(chars)

    text = _strip_legal_description(text)
    text = _fold_annotation_parens(text)

    placeholders: list[str] = []

    def protect(match: re.Match[str]) -> str:
        placeholders.append(match.group(0))
        return f"__P{len(placeholders) - 1}__"

    text = PAREN_CONTENT_RE.sub(protect, text)
    text = ET_AL_RE.sub("", text)
    text = re.sub(r"(?<=\s)\d+필지(?=\s|$)", "", text)
    for i, value in enumerate(placeholders):
        text = text.replace(f"__P{i}__", value)

    text = re.sub(r"(\d+호)호", r"\1", text)
    text = re.sub(r"(\d+)번지\s*(\d+\s*호)", r"\1 \2", text)
    text = re.sub(r"(\d+)번지", r"\1", text)
    text = re.sub(r"(\d+)의(\d+)", r"\1-\2", text)

    # 실전 실패 사례에서 나온 깨진 동·층·호 표기 복원 규칙들.
    text = re.sub(r"(제\s*\d+)등(?![가-힣])", r"\1동", text)  # 제2등 → 제2동 오타
    # '비제2층'처럼 식별동 글자가 층 표기에 붙은 경우: 동 정보는 살리고 층만 뗀다.
    text = re.sub(r"(?<![가-힣0-9])(비|에이|씨|디)(?=제\s*\d+\s*층)", r"\1동 ", text)
    # '오피스텔410호'의 군더더기 어휘 제거. 건물명 끝의 '~오피스텔'은 보존한다.
    text = re.sub(r"(?<![가-힣])(?:제\s*)?오피스텔(?=\s*\d{1,4}\s*호)", " ", text)
    text = re.sub(r"(?<![가-힣])제?(비|에이|씨|디)-(?=\d{1,4}\s*호)", r"\1동 ", text)
    text = re.sub(r"--\s*(\d{1,4})(?=\s|$)", r" \1호", text)  # '--408' 꼬리 호수
    text = re.sub(r"(?<=\d)\s*-동(?=\s|$)", " ", text)  # '714-14-동' 고아 동
    text = re.sub(r"(?<![가-힣0-9])\d+\s*층동(?=\s|$)", " ", text)  # '4층동' 잔재
    text = re.sub(r"(?<=[0-9])[ㅏ-ㅣ]+", "", text)  # '제1ㅣ동' 같은 모음 자모 잔재

    # Common missing-space repairs from the legacy script.
    text = re.sub(r"(특별시|광역시|특별자치시|특별자치도)(?=[가-힣])", r"\1 ", text)
    text = re.sub(
        r"^(인천|서울|경기|부산|대구|광주|대전|울산|세종|경북|경남|충북|충남|전북|전남|강원|제주)"
        r"(?!특별|광역|도)(?=[가-힣])",
        r"\1 ",
        text,
    )
    text = re.sub(r"^([가-힣]+도)(?=[가-힣])", r"\1 ", text)
    head, tail = text[:80], text[80:]
    head = re.sub(r"([가-힣]{2,}시)([가-힣]{2,}구)(?=[가-힣\s]|$)", r"\1 \2", head)
    # 붙은 행정구역 복원. 리/로/길은 뒤에 한글이 이어지면 '하이파크시티일산
    # 파밀리에…' 같은 건물명 내부일 수 있어 가르지 않는다 (동/읍/면 뒤에 붙은
    # 건물명은 아래 시군구-법정동 규칙이 마저 가른다).
    head = re.sub(
        r"([가-힣]+(?:시|군|구))"
        r"([가-힣]{2,}(?:읍|면|동)|[가-힣]{2,}(?:리|로|길)(?![가-힣]))",
        r"\1 \2",
        head,
    )
    # '주안동정다운파크빌'처럼 시군구 바로 뒤에서 법정동과 건물명이 붙은 경우.
    head = re.sub(
        r"([가-힣]+(?:시|군|구)\s)([가-힣]{2,}동)(?=[가-힣]{3,})", r"\1\2 ", head
    )
    # '양평동4가' 같은 서수 법정동은 한 토큰이므로 동과 숫자를 가르지 않는다.
    head = re.sub(
        r"([가-힣]{2,}(?:읍|면|동|리))(?!\d+가(?![가-힣]))(\d)", r"\1 \2", head
    )
    head = re.sub(r"([가-힣]{2,}동\d+가)(?=\d)", r"\1 ", head)
    head = re.sub(r"(번길)(\d)", r"\1 \2", head)
    text = head + tail

    text = DEV_DISTRICT_PREFIX_RE.sub("", text)
    text = DEV_DISTRICT_TOKEN_RE.sub("", text)

    text = re.sub(
        r"(\d+(?:-\d+)?)"
        r"(?!(?:가|나|다|라|마|바|사|아|자|차|카|타|파|하)길)"
        r"(?!단지|차|동|호|관|블록|블럭|공구|구역|지구|번길|로트|신도시|세대)"
        r"([가-힣]{2,})",
        r"\1 \2",
        text,
    )
    return normalize_unit_dong(text)


def strip_api_unsafe_tokens(text: str) -> str:
    text = SQL_FILTER_RE.sub(" ", text)
    text = SQL_WORD_RE.sub(" ", text)
    return normalize_spaces(text)


def strip_floor_detail(text: str) -> str:
    """Remove floor-only detail while preserving building, dong, and ho detail."""
    return normalize_spaces(FLOOR_DETAIL_RE.sub(" ", text))


def _cut_detail(text: str) -> str:
    text = PAREN_CONTENT_RE.sub(" ", text)
    text = normalize_spaces(text)
    match = DETAIL_START_RE.search(text)
    if match:
        text = text[: match.start()]
    return normalize_spaces(text)


def _building_parts(cleaned: str) -> tuple[str, str, list[str]] | None:
    """지번 없는 주소에서 (행정구역 접두, 건물명, 동·호 토큰들)을 뽑는다.

    juso는 '수유동 한양아이빌'처럼 법정동+건물명으로도 검색되므로, 지번/도로명
    골격이 없어도 건물명이 있으면 검색어를 포기하지 않기 위한 마지막 폴백이다.
    동·호 토큰만 있고 건물명이 없으면(예: '간석동 503호') None을 돌려주고
    기존처럼 검색주소없음으로 남긴다.
    """
    match = BUILDING_PREFIX_RE.match(cleaned)
    if not match:
        return None
    prefix = match.group("prefix")
    name_tokens: list[str] = []
    unit_tokens: list[str] = []
    for token in cleaned[match.end() :].split():
        if BUILDING_UNIT_TOKEN_RE.match(token):
            unit_tokens.append(token)
        elif not BUILDING_NAME_TOKEN_RE.match(token) or not HANGUL_RE.search(token):
            continue  # 건물명이 될 수 없는 잔재 토큰은 버린다
        elif not unit_tokens or not name_tokens:
            # 상세 토큰 전의 연속 구간, 또는 괄호에서 승계돼 상세 뒤에 남은
            # 건물명(이름이 아직 없을 때)만 채택한다. 그 밖의 잔여 토큰은
            # 중복 표기로 보고 버린다.
            name_tokens.append(token)
    if not name_tokens or not unit_tokens:
        return None
    return prefix, " ".join(name_tokens), unit_tokens


def normalize_for_search(raw_addr: Any) -> NormalizedAddress:
    original = to_addr_str(raw_addr)
    cleaned = strip_floor_detail(
        strip_api_unsafe_tokens(preprocess_raw_address(original))
    )
    if not cleaned:
        return NormalizedAddress(
            original=original, query="", kind="empty", status="empty"
        )
    if cleaned in COMMON_INVALID_MARKERS:
        return NormalizedAddress(
            original=original, query="", kind="invalid", status="invalid_marker"
        )
    if _looks_malformed(cleaned):
        return NormalizedAddress(
            original=original,
            query="",
            kind="invalid",
            status="malformed",
            detail=cleaned,
        )

    base = cleaned

    road_match = ROAD_QUERY_RE.match(base)
    if road_match:
        return NormalizedAddress(
            original=original,
            query=cleaned,
            kind="road",
            status="ok",
            detail=cleaned[road_match.end() :].strip(),
        )

    lot_match = LOT_QUERY_RE.match(base)
    if lot_match:
        return NormalizedAddress(
            original=original,
            query=cleaned,
            kind="lot",
            status="ok",
            detail=cleaned[lot_match.end() :].strip(),
        )

    # 지번/도로명 골격은 없지만 '수유동 한양아이빌 402호'처럼 법정동+건물명이
    # 있으면 건물명 검색어로 살린다. juso가 건물명 검색을 지원하므로
    # 검색주소없음으로 버리는 것보다 확정 가능성이 높다.
    building = _building_parts(cleaned)
    if building is not None:
        prefix, name, units = building
        return NormalizedAddress(
            original=original,
            query=normalize_spaces(" ".join([prefix, name, *units])),
            kind="building",
            status="ok",
            detail=" ".join(units),
        )

    fallback = normalize_spaces(_cut_detail(cleaned))
    return NormalizedAddress(
        original=original,
        query="",
        kind="invalid",
        status="unrecognized",
        detail=fallback,
    )


def base_for_search(query: str, kind: str) -> str:
    """건물명/동/호 상세부를 뗀 주소 골격(시도~지번/건물번호).

    상세 포함 검색이 0건일 때 골격만으로 2차 검색해, 상세 표기 때문에
    멀쩡한 주소가 검색주소없음으로 빠지는 것을 막는다. 건물명 검색어(kind
    ``building``)의 골격은 행정구역+건물명이다.
    """
    if kind == "building":
        parts = _building_parts(query)
        return "" if parts is None else normalize_spaces(f"{parts[0]} {parts[1]}")
    pattern = ROAD_QUERY_RE if kind == "road" else LOT_QUERY_RE
    match = pattern.match(query)
    return normalize_spaces(match.group(0)) if match else ""


def compact_for_epost(query: str, kind: str) -> str:
    """Build the short Korea Post query form for provider-specific fallback."""
    cleaned = strip_floor_detail(strip_api_unsafe_tokens(preprocess_raw_address(query)))
    if not cleaned:
        return ""
    if kind == "road":
        match = ROAD_QUERY_RE.match(cleaned)
        if not match:
            return ""
        road_name = normalize_spaces(match.group("prefix")).split()[-1]
        return normalize_spaces(f"{road_name} {match.group('num')}")
    if kind == "lot":
        match = LOT_QUERY_RE.match(cleaned)
        if not match:
            return ""
        lot_area = normalize_spaces(match.group("prefix")).split()[-1]
        lot_no = re.sub(r"\s+", "", match.group("num"))
        return normalize_spaces(f"{lot_area} {lot_no}")
    return ""


def _looks_malformed(text: str) -> bool:
    if len(text) < 5:
        return True
    if not HANGUL_RE.search(text):
        return True
    if not ADDRESS_TOKEN_RE.search(text):
        return True
    if not ADDRESS_NUMBER_RE.search(text):
        return True
    return False
