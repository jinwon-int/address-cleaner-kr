from __future__ import annotations

from dataclasses import dataclass
import re
from typing import Any

try:
    import pandas as pd
except Exception:  # pragma: no cover - pandas is optional for plain text use.
    pd = None


PAREN_CONTENT_RE = re.compile(r"\([^)]*\)")
ZIPCODE_RE = re.compile(r"^\s*\d{5}\s+")
ET_AL_RE = re.compile(r"외\s*\d+\s*(필지|건|목록)")
DETAIL_START_RE = re.compile(
    r"\s+(?:"
    r"(?:제)?\d+동|(?:제)?[A-Za-z가-힣]동|(?:제)?\d+층|지하\s*\d*층?|"
    r"(?:제)?\d+호|[가-힣A-Za-z0-9_-]+(?:아파트|빌라|펠리스|타운|빌|맨션|하우스|오피스텔|주택|연립|다세대|상가)"
    r")\b"
)
ROAD_SUFFIXES = "대로|로|길|번길"
ROAD_QUERY_RE = re.compile(
    rf"^(?P<prefix>.+?\s[가-힣0-9·.\-]+(?:{ROAD_SUFFIXES}))\s*"
    r"(?P<num>\d+(?:-\d+)?)\b"
)
LOT_QUERY_RE = re.compile(
    r"^(?P<prefix>.+?\s[가-힣0-9]+(?:동|읍|면|리|가))\s*"
    r"(?P<num>산?\s*\d+(?:-\d+)?)\b"
)
SQL_FILTER_RE = re.compile(r"[%=><\[\]]")
SQL_WORD_RE = re.compile(
    r"\b(OR|SELECT|INSERT|DELETE|UPDATE|CREATE|DROP|EXEC|UNION|FETCH|DECLARE|TRUNCATE)\b",
    re.IGNORECASE,
)
COMMON_INVALID_MARKERS = {"주소없음", "미상", "없음", "미정", "확인중", "해당없음", "불명", "없다"}


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
    if pd is not None:
        try:
            if pd.isna(raw_addr):
                return ""
        except (TypeError, ValueError):
            pass
    text = str(raw_addr).strip()
    return "" if text.lower() == "nan" else text


def normalize_spaces(text: str) -> str:
    return re.sub(r"\s+", " ", text).strip()


def preprocess_raw_address(raw_addr: Any) -> str:
    """Clean noisy Excel/HUG address strings before query extraction."""
    text = to_addr_str(raw_addr)
    if not text:
        return ""

    text = text.replace("\t", " ")
    text = re.sub(r"[\x00-\x1f]", " ", text)
    text = ZIPCODE_RE.sub("", text).strip()

    sido_keywords = [
        "서울특별시", "부산광역시", "대구광역시", "인천광역시", "광주광역시",
        "대전광역시", "울산광역시", "세종특별자치시", "경기도", "강원특별자치도",
        "충청북도", "충청남도", "전북특별자치도", "전라남도", "경상북도",
        "경상남도", "제주특별자치도",
    ]
    for sido in sido_keywords:
        if text.count(sido) >= 2:
            text = text[text.rfind(sido):]
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

    placeholders: list[str] = []
    def protect(match: re.Match[str]) -> str:
        placeholders.append(match.group(0))
        return f"__P{len(placeholders)-1}__"

    text = PAREN_CONTENT_RE.sub(protect, text)
    text = ET_AL_RE.sub("", text)
    text = re.sub(r"(?<=\s)\d+필지(?=\s|$)", "", text)
    for i, value in enumerate(placeholders):
        text = text.replace(f"__P{i}__", value)

    text = re.sub(r"(\d+호)호", r"\1", text)
    text = re.sub(r"(\d+)번지\s*(\d+\s*호)", r"\1 \2", text)
    text = re.sub(r"(\d+)번지", r"\1", text)
    text = re.sub(r"(\d+)의(\d+)", r"\1-\2", text)

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
    head = re.sub(r"([가-힣]+(?:시|군|구))([가-힣]{2,}(?:읍|면|동|리|로|길))", r"\1 \2", head)
    head = re.sub(r"([가-힣]{2,}(?:읍|면|동|리))(\d)", r"\1 \2", head)
    head = re.sub(r"(번길)(\d)", r"\1 \2", head)
    text = head + tail

    text = re.sub(
        r"(\d+(?:-\d+)?)"
        r"(?!(?:가|나|다|라|마|바|사|아|자|차|카|타|파|하)길)"
        r"(?!단지|차|동|호|관|블록|공구|구역|지구)"
        r"([가-힣]{2,})",
        r"\1 \2",
        text,
    )
    return normalize_spaces(text)


def strip_api_unsafe_tokens(text: str) -> str:
    text = SQL_FILTER_RE.sub(" ", text)
    text = SQL_WORD_RE.sub(" ", text)
    return normalize_spaces(text)


def _cut_detail(text: str) -> str:
    text = PAREN_CONTENT_RE.sub(" ", text)
    text = normalize_spaces(text)
    match = DETAIL_START_RE.search(text)
    if match:
        text = text[: match.start()]
    return normalize_spaces(text)


def normalize_for_search(raw_addr: Any) -> NormalizedAddress:
    original = to_addr_str(raw_addr)
    cleaned = strip_api_unsafe_tokens(preprocess_raw_address(original))
    if not cleaned:
        return NormalizedAddress(original=original, query="", kind="empty", status="empty")
    if cleaned in COMMON_INVALID_MARKERS:
        return NormalizedAddress(original=original, query="", kind="invalid", status="invalid_marker")

    base = cleaned

    road_match = ROAD_QUERY_RE.match(base)
    if road_match:
        query = normalize_spaces(f"{road_match.group('prefix')} {road_match.group('num')}")
        return NormalizedAddress(original=original, query=query, kind="road", status="ok", detail=cleaned[len(query):].strip())

    lot_match = LOT_QUERY_RE.match(base)
    if lot_match:
        lot_no = re.sub(r"\s+", "", lot_match.group("num"))
        query = normalize_spaces(f"{lot_match.group('prefix')} {lot_no}")
        return NormalizedAddress(original=original, query=query, kind="lot", status="ok", detail=cleaned[len(query):].strip())

    fallback = normalize_spaces(_cut_detail(cleaned))
    status = "fallback" if fallback else "empty"
    return NormalizedAddress(original=original, query=fallback, kind="fallback", status=status)
