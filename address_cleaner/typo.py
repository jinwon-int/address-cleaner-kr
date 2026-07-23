"""오타 교정 규칙 저장소 — 일반(normalizer)·등기(registry) 정제가 공유하는 공용 모듈.

교정 후보 수확(`excel --corrections-out`)에서 사람이 검토해 승격한 규칙을
`--typo-rules` JSON으로 양쪽 모드에 되먹이는 학습 루프의 단일 저장소다.
"""

from __future__ import annotations

import json
import re
from pathlib import Path
from typing import Any, Iterable

# registry/normalize.py의 norm()과 같은 규칙의 사본. typo_fix가 registry 모듈을
# 되부르면 순환 import가 되므로 여기 사적으로 둔다 (registry는 이 모듈을 import).
_SPECIAL_CHARS = re.compile(r"[%,=><\[\]]+")


def _norm(value: Any) -> str:
    if value is None:
        return ""
    s = re.sub(r"\s+", " ", str(value).replace("　", " ")).strip()
    s = _SPECIAL_CHARS.sub(" ", s)
    s = re.sub(r"[,，]+", " ", s)
    return re.sub(r"\s+", " ", s).strip()


# 데이터에서 실제로 발견된 오타 교정 규칙. CLI의 --typo-rules JSON으로
# 코드 수정 없이 규칙을 추가할 수 있다.
BASE_TYPO_REPLACEMENTS: list[tuple[str, str]] = [
    ("서울틀벽시", "서울특별시"),
    ("서울특벽시", "서울특별시"),
    ("서욽특별시", "서울특별시"),
    ("서울시", "서울특별시"),
    ("인천광여깃", "인천광역시"),
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


def apply_typo_replacements(text: str) -> str:
    """기본·추가 치환 규칙만 적용한다 (공백/특수문자 정규화 없음).

    일반 정제기(normalizer)는 쉼표·괄호를 자체 규칙으로 다루므로,
    norm()까지 수행하는 typo_fix 대신 이 순수 치환을 원문 정리 단계에 끼운다.
    """
    for a, b in BASE_TYPO_REPLACEMENTS + _extra_typo_replacements:
        text = text.replace(a, b)
    return text


def typo_fix(value: Any) -> str:
    s = _norm(value)
    s = re.sub(r"^\d{5}\s+", "", s)
    s = apply_typo_replacements(s)
    # 시/도 약칭은 문자열 시작에서만 확장한다. 중간 치환은 '시민로' 같은 도로명이나
    # '서울 빌라' 같은 건물명까지 훼손한다.
    s = re.sub(r"^서울\s+", "서울특별시 ", s)
    s = re.sub(r"^인천\s+", "인천광역시 ", s)
    s = re.sub(r"^경기\s+", "경기도 ", s)
    s = re.sub(r"인천광역시\s+남구\b", "인천광역시 미추홀구", s)
    # 산 지번은 '산12-3' 표기로 통일해 지번 파싱과 Juso 지번주소 대조를 일관되게 한다.
    s = re.sub(r"((?:동|가|리)\s+)산\s+(\d)", r"\1산\2", s)
    return _norm(s)
