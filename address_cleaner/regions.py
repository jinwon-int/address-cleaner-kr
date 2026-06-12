"""행정구역 상수와 법정동 사전 — 일반/등기 정제가 공유하는 단일 출처.

시/도 명칭이 normalizer.py와 registry/normalize.py에 따로 있어
한쪽만 고치는 사고가 났던 것을 막기 위해 여기로 모은다.
"""

from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path

# 현행 시/도 17개. 정규식 대안(alternation)에서 긴 이름이 먼저 매치돼야 하므로
# '강원특별자치도'가 '강원도'보다 앞에 오도록 현행 → 구명칭 순서를 유지한다.
CURRENT_SIDO_NAMES = [
    "서울특별시", "부산광역시", "대구광역시", "인천광역시", "광주광역시",
    "대전광역시", "울산광역시", "세종특별자치시", "경기도", "강원특별자치도",
    "충청북도", "충청남도", "전북특별자치도", "전라남도", "경상북도",
    "경상남도", "제주특별자치도",
]

# 개편 전 명칭. 원주소 데이터에 여전히 흔하다.
LEGACY_SIDO_NAMES = ["강원도", "전라북도", "제주도"]

ALL_SIDO_NAMES = CURRENT_SIDO_NAMES + LEGACY_SIDO_NAMES

SIDO_RE = "|".join(ALL_SIDO_NAMES)


@dataclass(frozen=True)
class AdminDict:
    """법정동코드 전체자료 기반 오프라인 행정구역 사전.

    행정안전부 '법정동코드 전체자료' 텍스트 파일(탭 구분:
    법정동코드 / 법정동명 / 폐지여부)을 읽어, 주소에 등장한 행정구역
    조합이 실존하는지 API 호출 없이 검사한다.
    """

    ngrams: frozenset[str]

    @classmethod
    def load(cls, path: str | Path) -> "AdminDict":
        raw = Path(path).read_bytes()
        try:
            text = raw.decode("utf-8")
        except UnicodeDecodeError:
            # 공공데이터 배포본은 EUC-KR(cp949)인 경우가 많다.
            text = raw.decode("cp949")
        ngrams: set[str] = set()
        for line in text.splitlines():
            parts = line.rstrip("\n").split("\t")
            if len(parts) < 2 or not parts[0].strip().isdigit():
                continue  # 헤더/빈 줄
            if len(parts) >= 3 and "폐지" in parts[2]:
                continue
            tokens = parts[1].split()
            # '서울특별시 강남구 역삼동'에서 나올 수 있는 연속 토큰 조합 전부:
            # 시군구 생략('강남구 역삼동')이나 시군구만 적은 도로명 주소도 검사 가능.
            for i in range(len(tokens)):
                for j in range(i + 1, len(tokens) + 1):
                    ngrams.add(" ".join(tokens[i:j]))
        if not ngrams:
            raise ValueError(f"법정동 사전이 비어 있음: {path}")
        return cls(ngrams=frozenset(ngrams))

    def contains(self, combo: str) -> bool:
        return combo in self.ngrams
