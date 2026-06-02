# -*- coding: utf-8 -*-
import os
import re
import time
import json
import hashlib
import requests
import urllib3
import pandas as pd
import logging
from datetime import datetime
from typing import Optional, Tuple, Dict, List
from functools import lru_cache
from difflib import SequenceMatcher
from urllib3.util.retry import Retry
from requests.adapters import HTTPAdapter

# ===== SSL 경고 숨기기 =====
urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

# ===== 로깅 설정 =====
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler(f'address_processing_{datetime.now():%Y%m%d_%H%M%S}.log', encoding='utf-8'),
        logging.StreamHandler()
    ]
)
logger = logging.getLogger(__name__)

# ==============================================================================
# [설정] API 키 및 환경 설정
# ==============================================================================
VWORLD_KEY = os.getenv("VWORLD_KEY", "")
JUSO_API_KEY = os.getenv("JUSO_CONFIRM_KEY", "")

FOLDER = r"C:\Users\user\Desktop\강제경매 사전심사\취합"
TARGET_COL_INDEX = 8
STATS_FILE_PATH = r"C:\Users\user\Desktop\강제경매 사전심사\낙찰가율 통계\CourtAuction_Sigungu_NakchalRate.xlsx"

API_DELAY = 0.05
MAX_RETRIES = 3
TIMEOUT_SEC = 3
MAX_WORKERS = 5

FILTER_OPTS = {
    "KEYWORD": True, "PARENTHESES": True, "UNDER_100": True,
    "NO_HO": True, "ONLY_FLOOR": True, "BASEMENT": True,
    "MOUNTAIN": True, "SHARE": True, "NO_HINT": True,
    "NO_BUNJI": True
}

# ==============================================================================
# [정규식 정의]
# ==============================================================================
PAREN_CONTENT_RE = re.compile(r'\([^)]*\)')
ZIPCODE_RE = re.compile(r'^\s*\d{5}\s+')
ET_AL_RE = re.compile(r'외\s*\d+\s*(필지|건|목록)')

HO_RE = re.compile(r'(?:제)?(?P<ho>\d+)호\b')
BUILDING_DONG_ALNUM_RE = re.compile(r'\b(?:제)?(?P<dong>(\d+동|[A-Za-z]동))\b')
BUILDING_DONG_KO_RE = re.compile(r'\b(?:제)?(?P<dong>(?:가|나|다|라|마|바|사|아|자|차|카|타|파|하|에이|비|씨|디)동)\b')
FLOOR_RE = re.compile(r'(지하|제)?\s*\d+층')
MIN_LOCATION_TOKEN_RE = re.compile(r'(시|군|구|읍|면|동|리|가|길|로)')

COMPLEX_DONG_HO_PATTERNS = [
    re.compile(r'(?P<dong>\d+)동\s*(?P<ho>\d+)호'),
    re.compile(r'(?P<dong>[가-힣])동\s*(?P<ho>\d+)호'),
    re.compile(r'(?P<dong>\d+)-(?P<ho>\d+)호'),
    re.compile(r'(?P<dong>[A-Z])(?P<ho>\d+)호'),
]

BAD_PREFIX_DONG_RE = re.compile(r'(?<![0-9a-zA-Z가-힣])동\b')
SEDAE_RE = re.compile(r'\b\S*세대\b')

COMMON_INVALID_MARKERS = ["주소없음", "미상", "없음", "미정", "확인중", "해당없음", "불명", "없다"]

# ★★★ [변경2] BASEMENT_PATTERNS는 더 이상 사용하지 않음 (check_basement 내부에서 직접 처리) ★★★
BASEMENT_PATTERNS = [
    re.compile(r'지하\s*\d*층?', re.IGNORECASE),
    re.compile(r'지층', re.IGNORECASE),
    re.compile(r'반지하', re.IGNORECASE),
]

STATS_1Y: Dict[Tuple[str, str], float] = {}
STATS_3M: Dict[Tuple[str, str], float] = {}

session = requests.Session()
retry_strategy = Retry(
    total=MAX_RETRIES, backoff_factor=0.3,
    status_forcelist=[429, 500, 502, 503, 504],
    allowed_methods=["GET"], raise_on_status=False,
)
adapter = HTTPAdapter(max_retries=retry_strategy, pool_connections=20, pool_maxsize=20)
session.mount('https://', adapter)
session.mount('http://', adapter)


# ==============================================================================
# [개선] NaN 처리 유틸리티
# ==============================================================================
def to_addr_str(raw_addr) -> str:
    if raw_addr is None:
        return ""
    try:
        if pd.isna(raw_addr):
            return ""
    except (TypeError, ValueError):
        pass
    s = str(raw_addr).strip()
    return "" if s.lower() == "nan" else s


# ==============================================================================
# ★★★ [변경1] 전처리 함수 추가 ★★★
# ==============================================================================
def preprocess_raw_address(addr_str: str) -> str:
    """
    API 호출 전 원본 주소 전처리
    - 우편번호/탭 제거 (~441건)
    - 반복 주소 제거 (~11건)
    - 콤마 정리 (~336건)
    - 외N필지 제거
    - 중복 동호 제거 (~113건)
    - 공백 없는 주소 분리 (~561건)
    순서: 기본정리 → 콤마 → 필지 → 중복동호 → 기타패턴 → 공백삽입(마지막)
    """
    if not addr_str:
        return ""

    s = addr_str

    # ═══ 1. 기본 정리 ═══
    s = s.replace('\t', ' ')
    s = re.sub(r'[\x00-\x1f]', ' ', s)
    s = re.sub(r'^\s*\d{5}\s+', '', s)  # 우편번호
    s = s.strip()

    # ═══ 2. 반복 주소 제거 ═══
    sido_keywords = [
        '서울특별시', '부산광역시', '대구광역시', '인천광역시',
        '광주광역시', '대전광역시', '울산광역시', '세종특별자치시',
        '경기도', '강원특별자치도', '충청북도', '충청남도',
        '전북특별자치도', '전라남도', '경상북도', '경상남도', '제주특별자치도'
    ]
    for sido in sido_keywords:
        if s.count(sido) >= 2:
            s = s[s.rfind(sido):]
            break

    # ═══ 3. 콤마 정리 (괄호 밖만 공백으로) ═══
    result_chars, depth = [], 0
    for ch in s:
        if ch == '(':
            depth += 1
        elif ch == ')':
            depth = max(0, depth - 1)
        if ch == ',' and depth == 0:
            result_chars.append(' ')
        else:
            result_chars.append(ch)
    s = ''.join(result_chars)

    # ═══ 4. "외 N필지" 제거 (괄호 보호) ═══
    paren_parts = re.findall(r'\([^)]*\)', s)
    for i, pp in enumerate(paren_parts):
        s = s.replace(pp, f'__P{i}__', 1)
    s = re.sub(r'외\s*\d+\s*필지', '', s)
    s = re.sub(r'(?<=\s)\d+필지(?=\s|$)', '', s)
    for i, pp in enumerate(paren_parts):
        s = s.replace(f'__P{i}__', pp, 1)

    # ═══ 5. 중복 동호 제거 (공백삽입 전!) ═══
    s = re.sub(r'(\d+동\s*\d+호)\s+\1', r'\1', s)
    s = re.sub(r'([A-Za-z가-힣]동\s+\d+호)\s+\1', r'\1', s)
    s = re.sub(r'\s+\d+층동\s+\d+호호?\s*$', '', s)
    s = re.sub(r'\s+[가-힣A-Za-z]+동\d+호\s*$', '', s)
    s = re.sub(r'\s+동\d+호\s*$', '', s)
    s = re.sub(r'\s+--\d+\s*$', '', s)
    s = re.sub(r'\s+\d+-\d+\s*$', '', s)
    s = re.sub(r'\s+0-\d+\s*$', '', s)
    match_dh = re.findall(r'(\d+동)\s*(\d+호)', s)
    if len(match_dh) >= 2 and match_dh[-1] == match_dh[-2]:
        parts = list(re.finditer(f"{match_dh[-1][0]}\\s*{match_dh[-1][1]}", s))
        if len(parts) >= 2:
            s = s[:parts[-1].start()].rstrip()

    # ═══ 6. 기타 오류 패턴 ═══
    s = re.sub(r'(\d+호)호', r'\1', s)       # 503호호 → 503호
    s = re.sub(r'제아파트(\d+호)', r'\1', s)  # 제아파트603호 → 603호
    # [PATCH-6] '번지' 토큰 정규화 (공백삽입 7-4 이전에 처리)
    s = re.sub(r'(\d+)번지\s*(\d+\s*호)', r'\1 \2', s)  # 825번지15호 → 825 15호
    s = re.sub(r'(\d+)번지', r'\1', s)                  # 1234번지 → 1234

    # ═══ 7. 공백 삽입 (가장 마지막!) ═══
    # 7-1. 시도 풀네임
    s = re.sub(r'(특별시|광역시|특별자치시|특별자치도)(?=[가-힣])', r'\1 ', s)
    # 7-2. 약칭 시도 (문자열 시작, 풀네임 뒤가 아닌 경우만)
    s = re.sub(
        r'^(인천|서울|경기|부산|대구|광주|대전|울산|세종|경북|경남|충북|충남|전북|전남|강원|제주)(?!특별|광역|도)(?=[가-힣])',
        r'\1 ', s
    )
    # 7-3. 행정구역 경계 (앞 80자만, 건물명 보호)
    head, tail = s[:80], s[80:]
    # [PATCH-1] 일반시 + 자치구 경계 분리 (수원시영통구 → 수원시 영통구)
    head = re.sub(r'([가-힣]{2,}시)([가-힣]{2,}구)(?=[가-힣\s]|$)', r'\1 \2', head)
    head = re.sub(r'([가-힣]+(?:시|군|구))([가-힣]{2,}(?:읍|면|동|리|로|길))', r'\1 \2', head)
    head = re.sub(r'([가-힣]{2,}(?:읍|면|동|리))(\d)', r'\1 \2', head)
    head = re.sub(r'(\d가)(\d)', r'\1 \2', head)
    head = re.sub(r'(번길)(\d)', r'\1 \2', head)
    s = head + tail
    # 7-4. 번지+건물명 (도로명접미 제외 + 단지/차/관/동/호/블록 등 건물명 토큰 보호)
    s = re.sub(
        r'(\d+(?:-\d+)?)'
        r'(?!(?:가|나|다|라|마|바|사|아|자|차|카|타|파|하)길)'
        r'(?!단지|차|동|호|관|블록|공구|구역|지구)'
        r'([가-힣]{2,})', r'\1 \2', s)

    # ═══ 최종 ═══
    s = re.sub(r'\s+', ' ', s).strip()
    return s


# ==============================================================================
# [유사도 매칭 시스템]
# ==============================================================================
class SimilarityMatcher:
    def __init__(self, threshold=0.85):
        self.threshold = threshold
        self.success_history = []
    
    def add_success(self, addr: str, result: dict):
        self.success_history.append((addr, result))
        if len(self.success_history) > 1000:
            self.success_history = self.success_history[-1000:]
    
    def find_similar(self, target: str):
        best_match = None
        best_score = 0
        for candidate, result in self.success_history:
            score = SequenceMatcher(None, target, candidate).ratio()
            if score > best_score and score >= self.threshold:
                best_score = score
                best_match = result
        return best_match, best_score

similarity_matcher = SimilarityMatcher()


# ==============================================================================
# [처리 모니터]
# ==============================================================================
class ProcessingMonitor:
    def __init__(self):
        self.stats = {
            'total': 0, 'success': 0, 'failed': 0,
            'vworld': 0, 'juso': 0, 'grammar': 0,
            'deleted': 0, 'similar': 0,
            'building': 0, 'cross_validated': 0, 'low_confidence': 0
        }
        self.confidence_scores = []
        self.start_time = time.time()
    
    def update(self, result: dict):
        self.stats['total'] += 1
        status = result.get('상태', '')
        confidence = result.get('신뢰도', 0)
        if '유사' in status:
            self.stats['similar'] += 1; self.stats['success'] += 1
        elif '삭제' in status:
            self.stats['deleted'] += 1
        elif 'VWorld' in status:
            self.stats['vworld'] += 1; self.stats['success'] += 1
        elif 'Juso' in status:
            self.stats['juso'] += 1; self.stats['success'] += 1
        elif '건물명' in status:
            self.stats['building'] += 1; self.stats['success'] += 1
        elif '문법' in status:
            self.stats['grammar'] += 1; self.stats['success'] += 1
        else:
            self.stats['failed'] += 1
        if result.get('교차검증') == 'Y':
            self.stats['cross_validated'] += 1
        if confidence > 0:
            self.confidence_scores.append(confidence)
            if confidence < 60:
                self.stats['low_confidence'] += 1
    
    def print_progress(self, clear=True):
        elapsed = time.time() - self.start_time
        rate = self.stats['total'] / elapsed if elapsed > 0 else 0
        progress_text = (
            f"\r진행: {self.stats['total']}건 | "
            f"성공: {self.stats['success']} | "
            f"삭제: {self.stats['deleted']} | "
            f"속도: {rate:.1f}건/초 | "
            f"경과: {elapsed:.0f}초"
        )
        if clear:
            print(progress_text, end='', flush=True)
        else:
            print(progress_text)
    
    def print_summary(self):
        print("\n" + "=" * 70)
        print(f" [처리 완료] 총 {self.stats['total']}건")
        print("=" * 70)
        print(f" - 성공: {self.stats['success']}건")
        print(f"   └ VWorld: {self.stats['vworld']}건")
        print(f"   └ Juso API: {self.stats['juso']}건")
        print(f"   └ 건물명검색: {self.stats['building']}건")
        print(f"   └ 문법분석: {self.stats['grammar']}건")
        print(f"   └ 유사매칭: {self.stats['similar']}건")
        print(f" - 삭제: {self.stats['deleted']}건")
        print(f" - 실패: {self.stats['failed']}건")
        if self.confidence_scores:
            avg_confidence = sum(self.confidence_scores) / len(self.confidence_scores)
            print(f"\n [신뢰도 분석]")
            print(f" - 평균 신뢰도: {avg_confidence:.1f}점")
            print(f" - 교차검증: {self.stats['cross_validated']}건")
            print(f" - 낮은신뢰도(<60): {self.stats['low_confidence']}건")
        elapsed = time.time() - self.start_time
        print(f"\n - 처리시간: {elapsed:.1f}초")
        print(f" - 평균속도: {self.stats['total']/elapsed:.2f}건/초")
        print("=" * 70)


# ==============================================================================
# [토큰 우선순위]
# ==============================================================================
ADDRESS_ORDER = {
    '특별시': 1, '광역시': 1, '도': 1, '특별자치시': 1, '특별자치도': 1,
    '시': 2, '군': 2, '구': 3,
    '읍': 4, '면': 4, '동': 4, '리': 4, '가': 4,
    '로': 5, '길': 5,
}

def get_token_order(token: str) -> int:
    for suffix, order in ADDRESS_ORDER.items():
        if token.endswith(suffix):
            return order
    if re.match(r'^(산)?\d+(-\d+)?$', token):
        return 6
    return 99


# ==============================================================================
# [Smart Cleaning & Utils]
# ==============================================================================
def remove_repeated_phrase(text: str) -> str:
    if not text:
        return ""
    tokens = text.split()
    if not tokens:
        return ""
    full_str = text.replace(" ", "")
    mid = len(full_str) // 2
    if full_str[:mid] == full_str[mid:]:
        half_token_idx = len(tokens) // 2
        part1 = "".join(tokens[:half_token_idx])
        part2 = "".join(tokens[half_token_idx:])
        if part1 == part2:
            return " ".join(tokens[:half_token_idx])
    match = re.match(r'^(.+?)\s+\1$', text)
    if match:
        return match.group(1)
    return text


def fix_admin_typos(addr_str: str) -> str:
    if not addr_str:
        return ""
    addr_str = re.sub(r'(\d+)충', r'\1층', addr_str)
    addr_str = re.sub(r'\d+(로트|블록|공구|지구)', '', addr_str)
    sido_corrections = {
        "서욽특별시": "서울특별시", "서욽시": "서울특별시", "서울시": "서울특별시", "서울": "서울특별시",
        "경기": "경기도",
        "강원": "강원특별자치도", "강원도": "강원특별자치도",
        "충복": "충청북도", "충남": "충청남도",
        "전복": "전북특별자치도", "전북": "전북특별자치도", "전라북도": "전북특별자치도",
        "전남": "전라남도",
        "경상복도": "경상북도", "경복": "경상북도", "경북": "경상북도",
        "경상남도": "경상남도", "경남": "경상남도",
        "제주": "제주특별자치도", "부산": "부산광역시",
        "대구": "대구광역시", "인천": "인천광역시",
        "광주": "광주광역시", "대전": "대전광역시",
        "울산": "울산광역시", "세종": "세종특별자치시", "세종시": "세종특별자치시"
    }
    for bad, good in sido_corrections.items():
        pattern = re.compile(f'^{bad}(?=\\s|$)', re.IGNORECASE)
        if pattern.match(addr_str):
            addr_str = pattern.sub(good, addr_str)
            break
    return addr_str


def remove_bad_tokens(addr_str: str) -> str:
    if not addr_str:
        return ""
    addr_str = re.sub(r'\b동(?=\d+호)', '', addr_str)
    temp = FLOOR_RE.sub(" ", addr_str)
    temp = BAD_PREFIX_DONG_RE.sub("", temp)
    temp = SEDAE_RE.sub("", temp)
    temp = re.sub(r'(^|\s)동(\s|$)', r'\1\2', temp)
    return temp.strip()


def normalize_bunji(addr: str) -> str:
    addr = re.sub(r'(\d+)번지(\d+)호?', r'\1-\2', addr)
    addr = re.sub(r'(\d+)의(\d+)', r'\1-\2', addr)
    addr = re.sub(r'(\d+)번지\b', r'\1', addr)
    return addr


def classify_and_reorder_tokens(addr_str: str) -> Tuple[str, str, bool]:
    tokens = addr_str.split()
    address_parts = []
    building_parts = []
    has_bunji = False
    for t in tokens:
        order = get_token_order(t)
        if order <= 6:
            address_parts.append((order, t))
            if order == 6:
                has_bunji = True
        else:
            building_parts.append(t)
    address_parts.sort(key=lambda x: x[0])
    sorted_addr = " ".join([t for _, t in address_parts])
    building_name = "".join(building_parts)
    return sorted_addr, building_name, has_bunji


def split_addr_and_building(text: str) -> Tuple[str, str]:
    tokens = text.split()
    addr_parts = []
    building_parts = []
    valid_suffixes = ('시', '도', '군', '구', '읍', '면', '동', '리', '가', '로', '길')
    for t in tokens:
        is_addr = False
        if re.match(r'^(산)?\d', t):
            is_addr = True
        elif t.endswith(valid_suffixes):
            is_addr = True
        if is_addr:
            addr_parts.append(t)
        else:
            building_parts.append(t)
    return " ".join(addr_parts), "".join(building_parts)


def mark_uncertain_case(result: dict) -> str:
    if '삭제' in result.get('상태', ''):
        return ""
    confidence = result.get('신뢰도', 100)
    status = result.get('상태', '')
    is_uncertain = (confidence < 70 or '재검색' in status or '문법분석' in status or '폴백' in status)
    if is_uncertain:
        if confidence < 60:
            return "높음"
        else:
            return "보통"
    return "N"


def calculate_confidence_score(original: str, api_result: dict, method: str) -> float:
    score = 0
    final_addr = api_result.get('최종주소', '')
    if not final_addr:
        return 0
    similarity = SequenceMatcher(None, original.lower(), final_addr.lower()).ratio()
    score += similarity * 40
    method_scores = {
        'VWorld': 30, 'JusoAPI': 28, 'Juso': 28,
        '건물명검색-번지없음': 25, '건물명검색': 22, '재검색': 20, '문법분석': 10
    }
    for key, value in method_scores.items():
        if key in method:
            score += value
            break
    has_sido = any(x in final_addr for x in ['특별시', '광역시', '도', '특별자치시', '특별자치도'])
    has_sigungu = bool(re.search(r'[가-힣]+[시군구]\b', final_addr))
    has_detail = api_result.get('동') or api_result.get('호')
    if has_sido: score += 7
    if has_sigungu: score += 7
    if has_detail: score += 6
    if api_result.get('동') and api_result.get('호'):
        score += 10
    elif api_result.get('호'):
        score += 5
    return min(100, score)


def validate_dong_ho_range(dong: str, ho: str) -> Tuple[bool, str]:
    if ho:
        try:
            ho_nums = re.findall(r'\d+', ho)
            if ho_nums:
                ho_num = int(ho_nums[-1])
                if ho_num > 10000:
                    return False, f"비정상호수({ho_num})"
                if ho_num < 10:
                    return False, f"비정상호수({ho_num})"
        except:
            pass
    if dong and not is_valid_apt_dong(dong):
        return False, f"비정상동명({dong})"
    return True, "정상"


def calculate_distance(lat1: float, lon1: float, lat2: float, lon2: float) -> float:
    from math import radians, sin, cos, sqrt, atan2
    lat1, lon1 = radians(lat1), radians(lon1)
    lat2, lon2 = radians(lat2), radians(lon2)
    dlat = lat2 - lat1
    dlon = lon2 - lon1
    a = sin(dlat/2)**2 + cos(lat1) * cos(lat2) * sin(dlon/2)**2
    c = 2 * atan2(sqrt(a), sqrt(1-a))
    return 6371 * c


def validate_jibun_road_conversion(jibun_addr: str, road_addr: str) -> Tuple[bool, float]:
    if not jibun_addr or not road_addr:
        return True, 0
    xy_jibun = vworld_getcoord_cached(jibun_addr, "parcel")
    xy_road = vworld_getcoord_cached(road_addr, "road")
    if xy_jibun and xy_road:
        distance_km = calculate_distance(xy_jibun[1], xy_jibun[0], xy_road[1], xy_road[0])
        distance_m = distance_km * 1000
        if distance_m < 200:
            return True, distance_m
        else:
            logger.warning(f"지번-도로명 거리 차이: {distance_m:.0f}m")
            return False, distance_m
    return True, 0


class SmartAddressCleaner:
    SIDO_MAP = {
        "서울": "서울특별시", "부산": "부산광역시", "대구": "대구광역시",
        "인천": "인천광역시", "광주": "광주광역시", "대전": "대전광역시",
        "울산": "울산광역시", "세종": "세종특별자치시",
        "경기": "경기도", "강원": "강원특별자치도", "충북": "충청북도",
        "충남": "충청남도", "전북": "전북특별자치도", "전남": "전라남도",
        "경북": "경상북도", "경남": "경상남도", "제주": "제주특별자치도"
    }
    MISSING_GU_MAP = {
        "부천시": {
            "심곡동": "원미구", "원미동": "원미구", "춘의동": "원미구", "도당동": "원미구",
            "약대동": "원미구", "중동": "원미구", "상동": "원미구", "역곡동": "원미구", "소사동": "원미구",
            "소사본동": "소사구", "심곡본동": "소사구", "범박동": "소사구", "괴안동": "소사구",
            "송내동": "소사구", "옥길동": "소사구", "계수동": "소사구",
            "오정동": "오정구", "여월동": "오정구", "작동": "오정구", "원종동": "오정구",
            "고강동": "오정구", "대장동": "오정구", "삼정동": "오정구", "내동": "오정구", "성곡동": "오정구"
        }
    }

    @classmethod
    def clean(cls, addr_str: str) -> str:
        if not addr_str:
            return ""
        temp_check = addr_str.split()
        if len(temp_check) > 4 and temp_check[0] == temp_check[len(temp_check) // 2]:
            addr_str = remove_repeated_phrase(addr_str)
        temp = fix_admin_typos(addr_str)
        temp = remove_bad_tokens(temp)
        temp = ZIPCODE_RE.sub("", temp).strip()
        temp = ET_AL_RE.sub(" ", temp)
        tokens = temp.split()
        if not tokens:
            return ""
        new_tokens = []
        is_bucheon = False
        first = tokens[0]
        matched_prov = next((v for k, v in cls.SIDO_MAP.items() if k in first or v in first), first)
        if "부천" in temp:
            is_bucheon = True
        if matched_prov == "부천시":
            matched_prov = "경기도"
            new_tokens = ["경기도", "부천시"]
            start_idx = 1
        else:
            new_tokens.append(matched_prov)
            start_idx = 1
        for i in range(start_idx, len(tokens)):
            t = tokens[i]
            if t == "동":
                continue
            if is_bucheon and t in cls.MISSING_GU_MAP["부천시"]:
                gu_name = cls.MISSING_GU_MAP["부천시"][t]
                if new_tokens[-1] != gu_name and not new_tokens[-1].endswith("구"):
                    new_tokens.append(gu_name)
            new_tokens.append(t)
        return " ".join(new_tokens)


# ==============================================================================
# [주소 정제 및 파싱 로직]
# ==============================================================================
def normalize_spaces(text: str) -> str:
    if not isinstance(text, str):
        return ""
    text = text.replace("번지", "")
    return re.sub(r'\s+', ' ', text).strip()

def normalize_dong_token(token: str | None) -> str | None:
    if not token: return None
    token = token.strip().lstrip("제")
    if token.endswith("동"): token = token[:-1]
    return token or None

def normalize_ho_token(token: str | None) -> str | None:
    if not token: return None
    token = token.strip().lstrip("제")
    if token.endswith("호"): token = token[:-1]
    return token or None

def is_valid_apt_dong(dong_str: str) -> bool:
    if not dong_str: return False
    dong_clean = dong_str.replace('동', '').strip()
    if not dong_clean: return False
    if dong_clean.isdigit(): return True
    if len(dong_clean) == 1 and dong_clean.isupper() and dong_clean.isalpha(): return True
    if len(dong_clean) == 1 and '가' <= dong_clean <= '힣': return True
    # [PATCH-7] 철자형 영문동 (에이동/비동/씨동 등) - 정규식과 일관성 유지
    if dong_clean in ('에이', '비', '씨', '디', '이', '에프', '지'): return True
    return False


def extract_dong_ho_enhanced(addr: str):
    text = remove_bad_tokens(addr)
    for pattern in COMPLEX_DONG_HO_PATTERNS:
        match = pattern.search(text)
        if match:
            dong = match.group('dong')
            ho = match.group('ho')
            start, end = match.span()
            if start > 0 and text[start-1].isdigit():
                continue
            if not is_valid_apt_dong(dong):
                continue
            clean_text = pattern.sub(' ', text)
            return normalize_spaces(clean_text), dong, ho
    dong, ho = None, None
    m0 = re.search(r'\b(?:제)?(?P<dong>\d+)-(?P<ho>\d+)호\b', text) or \
         re.search(r'\b(?:제)?(?P<dong>\d+)동\s*(?P<ho>\d+)호\b', text)
    if m0:
        dong_candidate = m0.group('dong') + "동"
        if is_valid_apt_dong(dong_candidate):
            dong = normalize_dong_token(dong_candidate)
            ho = normalize_ho_token(m0.group('ho') + "호")
            text = text.replace(m0.group(0), " ")
    if dong is None or ho is None:
        m1 = re.search(r'\b(?:제)?(?P<dong>[A-Za-z가-힣])동\s*(?P<ho>\d+)호\b', text)
        if m1:
            dong_candidate = m1.group('dong') + "동"
            if is_valid_apt_dong(dong_candidate):
                dong = normalize_dong_token(dong_candidate)
                ho = normalize_ho_token(m1.group('ho') + "호")
                text = text.replace(m1.group(0), " ")
    if dong is None:
        d_al = BUILDING_DONG_ALNUM_RE.search(text) or BUILDING_DONG_KO_RE.search(text)
        if d_al:
            dong_candidate = d_al.group('dong')
            if is_valid_apt_dong(dong_candidate):
                dong = normalize_dong_token(dong_candidate)
                text = text.replace(d_al.group(0), " ")
    if ho is None:
        h = HO_RE.search(text)
        if h:
            ho = normalize_ho_token(h.group('ho') + "호")
            text = text.replace(h.group(0), " ")
    clean_text = normalize_spaces(text)
    return clean_text, dong, ho

def extract_dong_ho_precise(addr: str):
    return extract_dong_ho_enhanced(addr)

def regex_fallback_parsing(addr_str: str):
    cleaned_base = SmartAddressCleaner.clean(addr_str)
    clean_addr_no_dh, dong, ho = extract_dong_ho_precise(cleaned_base)
    clean_addr = PAREN_CONTENT_RE.sub(" ", clean_addr_no_dh)
    clean_addr = normalize_spaces(clean_addr)
    return clean_addr.strip(), dong, ho

def extract_si_sgg_simple(addr: str):
    if not isinstance(addr, str): return None, None
    parts = addr.strip().split()
    if len(parts) >= 2: return parts[0], parts[1]
    return None, None

def is_building_name_only(addr_str: str) -> Tuple[bool, str]:
    if not addr_str: return False, ""
    temp = addr_str
    temp = re.sub(r'\d+동\s*\d+호', ' ', temp)
    temp = re.sub(r'[가-힣A-Za-z]동\s*\d+호', ' ', temp)
    temp = re.sub(r'\d+호', ' ', temp)
    temp = re.sub(r'\d+동', ' ', temp)
    temp = normalize_spaces(temp)
    has_sido = any(x in temp for x in ['특별시', '광역시', '도', '특별자치시', '특별자치도'])
    has_sigungu = bool(re.search(r'[가-힣]+[시군구]\b', temp))
    has_eupmyeondong = bool(re.search(r'[가-힣]+[읍면동리가]\b', temp))
    has_road = bool(re.search(r'[가-힣]+[로길]\b', temp))
    has_bunji = bool(re.search(r'\b(산)?\d+(-\d+)?\b', temp))
    has_address_element = has_sido or has_sigungu or has_eupmyeondong or has_road or has_bunji
    if not has_address_element and len(temp) >= 2:
        building_name = PAREN_CONTENT_RE.sub(' ', temp)
        building_name = normalize_spaces(building_name)
        if len(building_name) >= 2 and not building_name.isdigit():
            return True, building_name
    return False, ""

def extract_building_name(addr_str: str) -> str:
    if not addr_str: return ""
    building_patterns = [
        r'([가-힣A-Za-z0-9]+(?:아파트|APT|apt))', r'([가-힣A-Za-z0-9]+(?:빌라|Villa|villa))',
        r'([가-힣A-Za-z0-9]+(?:오피스텔|officetel))', r'([가-힣A-Za-z0-9]+(?:타워|Tower|tower|TOWER))',
        r'([가-힣A-Za-z0-9]+(?:빌딩|Building|building|Bldg))', r'([가-힣A-Za-z0-9]+(?:센터|Center|center))',
        r'([가-힣A-Za-z0-9]+(?:플라자|Plaza|plaza))', r'([가-힣A-Za-z0-9]+(?:파크|Park|park))',
        r'([가-힣A-Za-z0-9]+(?:힐스|Hills|hills))', r'([가-힣A-Za-z0-9]+(?:캐슬|Castle|castle))',
        r'([가-힣A-Za-z0-9]+(?:팰리스|Palace|palace))', r'([가-힣A-Za-z0-9]+(?:스카이|Sky|sky))',
        r'([가-힣A-Za-z0-9]+(?:푸르지오|자이|래미안|e편한세상|롯데캐슬|힐스테이트|더샵))',
    ]
    for pattern in building_patterns:
        match = re.search(pattern, addr_str, re.IGNORECASE)
        if match: return match.group(1)
    return ""

def extract_building_name_from_paren(addr_str: str) -> str:
    if not addr_str: return ""
    paren_match = re.search(r'\(([^)]+)\)', addr_str)
    if not paren_match: return ""
    content = paren_match.group(1)
    # [PATCH-3] 괄호 안 토큰은 통째로 건물명으로 사용 (아이파크분당→아이파크 절단 방지)
    if ',' in content:
        parts = content.split(',')
        for part in reversed(parts):
            part = part.strip()
            if not part:
                continue
            # 순수 행정동/리/가 토큰은 건물명 아님 → 건너뜀
            if part.endswith(('동', '리', '가')) and ' ' not in part:
                continue
            return part  # 절단 없이 전체 반환
        return parts[-1].strip()
    content = content.strip()
    if content.endswith(('동', '리', '가')) and ' ' not in content:
        return ""
    return content

def check_has_bunji(addr_str: str, dong_str: str = None, ho_str: str = None) -> bool:
    if not addr_str: return False
    temp = addr_str
    if dong_str: temp = temp.replace(dong_str, ' ')
    if ho_str: temp = temp.replace(ho_str, ' ')
    temp = re.sub(r'\d+동', '', temp)
    temp = re.sub(r'\d+호', '', temp)
    temp = re.sub(r'[가-힣A-Z]동', '', temp)
    return bool(re.search(r'\b(산)?\d+(-\d+)?\b', temp))

def standardize_grammar_result(addr_str: str, dong: str, ho: str) -> Tuple[str, bool, str]:
    if not addr_str: return "", False, "주소없음"
    addr = normalize_bunji(addr_str)
    sorted_addr, building_name, has_bunji = classify_and_reorder_tokens(addr)
    has_road_number = bool(re.search(r'(로|길)\s+\d+', addr))
    if not has_bunji and not has_road_number:
        return "", False, "번지/건물번호없음"
    if not ho: return "", False, "호정보없음"
    if ho:
        nums = re.findall(r'\d+', ho)
        if nums and int(nums[-1]) < 100:
            return "", False, f"100호미만({ho})"
    result_parts = [sorted_addr]
    if building_name: result_parts.append(building_name)
    if dong: result_parts.append(f"{dong}동")
    if ho: result_parts.append(f"{ho}호")
    return " ".join(result_parts), True, "통과"

def validate_address_quality(result: dict) -> Tuple[bool, str]:
    addr = result.get('최종주소', '')
    if not addr: return False, "주소없음"
    required_elements = {
        '시도': any(x in addr for x in ['특별시', '광역시', '도', '특별자치시', '특별자치도']),
        '시군구': any(x in addr for x in ['시', '군', '구']),
        '번지': bool(re.search(r'\d+(-\d+)?', addr)),
    }
    missing = [k for k, v in required_elements.items() if not v]
    if missing: return False, f"필수요소부족({','.join(missing)})"
    if len(addr) < 10: return False, "주소길이부족"
    tokens = addr.split()
    unique_tokens = set(tokens)
    if len(tokens) - len(unique_tokens) > 1: return False, "중복토큰과다"
    return True, "정상"


# ★★★ [변경2] check_basement 함수 교체 - B동 오분류 방지 ★★★
def check_basement(addr_str: str) -> bool:
    """
    지하/반지하 체크 (개선)
    - B동(건물동)을 지하로 오분류하지 않음
    - B01, B02 등 100 미만은 지하로 판단, B동은 제외
    """
    if not addr_str:
        return False

    # 1. 명시적 지하 키워드
    if re.search(r'지하\s*\d*층?', addr_str):
        return True
    if re.search(r'지층', addr_str):
        return True
    if re.search(r'반지하', addr_str):
        return True

    # 2. B호수 판별 (B동 제외)
    temp = re.sub(r'[Bb]동', '', addr_str)
    temp = re.sub(r'비동', '', temp)

    # B + 숫자(100 미만) = 지하 (B01, B02 등)
    # B + 숫자(100 이상) = 건물동 가능성 (B702호 등)
    b_match = re.search(r'\bB(\d{1,3})호?\b', temp, re.IGNORECASE)
    if b_match:
        b_num = int(b_match.group(1))
        if b_num < 100:
            return True

    # "비01호", "비02호" 패턴
    if re.search(r'비0?\d{1,2}호', temp):
        b_match2 = re.search(r'비0?(\d{1,2})호', temp)
        if b_match2 and int(b_match2.group(1)) < 100:
            return True

    return False


def check_filters_final(addr_str: str, ho_str: str) -> Tuple[bool, str]:
    if not addr_str: return False, "주소없음"
    DELETE_KEYWORDS = ["차량", "자동차", "중기", "지분", "대지", "임야", "목록", "입찰외"]
    for kw in DELETE_KEYWORDS:
        if kw in addr_str: return False, f"삭제({kw})"
    if FILTER_OPTS["SHARE"] and ("지분" in addr_str or "1/" in addr_str):
        return False, "지분매각"
    if FILTER_OPTS["MOUNTAIN"] and re.search(r'\s산\d+', addr_str):
        return False, "임야(산)"
    if FILTER_OPTS["BASEMENT"] and check_basement(addr_str):
        return False, "지하/반지하"
    if FILTER_OPTS["NO_BUNJI"]:
        temp_check = addr_str
        temp_check = HO_RE.sub("", temp_check)
        temp_check = BUILDING_DONG_ALNUM_RE.sub("", temp_check)
        temp_check = BUILDING_DONG_KO_RE.sub("", temp_check)
        temp_check = FLOOR_RE.sub("", temp_check)
        has_real_bunji = re.search(r'\d', temp_check)
        if not has_real_bunji: return False, "번지/건물번호 없음"
    if FILTER_OPTS["NO_HINT"]:
        text_removed = re.sub(r'\d', '', addr_str)
        if not MIN_LOCATION_TOKEN_RE.search(text_removed): return False, "위치단서부족"
    if FILTER_OPTS["NO_HO"] and not ho_str: return False, "호 정보 없음"
    if FILTER_OPTS["UNDER_100"] and ho_str:
        try:
            nums = re.findall(r'\d+', ho_str)
            if nums and int(nums[-1]) < 100: return False, f"100호 미만({ho_str})"
        except: pass
    return True, "통과"


# ==============================================================================
# [API 함수]
# ==============================================================================
@lru_cache(maxsize=5000)
def vworld_getcoord_cached(addr: str, addr_type: str = "road"):
    return vworld_getcoord(addr, addr_type)

def vworld_getcoord(addr: str, addr_type: str = "road"):
    url = "https://api.vworld.kr/req/address"
    params = {"service": "address", "request": "getcoord", "version": "2.0",
              "crs": "epsg:4326", "address": addr, "format": "json", "type": addr_type, "key": VWORLD_KEY}
    try:
        r = session.get(url, params=params, timeout=TIMEOUT_SEC, verify=False)
        data = r.json()
        if data["response"]["status"] == "OK":
            p = data["response"]["result"]["point"]
            return float(p["x"]), float(p["y"])
    except Exception as e:
        logger.debug(f"VWorld 좌표 조회 실패 ({addr_type}): {addr[:30]}... - {e}")
    return None

def vworld_getcoord_with_fallback(addr: str):
    result = vworld_getcoord_cached(addr, "road")
    if result: return result
    return vworld_getcoord_cached(addr, "parcel")

@lru_cache(maxsize=5000)
def vworld_reverse_geocode_cached(x: float, y: float):
    return vworld_reverse_geocode(x, y)

def vworld_reverse_geocode(x, y):
    url = "https://api.vworld.kr/req/address"
    params = {"service": "address", "request": "getaddress", "version": "2.0",
              "crs": "epsg:4326", "point": f"{x},{y}", "format": "json", "type": "BOTH", "key": VWORLD_KEY}
    try:
        r = session.get(url, params=params, timeout=TIMEOUT_SEC, verify=False)
        data = r.json()
        if data["response"]["status"] == "OK":
            res = data["response"]["result"]
            road = next((i["text"] for i in res if i["type"] == "road"), None)
            parcel = next((i["text"] for i in res if i["type"] == "parcel"), None)
            si, sgg = None, None
            target = road if road else parcel
            if target:
                parts = target.split()
                if len(parts) >= 2: si, sgg = parts[0], parts[1]
            return road, parcel, si, sgg
    except Exception as e:
        logger.debug(f"VWorld 역지오코딩 실패: ({x}, {y}) - {e}")
    return None, None, None, None

@lru_cache(maxsize=5000)
def search_juso_gov_cached(keyword: str):
    return search_juso_gov(keyword)

def search_juso_gov(keyword: str):
    url = "https://business.juso.go.kr/addrlink/addrLinkApi.do"
    params = {"confmKey": JUSO_API_KEY, "currentPage": 1, "countPerPage": 1, "keyword": keyword, "resultType": "json"}
    try:
        r = session.get(url, params=params, timeout=TIMEOUT_SEC, verify=False)
        data = r.json()
        if data['results']['common']['errorCode'] == "0" and int(data['results']['common']['totalCount']) > 0:
            return data['results']['juso'][0]
    except Exception as e:
        logger.debug(f"Juso API 실패: {keyword[:30]}... - {e}")
    return None

def search_building_name(building_name: str, dong: str = None, ho: str = None, region_hint: str = None):
    if not building_name: return None
    clean_name = normalize_spaces(building_name)
    if len(clean_name) < 3: return None
    logger.info(f"건물명 검색 시도: {clean_name}" + (f" (지역: {region_hint})" if region_hint else ""))
    def search_building_multiple(keyword, count=10):
        url = "https://business.juso.go.kr/addrlink/addrLinkApi.do"
        params = {"confmKey": JUSO_API_KEY, "currentPage": 1, "countPerPage": count, "keyword": keyword, "resultType": "json"}
        try:
            r = session.get(url, params=params, timeout=TIMEOUT_SEC, verify=False)
            data = r.json()
            if data['results']['common']['errorCode'] == "0" and int(data['results']['common']['totalCount']) > 0:
                return data['results']['juso']
        except: pass
        return None
    def find_exact_match(results, target_name, region_filter=None):
        if not results: return None
        def normalize_for_compare(text):
            if not text: return ""
            text = re.sub(r'[\s\-_]', '', text.lower())
            for suffix in ['아파트', 'apt', '빌라', 'villa', '오피스텔', 'officetel']:
                text = text.replace(suffix, '')
            return text
        target_normalized = normalize_for_compare(target_name)
        exact_matches, partial_matches = [], []
        for result in results:
            bdNm = result.get('bdNm', '')
            if not bdNm: continue
            if region_filter:
                full_region = f"{result.get('siNm', '')} {result.get('sggNm', '')}"
                if region_filter not in full_region: continue
            bdNm_normalized = normalize_for_compare(bdNm)
            if bdNm_normalized == target_normalized:
                exact_matches.append(result)
            elif target_normalized in bdNm_normalized:
                partial_matches.append(result)
        if exact_matches: return exact_matches[0]
        if partial_matches:
            partial_matches.sort(key=lambda x: len(x.get('bdNm', '')))
            return partial_matches[0]
        return None
    search_queries = []
    if region_hint: search_queries.append((f"{region_hint} {clean_name}", region_hint))
    search_queries.append((clean_name, region_hint))
    for search_keyword, region_filter in search_queries:
        results = search_building_multiple(search_keyword, count=20)
        if results:
            exact_result = find_exact_match(results, clean_name, region_filter)
            if exact_result: return exact_result
    variations = []
    if not any(x in clean_name.lower() for x in ['아파트', 'apt']): variations.append(f"{clean_name}아파트")
    if ' ' in clean_name: variations.append(clean_name.replace(' ', ''))
    for var in variations:
        search_keyword = f"{region_hint} {var}" if region_hint else var
        results = search_building_multiple(search_keyword, count=20)
        if results:
            exact_result = find_exact_match(results, clean_name, region_hint)
            if exact_result: return exact_result
    return None

def cross_validate_address(addr_str: str, building_name: str = None, region_hint: str = None):
    results = []
    xy = vworld_getcoord_with_fallback(addr_str)
    if xy:
        road, parcel, si, sgg = vworld_reverse_geocode_cached(*xy)
        if road or parcel:
            results.append({'method': 'VWorld', 'road': road, 'jibun': parcel, 'si': si, 'sgg': sgg, 'xy': xy})
    juso_res = search_juso_gov_cached(addr_str)
    if juso_res:
        results.append({'method': 'Juso', 'road': juso_res['roadAddrPart1'], 'jibun': juso_res['jibunAddr'],
                        'si': juso_res['siNm'], 'sgg': juso_res['sggNm'], 'bdNm': juso_res.get('bdNm', ''), 'juso_obj': juso_res})
    if building_name:
        bldg_res = search_building_name(building_name, region_hint=region_hint)
        if bldg_res:
            results.append({'method': '건물명', 'road': bldg_res['roadAddrPart1'], 'jibun': bldg_res['jibunAddr'],
                            'si': bldg_res['siNm'], 'sgg': bldg_res['sggNm'], 'bdNm': bldg_res.get('bdNm', ''), 'juso_obj': bldg_res})
    if not results: return None, None, False
    if len(results) == 1: return results[0], results[0]['method'], False
    from collections import Counter
    si_sgg_list = [(r['si'], r['sgg']) for r in results]
    si_sgg_counter = Counter(si_sgg_list)
    if si_sgg_counter:
        most_common_region, count = si_sgg_counter.most_common(1)[0]
        if count >= len(results) / 2:
            matching_results = [r for r in results if (r['si'], r['sgg']) == most_common_region]
            priority = {'건물명': 3, 'Juso': 2, 'VWorld': 1}
            matching_results.sort(key=lambda x: priority.get(x['method'], 0), reverse=True)
            selected = matching_results[0]
            return selected, f"{selected['method']}(교차검증)", True
    return results[0], results[0]['method'], False


# ==============================================================================
# [통계 로딩]
# ==============================================================================
def normalize_sido(name: str):
    if not name: return ""
    t = re.sub(r'\s+', '', str(name))
    for s in ["특별자치도", "특별자치시", "특별시", "광역시", "자치도", "자치시", "도"]:
        if t.endswith(s): return t[:-len(s)]
    return t

def normalize_sigungu(name: str):
    return re.sub(r'\s+', '', str(name)) if name else ""

def load_stats():
    if not os.path.isfile(STATS_FILE_PATH):
        logger.warning(f"통계 파일 없음: {STATS_FILE_PATH}"); return
    try:
        df = pd.read_excel(STATS_FILE_PATH).dropna(subset=['낙찰가율'])
        df['낙찰가율_v'] = pd.to_numeric(df['낙찰가율'].astype(str).str.replace(r'[%|,]', '', regex=True), errors='coerce')
        for _, row in df.iterrows():
            k = (normalize_sido(row['시도명']), normalize_sigungu(row['시군구명']))
            if row['기간구분'] == '3개월': STATS_3M[k] = row['낙찰가율_v']
            elif row['기간구분'] == '1년': STATS_1Y[k] = row['낙찰가율_v']
        logger.info(f"통계 로딩 완료: 3개월 {len(STATS_3M)}건, 1년 {len(STATS_1Y)}건")
    except Exception as e:
        logger.error(f"통계 로딩 실패: {e}")

def get_rates(si, sgg):
    k = (normalize_sido(si), normalize_sigungu(sgg))
    return STATS_3M.get(k), STATS_1Y.get(k)

def print_statistics(df_result):
    if '상태' not in df_result.columns: return
    counts = df_result['상태'].value_counts()
    total = len(df_result)
    print("\n" + "=" * 30)
    print(f" [처리 결과 통계] 총 {total}건")
    print("=" * 30)
    stats_summary = {"VWorld": 0, "Juso": 0, "건물명검색": 0, "문법분석": 0, "캐시": 0}
    for key, val in counts.items():
        if "캐시" in key: stats_summary["캐시"] += val
        elif "VWorld" in key: stats_summary["VWorld"] += val
        elif "Juso" in key: stats_summary["Juso"] += val
        elif "건물명" in key: stats_summary["건물명검색"] += val
        elif "문법" in key: stats_summary["문법분석"] += val
    for k, v in stats_summary.items():
        ratio = (v / total * 100) if total > 0 else 0
        print(f" - {k}: {v}건 ({ratio:.1f}%)")
    print("=" * 30 + "\n")


# ==============================================================================
# [단일 주소 처리]
# ==============================================================================
def process_single_address(raw_addr, idx: int) -> dict:
    # NaN 안전 처리
    addr_str = to_addr_str(raw_addr)
    if not addr_str:
        return {"상태": "삭제(주소없음)", "신뢰도": 0}

    # ★★★ [변경3] 전처리 호출 추가 ★★★
    addr_str = preprocess_raw_address(addr_str)
    if not addr_str:
        return {"상태": "삭제(전처리후공백)", "신뢰도": 0}

    # [Step 0] 1차 필터링
    DELETE_KEYWORDS = ["차량", "자동차", "중기", "지분", "대지", "임야", "목록", "입찰외"]
    for kw in DELETE_KEYWORDS:
        if kw in addr_str:
            return {"상태": f"삭제({kw})", "신뢰도": 0}

    building_in_paren = extract_building_name_from_paren(addr_str)
    
    # [Step 1] 스마트 정제
    smart_cleaned = SmartAddressCleaner.clean(addr_str)
    clean_addr, dong_temp, ho_temp = extract_dong_ho_precise(normalize_spaces(smart_cleaned))
    
    region_hint = None
    parts = smart_cleaned.split()
    if len(parts) >= 2:
        if any(parts[1].endswith(x) for x in ['시', '군', '구']):
            region_hint = parts[1]
    
    has_bunji = check_has_bunji(clean_addr, dong_temp, ho_temp)
    query_addr = PAREN_CONTENT_RE.sub(" ", clean_addr).strip()

    final_road, final_jibun, final_si, final_sgg, method = None, None, None, None, "실패"
    final_dong, final_ho = dong_temp, ho_temp
    cross_validated = False

    # [Step 1.5] 번지 없으면 건물명 검색 우선
    if not has_bunji and building_in_paren:
        bldg_result = search_building_name(building_in_paren, final_dong, final_ho, region_hint)
        if bldg_result:
            final_road = bldg_result['roadAddrPart1']
            final_jibun = bldg_result['jibunAddr']
            final_si = bldg_result['siNm']
            final_sgg = bldg_result['sggNm']
            method = "성공(건물명검색-번지없음)"
            if bldg_result.get('bdNm') and bldg_result['bdNm'] not in final_road:
                final_road += f" ({bldg_result['bdNm']})"

    # [Step 2] 교차 검증 시도
    if not final_road and has_bunji:
        validated_result, validated_method, is_cross_validated = cross_validate_address(query_addr, building_in_paren, region_hint)
        if validated_result:
            final_road = validated_result['road']
            final_jibun = validated_result['jibun']
            final_si = validated_result['si']
            final_sgg = validated_result['sgg']
            method = validated_method
            cross_validated = is_cross_validated
            if validated_result.get('bdNm') and validated_result['bdNm'] not in final_road:
                final_road += f" ({validated_result['bdNm']})"

    # [Step 3] VWorld
    if not final_road:
        xy = vworld_getcoord_with_fallback(query_addr)
        if xy:
            road, parcel, si, sgg = vworld_reverse_geocode_cached(*xy)
            if road or parcel:
                final_road, final_jibun, final_si, final_sgg = road, parcel, si, sgg
                method = "성공(VWorld)"

    # [Step 4] Juso API
    if not final_road:
        juso_res = search_juso_gov_cached(query_addr)
        if juso_res:
            final_road, final_jibun = juso_res['roadAddrPart1'], juso_res['jibunAddr']
            final_si, final_sgg = juso_res['siNm'], juso_res['sggNm']
            method = "성공(JusoAPI)"
            if juso_res.get('bdNm') and not final_dong:
                final_road += f" ({juso_res['bdNm']})"

    # [Step 5] 재시도
    if not final_road:
        local_addr, local_dong, local_ho = regex_fallback_parsing(addr_str)
        retry_query = local_addr
        final_dong = local_dong or final_dong
        final_ho = local_ho or final_ho
        validated_result, validated_method, is_cross_validated = cross_validate_address(retry_query, building_in_paren, region_hint)
        if validated_result:
            final_road = validated_result['road']
            final_jibun = validated_result['jibun']
            final_si = validated_result['si']
            final_sgg = validated_result['sgg']
            method = f"성공(재검색-{validated_method})"
            cross_validated = is_cross_validated
            if validated_result.get('bdNm') and validated_result['bdNm'] not in final_road:
                final_road += f" ({validated_result['bdNm']})"

    # [Step 6] 건물명 재검색
    if not final_road and building_in_paren:
        bldg_result = search_building_name(building_in_paren, final_dong, final_ho, region_hint)
        if bldg_result:
            final_road = bldg_result['roadAddrPart1']
            final_jibun = bldg_result['jibunAddr']
            final_si = bldg_result['siNm']
            final_sgg = bldg_result['sggNm']
            method = "성공(건물명검색-폴백)"
            if bldg_result.get('bdNm') and bldg_result['bdNm'] not in final_road:
                final_road += f" ({bldg_result['bdNm']})"

    # [Step 7] 주소에서 건물명 추출
    if not final_road and not building_in_paren:
        extracted_bldg = extract_building_name(addr_str)
        if extracted_bldg:
            bldg_result = search_building_name(extracted_bldg, final_dong, final_ho, region_hint)
            if bldg_result:
                final_road = bldg_result['roadAddrPart1']
                final_jibun = bldg_result['jibunAddr']
                final_si = bldg_result['siNm']
                final_sgg = bldg_result['sggNm']
                method = "성공(건물명검색-추출)"
                if bldg_result.get('bdNm') and bldg_result['bdNm'] not in final_road:
                    final_road += f" ({bldg_result['bdNm']})"

    # [Step 8] 유사도 매칭
    if not final_road:
        similar_result, score = similarity_matcher.find_similar(addr_str)
        if similar_result:
            similar_copy = similar_result.copy()
            similar_copy['상태'] = f"{similar_copy['상태']}(유사매칭)"
            return similar_copy

    # [Step 9] 문법분석 폴백
    if not final_road:
        standardized, is_valid, reason = standardize_grammar_result(local_addr, final_dong, final_ho)
        if not is_valid:
            return {"상태": f"삭제({reason})", "신뢰도": 0}
        final_road = standardized
        final_jibun = standardized
        method = "성공(문법분석)"
        final_si, final_sgg = extract_si_sgg_simple(local_addr)

    # [Step 10] API 성공 케이스
    if method != "성공(문법분석)":
        detail = ""
        if final_dong: detail += f" {final_dong}동"
        if final_ho: detail += f" {final_ho}호"
        full_addr = f"{final_road}{detail}".strip()
        dong_ho_valid, dong_ho_reason = validate_dong_ho_range(final_dong, final_ho)
        if not dong_ho_valid:
            return {"상태": f"삭제({dong_ho_reason})", "신뢰도": 0}
        passed, reason = check_filters_final(full_addr, final_ho)
        if not passed:
            return {"상태": f"삭제({reason})", "신뢰도": 0}
        final_road = full_addr
        if final_jibun:
            final_jibun = f"{final_jibun}{detail}".strip()

    # [Step 11] 지번-도로명 좌표 검증
    coord_valid = True
    coord_distance = 0
    if final_road and final_jibun and final_road != final_jibun:
        coord_valid, coord_distance = validate_jibun_road_conversion(final_jibun, final_road)

    # [Step 12] 낙찰가율 조회
    r3, r1 = get_rates(final_si, final_sgg)

    result = {
        "최종주소": final_road,
        "도로명_완성": final_road,
        "지번_완성": final_jibun or final_road,
        "동": final_dong,
        "호": final_ho,
        "상태": method,
        "낙찰가율_3M": f"{r3}%" if r3 else "",
        "낙찰가율_1Y": f"{r1}%" if r1 else "",
        "교차검증": "Y" if cross_validated else "N",
        "좌표거리_m": f"{coord_distance:.0f}" if coord_distance > 0 else ""
    }

    # [Step 13] 신뢰도
    confidence = calculate_confidence_score(addr_str, result, method)
    result['신뢰도'] = confidence

    # [Step 14] 품질 검증
    is_valid, reason = validate_address_quality(result)
    if not is_valid:
        result['상태'] = f"삭제({reason})"
        result['신뢰도'] = 0
        result['검토필요'] = ""
    else:
        if confidence < 60:
            result['상태'] = f"{result['상태']}(낮은신뢰도)"
        result['검토필요'] = mark_uncertain_case(result)
        similarity_matcher.add_success(addr_str, result)

    return result


# ==============================================================================
# [배치 처리]
# ==============================================================================
def process_addresses_batch(addresses: List[Tuple[int, str]], monitor: ProcessingMonitor) -> List[dict]:
    results = []
    for idx, raw_addr in addresses:
        time.sleep(API_DELAY)
        result = process_single_address(raw_addr, idx)
        results.append(result)
        monitor.update(result)
        if "삭제" in result.get("상태", ""):
            logger.debug(f"[{idx + 1}] {result['상태']}")
        else:
            logger.info(f"[{idx + 1}] {result['상태']}: {result.get('최종주소', '')[:50]}...")
        if (idx + 1) % 10 == 0:
            monitor.print_progress()
    return results


# ==============================================================================
# [메인 처리]
# ==============================================================================
def process_folder(folder):
    logger.info("=" * 70)
    logger.info("주소 정제 및 표준화 프로그램 시작 (v2 - 전처리 개선)")
    logger.info("=" * 70)
    load_stats()

    for fname in os.listdir(folder):
        if not fname.lower().endswith(".xlsx") or "_완료" in fname or "_최종완성본" in fname:
            continue
        full_path = os.path.join(folder, fname)
        logger.info("=" * 70)
        logger.info(f"[처리] {fname}")
        logger.info("=" * 70)
        try:
            df = pd.read_excel(full_path)
        except Exception as e:
            logger.error(f"파일 읽기 실패: {e}"); continue
        if df.shape[1] <= TARGET_COL_INDEX:
            logger.warning(f"컬럼 수 부족"); continue

        monitor = ProcessingMonitor()
        addresses = [(i, raw_addr) for i, raw_addr in df.iloc[:, TARGET_COL_INDEX].items()]
        logger.info(f"총 {len(addresses)}건 처리 시작...")

        results = process_addresses_batch(addresses, monitor)
        monitor.print_progress(clear=False)
        print()

        res_df = pd.DataFrame(results)
        cols_to_drop = [col for col in res_df.columns if col in df.columns]
        df_clean = df.drop(columns=cols_to_drop, errors='ignore')
        final = pd.concat([df_clean.reset_index(drop=True), res_df], axis=1)

        before_filter = len(final)
        final = final[~final['상태'].astype(str).str.contains("삭제")]
        after_filter = len(final)

        monitor.print_summary()
        logger.info(f"필터링: {before_filter}건 → {after_filter}건 (제외: {before_filter - after_filter}건)")

        if '검토필요' in final.columns:
            review_high = len(final[final['검토필요'] == '높음'])
            review_medium = len(final[final['검토필요'] == '보통'])
            total_review = review_high + review_medium
            if total_review > 0:
                logger.info(f"\n⚠️  검토 권장 케이스: {total_review}건")
                logger.info(f"   - 우선검토(신뢰도<60): {review_high}건")
                logger.info(f"   - 일반검토(신뢰도<70): {review_medium}건")

        outname = fname.replace(".xlsx", "_최종완성본.xlsx")
        out_path = os.path.join(folder, outname)
        final.to_excel(out_path, index=False)
        logger.info(f"저장 완료: {out_path}")
        logger.info(f"최종 결과: {len(final)}건")

    logger.info("=" * 70)
    logger.info("모든 파일 처리 완료")
    logger.info("=" * 70)


# ==============================================================================
# [실행]
# ==============================================================================
if __name__ == "__main__":
    try:
        process_folder(FOLDER)
    except KeyboardInterrupt:
        logger.warning("\n사용자에 의해 중단됨")
    except Exception as e:
        logger.error(f"예상치 못한 오류: {e}", exc_info=True)