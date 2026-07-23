# Changelog

이 프로젝트의 주요 변경 사항을 기록합니다. 형식은 [Keep a Changelog](https://keepachangelog.com/ko/1.1.0/)를,
버전은 [Semantic Versioning](https://semver.org/lang/ko/)을 따릅니다.

태그 규칙: 릴리스는 `v<버전>` 형태의 git 태그로 표시합니다 (예: `v0.3.0`).
태그 생성과 패키지 배포는 CONTRIBUTING의 승인 게이트에 따라 별도 승인 사항입니다.

## [Unreleased]

### Added
- 건물명 검색어(kind `building`): 지번/도로명 골격이 없어도 법정동+건물명이
  있으면(`수유동 한양아이빌 402호`, 괄호 주석 `(화곡동, 타운캐슬)`의 건물명 승계
  포함) 검색어를 포기하지 않음 — 실패 환류 엑셀의 `검색주소없음` 58건 중 41건 구제
- 검증 확정 골격으로 검색어 교체: 상세 포함 0건 → 골격 1건으로 확정된 행은
  후단 자동화(주소찾기 팝업)가 실패할 상세 검색어 대신 통한 골격을 검색어 열에
  기록 (`--keep-detail-query`로 비활성화, stats `query_rewritten`)
- 실패(우편번호미확정) 환류 정제 규칙: `외 N세대` 제거, `도면표시/㎡` 면적 기술
  절단(호수는 보존), `(법정동[, 건물명])` 괄호 주석 정리(건물명만 평문 승계),
  `(도로명 : …)` 주석 제거, `N블럭/N로트/N지구` 개발지구 표기 제거,
  `제N등→제N동`·`비제2층→비동`·`제지하층`·`오피스텔410호`·`제비-508호`·
  `--408`·`-동`·`4층동`·모음 자모 잔재 정리, 중복 동·호 구절 축약,
  `인천광여깃` 오타 규칙
- 라이브러리 공개 API: `verify_address()`/`VerifyResult`로 엑셀 없이 주소 1건 검증,
  `address_cleaner` 최상위에서 `normalize_for_search`, `JusoClient`,
  `KoreaPostRoadNameClient`, `SearchResult`, `process_workbook` export,
  `__version__`(pyproject 단일 출처)
- `py.typed` 마커 — 외부 사용자의 타입체커가 패키지 타입 힌트를 인식
- 일반 `excel` 모드 `--typo-rules` 지원: 교정 후보 수확 → 검토 → 승격 → 적용의
  학습 루프가 일반 모드 안에서 완결 (`address_cleaner/typo.py` 공용 모듈)
- 일반 `excel` 모드 `--mark-missing` 검증 병렬화 (`--workers`, 기본 8) —
  공용 검색 인프라(`juso_search.py`: RateLimiter/캐시/juso_query) 추출
- registry JSON 캐시 만료 (`--cache-max-age-days`, 기본 14일, 0이면 만료 없음) —
  캐시 엔트리에 `cached_at` 기록, 저장 시 만료 엔트리 청소
- ePost 엔드포인트 `EPOST_ENDPOINT` 환경변수 override
- CI 보강: mypy 타입체크, coverage 임계값, Windows(3.10/3.13) 러너, ruff format 체크
- clients/excel/history/CLI 목킹 테스트 및 직렬 vs 병렬 동등성 테스트
- 릴리스 빌드 워크플로 (`workflow_dispatch` 수동 전용, sdist/wheel 아티팩트까지만)
- `CHANGELOG.md` 신설 (0.1.0~0.3.0 git 이력 기반 소급)

### Changed
- README를 범용 소개(영문 요약 포함)로 재구성하고, 사내 HUG 연동 규약은
  `docs/hug-integration.md`로 분리

### Fixed
- `권광로175번길91`이 `권광로175 번길 91`로 갈라져 도로명이 깨지던 문제
- `양평동4가`·`문래동5가` 같은 서수 법정동이 `양평동 4가`로 갈라지던 문제
- `하이파크시티일산파밀리에…` 같은 건물명이 시/리 붙임 복원 규칙에 의해
  내부에서 갈라지던 문제
- `동탄2신도시…`, `씨10블럭…` 복합 단지명이 숫자-한글 분리 규칙에 갈라지던 문제
- `col_to_index("")`가 0(존재하지 않는 열)으로 조용히 통과하던 문제 → ValueError

## [0.3.0] - 2026-06-12

### Added
- 학습형 운영 도구화: 교정 후보 수확(`--corrections-out`), 검증 이력 SQLite DB
  (`--history`, 최근 결과 재사용·판정 변경 감지), 파워쉘 실패 환류 리포트(`feedback` 명령)
- registry 모드 Juso 검색 병렬 처리 (`--workers`, 전역 레이트리미터)
- 동/호 범위 가드(`unit_out_of_range`) 및 정제 견고성 보강

### Fixed
- 호수 앞 외톨이 '동' 정리 및 건물명 누수 차단

## [0.2.0] - 2026-06-12

### Added
- 정제 파이프라인 구조 개선: 시/도 상수 단일화(`regions.py`), 법정동 사전 오프라인
  검증(`--admin-dict`), 지번 변형 제안(`야당동 5717` → `57-17`)
- Juso 검증 2단계 검색(상세포함 → 골격) 및 검증상세 컬럼(`--detail-col`)

### Fixed
- `제1(상층하층)층` 복층 표기를 층 정보로 인식해 제거
- 정제-전산 연동에서 드러난 정제기 보완

## [0.1.0] - 2026-06-02

### Added
- 최초 릴리스: HUG 엑셀 주소 정제(`excel` 명령), juso.go.kr/우정사업본부 API 검증
  (`--mark-missing`), 단일 주소 정제(`normalize`), API 키 점검(`probe`)
- 법원 등기부등본 열람페이지 전체검색용 주소 생성(registry 모드) 통합
