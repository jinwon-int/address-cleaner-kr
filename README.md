# address-cleaner-kr (주소정제)

> **English** — A Python library and CLI that normalizes messy Korean addresses
> (from Excel sheets or free text) into API-searchable queries, then verifies
> them against the official `juso.go.kr` road-name address API and the Korea
> Post address service. It preserves building/unit detail, flags addresses that
> resolve to zero or multiple matches, and harvests correction candidates so the
> typo dictionary grows from real data. Documentation is in Korean.

지저분한 한국 주소(엑셀 원주소, 자유 서식 텍스트)를 `juso.go.kr`·우정사업본부
도로명주소조회서비스에서 검색 가능한 형태로 정제하고, 실제 API 검색으로
검증하는 Python 라이브러리 겸 CLI입니다.

## 하는 일

- 원주소에서 우편번호, `외 N필지`, 반복 주소, 깨진 공백 등 API 검색을 방해하는 잡음을 정리합니다.
- 건물명·동/호 같은 상세부는 보존하고 `제1층` 같은 층 정보만 제거합니다.
- 원주소가 잘못됐거나 주소 골격이 불명확하면 과감하게 검색어를 비우고 `검색주소없음`으로 표시합니다.
- API 검증 결과가 2건 이상이면 `2건이상검색`으로 표시해 사람이 확정하게 합니다.
- 검증 이력(SQLite)·교정 후보 수확·오타 규칙 적용으로 정제 규칙이 데이터에서 스스로 자랍니다.
- API 키는 코드에 저장하지 않고 환경변수에서 읽습니다.

## 설치

```bash
python -m venv .venv
. .venv/bin/activate
pip install -e ".[dev]"
```

패키지는 `py.typed`를 포함하므로 타입체커(mypy 등)가 타입 힌트를 인식합니다.

## 라이브러리로 쓰기

```python
from address_cleaner import normalize_for_search, verify_address

normalized = normalize_for_search("경기도 파주시 야당동 57-17 정우펠리스 제303동 제1층 제101호")
normalized.query   # '경기도 파주시 야당동 57-17 정우펠리스 제303동 제101호'
normalized.kind    # 'lot' (지번) | 'road' (도로명) | 'invalid' | 'empty'

# API 키(JUSO_CONFIRM_KEY / EPOST_SERVICE_KEY)가 환경변수에 있으면:
result = verify_address(normalized.query, normalized.kind)
result.verdict     # 'verified' | 'ambiguous' | 'missing'
result.detail      # 검색 경로/건수/표준주소/우편번호
result.correction  # 골격/지번 변형으로만 통한 경우의 교정 후보 (없으면 None)
```

엑셀 일괄 처리는 `address_cleaner.process_workbook`, 저수준 API 클라이언트는
`JusoClient`/`KoreaPostRoadNameClient`를 직접 쓸 수 있습니다.

## CLI 사용 예시

### 1) 주소 검색어 정제

단일 주소 정제:

```bash
address-cleaner normalize "경기도 파주시 야당동 57-17 정우펠리스 제303동 제1층 제101호"
```

위 예시는 `경기도 파주시 야당동 57-17 정우펠리스 제303동 제101호`처럼 층 정보만 제거하고 나머지 상세부는 포함한 검색어를 반환합니다.

Excel H열 원주소를 I열 검색어로 변환:

```bash
address-cleaner excel input.xlsx -o output.xlsx --source-col H --target-col I
```

로컬 정제 기준으로 잘못된 원주소를 M열에 표시:

```bash
address-cleaner excel input.xlsx -o output.xlsx --source-col H --target-col I --status-col M
```

API 키가 활성화된 뒤 실제 검색 결과까지 M열에 표시합니다. 기본 provider는 `both`입니다.

```bash
export JUSO_CONFIRM_KEY='...'
export EPOST_SERVICE_KEY='...'
address-cleaner excel input.xlsx -o output.xlsx --source-col H --target-col I --status-col M --mark-missing
```

Juso 키는 모든 모드에서 `JUSO_CONFIRM_KEY`, `JUSO_CONFM_KEY`, `JUSO_API_KEY` 어느 이름으로 설정해도 인식합니다.

특정 API만 검증할 수도 있습니다 (`--provider juso|epost|both|none`).

검증 상세를 함께 남기려면 `--detail-col`을 추가합니다(예: N열). 어떤 검색어로 몇 건이 나왔는지, 1건으로 확정되면 표준 도로명주소와 우편번호까지 기록되어 사람이 보완하거나 후단 자동화 실패(우편번호 미확정 등)를 사전 점검할 때 바로 쓸 수 있습니다.

```bash
address-cleaner excel input.xlsx -o output.xlsx --source-col H --target-col I --status-col M --detail-col N --mark-missing
```

상태 열 표시값:

- `검색주소없음`: 원주소가 잘못됐거나 API 검색 결과가 0건입니다.
- `2건이상검색`: API 검색 결과가 2건 이상이라 정확히 하나로 확정하기 어렵습니다.
- 빈칸: 검색주소가 하나로 확인됐거나, API 검증을 하지 않았지만 로컬 정제 기준으로는 검색 가능한 주소입니다.

검증 방식:

- `juso`: 상세부를 포함한 정제 검색어로 1차 검증하고, 0건이면 상세부를 뗀 주소 골격(시도~지번/건물번호)으로 한 번 더 검증합니다. 어느 단계에서 몇 건이 나왔는지는 `--detail-col`에 남습니다.
- `epost`: 상세부 포함 검색어가 맞지 않으면 우정사업본부 검색 특성에 맞춰 `도로명+건물번호` 또는 `동+번지` 형태의 짧은 검색어로 한 번 더 검증합니다.

행안부 [법정동코드 전체자료](https://www.code.go.kr/stdcode/regCodeL.do) 텍스트 파일이 있으면 API 호출 전에 행정구역 실존 여부를 오프라인으로 검사할 수 있습니다(UTF-8/CP949 모두 인식). 존재하지 않는 시군구/법정동 조합은 API를 부르지 않고 바로 `검색주소없음`으로 표시됩니다. 행정동 표기(예: 신정3동)는 법정동 사전에 없어 거짓 양성이 날 수 있으므로 상세 사유를 보고 판단하세요.

```bash
address-cleaner excel input.xlsx -o output.xlsx --source-col H --target-col I --status-col M --detail-col O --admin-dict 법정동코드_전체자료.txt --mark-missing
```

지번주소가 골격 검색까지 0건이면 붙여 쓴 지번 변형(`야당동 5717` → `57-17`)으로 후보를 찾아 검증상세에 제안만 남깁니다. 주소를 자동으로 바꾸지는 않습니다.

`--mark-missing` 검증은 기본 8개 워커로 병렬 처리하되, 등기 모드와 같은 전역 레이트리미터가 provider별 초당 호출 수를 자동으로 제한해 API 차단을 피합니다. 같은 원주소가 반복되는 행은 병렬 제출 전에 하나로 합쳐지고, 검증 이력(`--history`)에 있는 주소는 API를 부르지 않습니다. 직렬로 돌리려면 `--workers 1`을 줍니다.

### 2) 학습 루프: 이력·교정 후보·오타 규칙

```bash
address-cleaner excel input.xlsx -o output.xlsx --source-col H --target-col I \
  --status-col M --detail-col O --mark-missing \
  --history verify_history.sqlite --corrections-out corrections.json
```

- `--history`: 검증 결과를 SQLite에 누적합니다. 기한(기본 14일, `--history-max-age-days`) 내 같은 주소는 API를 부르지 않고 재사용하고, 이전 실행과 판정이 달라진 주소(예: 지난달 1건 → 이번달 0건)는 검증상세에 `⚠ 판정 변경`으로 표시합니다. 행정구역 개편·건물 멸실 신호이므로 사람이 확인해야 합니다.
- `--corrections-out`: 원문 그대로는 검색이 안 되고 골격/지번 변형으로만 통한 주소 쌍을 JSON 리포트로 수확합니다. 사람이 검토해 원주소를 고치거나 `--typo-rules` 규칙으로 승격하면, 오타 사전이 데이터에서 스스로 자랍니다.
- `--typo-rules`: 승격한 오타 교정 규칙 JSON을 일반 정제에 적용합니다(등기 모드와 같은 형식·같은 적용 지점). 수확 → 검토 → 승격 → 적용의 학습 루프가 일반 모드 안에서 완결됩니다.

```bash
# corrections.json 검토 후 규칙으로 승격한 rules.json 적용
# 형식: [["정우팰리스", "정우펠리스"]] 또는 {"replacements": [["...", "..."]]}
address-cleaner excel input.xlsx -o output.xlsx --source-col H --target-col I \
  --status-col M --mark-missing --typo-rules rules.json --corrections-out corrections.json
```

API 키 상태 확인:

```bash
address-cleaner probe epost "경기도 파주시 야당동 57-17"
address-cleaner probe juso "경기도 파주시 야당동 57-17"
```

`probe epost`도 Excel 검증과 동일하게 상세부 포함 검색이 맞지 않으면 짧은 ePost용 검색어로 재시도하고, 실제 사용한 검색어를 `query`로 출력합니다.

### 3) 법원 부동산등기부등본 열람페이지 전체검색용 주소 생성 (registry 모드)

강제경매/우선신청 엑셀에서 법원 부동산등기부등본 열람페이지의 전체검색에 넣을 주소를 만듭니다.

```bash
export JUSO_CONFM_KEY='...'
address-cleaner registry input.xlsx -o out
```

등기소 검색주소 생성 모드는 다음을 수행합니다.

- `최종주소` 또는 `대상 임대차계약 주소` 기준으로 Juso API 1차 검색
- 0건 또는 2건 이상이면 원문/도로명/지번/지번 변형 후보로 2차 좁힘 검색
- 법원 등기부등본 열람페이지 전체검색용 `지번주소 + 건물명 + 동 + 호` 검색어 생성
- `주소검토결과`를 `바로조회가능`, `검토후조회`, `보완필요`로 표기
- 0건/2건 이상 등 정확히 1건으로 확정되지 않은 경우 `JUSO_판정`, `JUSO_2차판정`, `등기_검토사유`, `등기소_검토필요` 시트에 검토 사유를 남김
- 출력 파일과 summary JSON을 `out/` 아래 생성

Juso 1차·2차 검색은 기본 8개 워커로 병렬 처리하되, 전역 레이트리미터가 초당 호출 수를 자동으로 제한해 API 차단을 피합니다. 캐시가 채워진 재실행은 거의 호출이 없고, 콜드 실행의 대기시간만 줄어듭니다. 직렬로 돌리려면 `--workers 1`을 줍니다.

JSON 캐시는 검증 이력과 같은 기준으로 기본 14일이 지나면 만료되어 다시 검색합니다(행정구역 개편·건물 멸실로 검색 결과가 달라질 수 있으므로). 만료 엔트리는 저장 시 파일에서도 정리됩니다. 만료 주기는 `--cache-max-age-days`로 바꾸고, `0`이면 만료 없이 이전 동작을 유지합니다.

```bash
address-cleaner registry input.xlsx -o out --workers 8
```

## 사내(HUG) 워크플로 연동

후단 자동입력(파워쉘)과의 열 규약(H/I/M/N/O), 권장 배치 실행, 실패 사례 환류
루틴, registry 입력 규약 등 사내 연동 문서는
[docs/hug-integration.md](docs/hug-integration.md)로 분리했습니다.

## 참고한 공식/오픈소스 자료

- 도로명주소 `juso.go.kr` 검색 API: `https://eng.juso.go.kr/addrlink/openApi/searchApi.do`
  - `confmKey`, `keyword`, `currentPage`, `countPerPage`, `resultType=json` 요청 구조와 `roadAddr`, `jibunAddr`, `zipNo` 등 응답 필드를 기준으로 구현했습니다.
- 우정사업본부 도로명주소조회서비스: `https://www.data.go.kr/data/15000124/openapi.do`
  - `ServiceKey`, `searchSe`, `srchwrd`, `countPerPage`, `currentPage` 요청 구조를 기준으로 구현했습니다.
- 우정사업본부 통합검색 5자리 우편번호조회서비스: `https://www.data.go.kr/en/data/15056971/openapi.do`
  - `srchwrd` 단일 검색어 기반의 통합 검색 경로를 보조 검토했습니다.
- GitHub `finecodekr/addresskr`: `https://github.com/finecodekr/addresskr`
  - `juso.go.kr` API 키를 환경변수로 주입하고 주소 파싱 결과를 구조화하는 방향을 참고했습니다.

## 릴리스

변경 이력은 [CHANGELOG.md](CHANGELOG.md)에 기록하며, 릴리스는 `v<버전>` 태그로
표시합니다. 릴리스 빌드 워크플로(`release-build`)는 수동(`workflow_dispatch`)
전용으로 sdist/wheel 아티팩트 업로드까지만 수행합니다. 태그 생성·PyPI 배포는
[CONTRIBUTING](CONTRIBUTING.md)의 승인 게이트에 따라 별도 승인 사항입니다.

## 보안 원칙

- API 키는 `.env` 또는 쉘 환경변수로만 둡니다.
- 원본 legacy 스크립트에 있던 하드코딩 키는 레포에 커밋하지 않습니다.
- `.env`, 로그, Excel 산출물은 기본적으로 Git 추적에서 제외합니다.
- 우정사업본부(ePost) 엔드포인트는 현재 기본값이 `http://`라 `ServiceKey`가 평문으로 전송됩니다. `EPOST_ENDPOINT` 환경변수로 엔드포인트를 바꿀 수 있으므로, 운영망에서 아래처럼 https 응답을 먼저 확인한 뒤 `.env`에 고정하세요. 확인되면 `clients.py`의 `EPOST_DEFAULT_ENDPOINT` 기본값도 https로 교체합니다. (2026-07-10 개발 환경에서는 프록시가 해당 도메인을 차단해 https 전환 여부를 검증하지 못했습니다 — 운영망 재확인 필요)

  ```bash
  EPOST_ENDPOINT='https://openapi.epost.go.kr/postal/retrieveNewAdressAreaCdService/retrieveNewAdressAreaCdService/getNewAddressListAreaCd' \
    address-cleaner probe epost "경기도 파주시 야당동 57-17"
  ```

## Public source visibility boundary

This repository is being prepared for possible public source visibility. A
public repository setting would be source-only: it would not approve release or
tag creation, package/image publication, production deploy/restart/reload,
database mutation, provider or Telegram sends, credential movement, history
rewrite, or any other live operation.

Runtime credentials and private operational data must stay outside the
repository. Example configuration must use placeholders only.
