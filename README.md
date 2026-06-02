# 주소정제

HUG 강제경매/사전심사 Excel 주소를 `juso.go.kr` 또는 우정사업본부 도로명주소조회서비스에서 검색 가능한 형태로 정제하는 Python 애플리케이션입니다.

## 하는 일

- H열 같은 원주소에서 건물명, 동/층/호, `외 N필지` 등 상세부를 제거합니다.
- I열에는 API 검색어로 확실한 `도로명+건물번호` 또는 `동/리+번지`만 씁니다.
- 원주소가 잘못됐거나 주소 골격이 불명확하면 과감하게 검색어를 비우고 N열에 `검색주소없음`을 표시합니다.
- API 키가 활성화된 뒤 실제 검색 결과가 2건 이상이면 N열에 `2건이상검색`을 표시합니다.
- API 키는 코드에 저장하지 않고 환경변수에서 읽습니다.

## 참고한 공식/오픈소스 자료

- 도로명주소 `juso.go.kr` 검색 API: `https://eng.juso.go.kr/addrlink/openApi/searchApi.do`
  - `confmKey`, `keyword`, `currentPage`, `countPerPage`, `resultType=json` 요청 구조와 `roadAddr`, `jibunAddr`, `zipNo` 등 응답 필드를 기준으로 구현했습니다.
- 우정사업본부 도로명주소조회서비스: `https://www.data.go.kr/data/15000124/openapi.do`
  - `ServiceKey`, `searchSe`, `srchwrd`, `countPerPage`, `currentPage` 요청 구조를 기준으로 구현했습니다.
- 우정사업본부 통합검색 5자리 우편번호조회서비스: `https://www.data.go.kr/en/data/15056971/openapi.do`
  - `srchwrd` 단일 검색어 기반의 통합 검색 경로를 보조 검토했습니다.
- GitHub `finecodekr/addresskr`: `https://github.com/finecodekr/addresskr`
  - `juso.go.kr` API 키를 환경변수로 주입하고 주소 파싱 결과를 구조화하는 방향을 참고했습니다.

## 설치

```bash
python -m venv .venv
. .venv/bin/activate
pip install -e '.[dev]'
```

## 사용 예시

단일 주소 정제:

```bash
address-cleaner normalize "경기도 파주시 야당동 57-17 정우펠리스 제303동 제1층 제101호"
```

Excel H열 원주소를 I열 검색어로 변환:

```bash
address-cleaner excel input.xlsx -o output.xlsx --source-col H --target-col I
```

로컬 정제 기준으로 잘못된 원주소를 N열에 표시:

```bash
address-cleaner excel input.xlsx -o output.xlsx --source-col H --target-col I --status-col N
```

API 키가 활성화된 뒤 실제 검색 결과까지 N열에 표시합니다. 기본 provider는 `both`입니다.

```bash
export JUSO_CONFIRM_KEY='...'
export EPOST_SERVICE_KEY='...'
address-cleaner excel input.xlsx -o output.xlsx --source-col H --target-col I --status-col N --mark-missing
```

특정 API만 검증할 수도 있습니다.

```bash
address-cleaner excel input.xlsx -o output.xlsx --source-col H --target-col I --status-col N --provider juso --mark-missing
address-cleaner excel input.xlsx -o output.xlsx --source-col H --target-col I --status-col N --provider epost --mark-missing
address-cleaner excel input.xlsx -o output.xlsx --source-col H --target-col I --status-col N --provider both --mark-missing
```

N열 표시값:

- `검색주소없음`: 원주소가 잘못됐거나 API 검색 결과가 0건입니다.
- `2건이상검색`: API 검색 결과가 2건 이상이라 정확히 하나로 확정하기 어렵습니다.
- 빈칸: 검색주소가 하나로 확인됐거나, API 검증을 하지 않았지만 로컬 정제 기준으로는 검색 가능한 주소입니다.

API 키 상태 확인:

```bash
address-cleaner probe epost "경기도 파주시 야당동 57-17"
address-cleaner probe juso "경기도 파주시 야당동 57-17"
```

## 보안 원칙

- API 키는 `.env` 또는 쉘 환경변수로만 둡니다.
- 원본 legacy 스크립트에 있던 하드코딩 키는 레포에 커밋하지 않습니다.
- `.env`, 로그, Excel 산출물은 기본적으로 Git 추적에서 제외합니다.
