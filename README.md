# 주소정제

HUG 강제경매/사전심사 Excel 주소를 `juso.go.kr` 또는 우정사업본부 도로명주소조회서비스에서 검색 가능한 형태로 정제하는 Python 애플리케이션입니다.

## 하는 일

- H열 같은 원주소에서 우편번호, `외 N필지`, 반복 주소, 깨진 공백 등 API 검색을 방해하는 잡음을 정리합니다.
- I열에는 정제된 주소를 쓰되, 건물명·동/호 같은 상세부는 보존하고 `제1층` 같은 층 정보만 제거합니다.
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

### 1) 일반 HUG 주소 검색어 정제

단일 주소 정제:

```bash
address-cleaner normalize "경기도 파주시 야당동 57-17 정우펠리스 제303동 제1층 제101호"
```

위 예시는 `경기도 파주시 야당동 57-17 정우펠리스 제303동 제101호`처럼 층 정보만 제거하고 나머지 상세부는 포함한 검색어를 반환합니다.

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

검증 방식:

- `juso`: 상세부를 포함한 정제 검색어로 검증합니다.
- `epost`: 상세부 포함 검색어가 맞지 않으면 우정사업본부 검색 특성에 맞춰 `도로명+건물번호` 또는 `동+번지` 형태의 짧은 검색어로 한 번 더 검증합니다.

API 키 상태 확인:

```bash
address-cleaner probe epost "경기도 파주시 야당동 57-17"
address-cleaner probe juso "경기도 파주시 야당동 57-17"
```

`probe epost`도 Excel 검증과 동일하게 상세부 포함 검색이 맞지 않으면 짧은 ePost용 검색어로 재시도하고, 실제 사용한 검색어를 `query`로 출력합니다.

### 2) 법원 부동산등기부등본 열람페이지 전체검색용 주소 생성

`registry-address-refiner` 레포의 등기소 검색주소 생성 로직은 이 레포로 통합되었습니다. 강제경매/우선신청 엑셀에서 법원 부동산등기부등본 열람페이지의 전체검색에 넣을 주소를 만들 때는 아래 명령을 사용합니다.

```bash
export JUSO_CONFM_KEY='...'
address-cleaner registry input.xlsx -o out
```

기존 자동화와 호환되도록 예전 명령도 유지합니다.

```bash
registry-address-refine input.xlsx -o out
```

입력 파일은 `대상 임대차계약 주소` 컬럼이 필요합니다. 있으면 `최종주소`, `도로명_완성`, `지번_완성`, `원본행`, `법무대리인` 컬럼을 추가로 활용합니다.

등기소 검색주소 생성 모드는 다음을 수행합니다.

- `최종주소` 또는 `대상 임대차계약 주소` 기준으로 Juso API 1차 검색
- 0건 또는 2건 이상이면 원문/도로명/지번/지번 변형 후보로 2차 좁힘 검색
- 법원 등기부등본 열람페이지 전체검색용 `지번주소 + 건물명 + 동 + 호` 검색어 생성
- `주소검토결과`를 `바로조회가능`, `검토후조회`, `보완필요`로 표기
- 0건/2건 이상 등 정확히 1건으로 확정되지 않은 경우 `JUSO_판정`, `JUSO_2차판정`, `등기_검토사유`, `등기소_검토필요` 시트에 검토 사유를 남김
- 출력 파일과 summary JSON을 `out/` 아래 생성

업무 원칙: 정제된 주소는 검색결과가 정확히 1건이어야 합니다. 0건 또는 2건 이상이 불가피한 행은 엑셀에 검토/보완 대상으로 남겨 수동 확인할 수 있게 합니다.

## 보안 원칙

- API 키는 `.env` 또는 쉘 환경변수로만 둡니다.
- 원본 legacy 스크립트에 있던 하드코딩 키는 레포에 커밋하지 않습니다.
- `.env`, 로그, Excel 산출물은 기본적으로 Git 추적에서 제외합니다.
