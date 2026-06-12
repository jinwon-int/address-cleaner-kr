# 주소정제

HUG 강제경매/사전심사 Excel 주소를 `juso.go.kr` 또는 우정사업본부 도로명주소조회서비스에서 검색 가능한 형태로 정제하는 Python 애플리케이션입니다.

## 하는 일

- H열 같은 원주소에서 우편번호, `외 N필지`, 반복 주소, 깨진 공백 등 API 검색을 방해하는 잡음을 정리합니다.
- I열에는 정제된 주소를 쓰되, 건물명·동/호 같은 상세부는 보존하고 `제1층` 같은 층 정보만 제거합니다.
- 원주소가 잘못됐거나 주소 골격이 불명확하면 과감하게 검색어를 비우고 M열에 `검색주소없음`을 표시합니다.
- API 키가 활성화된 뒤 실제 검색 결과가 2건 이상이면 M열에 `2건이상검색`을 표시합니다.
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

특정 API만 검증할 수도 있습니다.

```bash
address-cleaner excel input.xlsx -o output.xlsx --source-col H --target-col I --status-col M --provider juso --mark-missing
address-cleaner excel input.xlsx -o output.xlsx --source-col H --target-col I --status-col M --provider epost --mark-missing
address-cleaner excel input.xlsx -o output.xlsx --source-col H --target-col I --status-col M --provider both --mark-missing
```

검증 상세를 함께 남기려면 `--detail-col`을 추가합니다(예: N열). 어떤 검색어로 몇 건이 나왔는지, 1건으로 확정되면 표준 도로명주소와 우편번호까지 기록되어 사람이 보완하거나 후단 자동입력 실패(우편번호 미확정 등)를 사전 점검할 때 바로 쓸 수 있습니다.

```bash
address-cleaner excel input.xlsx -o output.xlsx --source-col H --target-col I --status-col M --detail-col N --mark-missing
```

M열 표시값:

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

검증 이력과 교정 후보까지 켠 권장 배치 실행:

```bash
address-cleaner excel input.xlsx -o output.xlsx --source-col H --target-col I \
  --status-col M --detail-col O --mark-missing \
  --history verify_history.sqlite --corrections-out corrections.json
```

- `--history`: 검증 결과를 SQLite에 누적합니다. 기한(기본 14일, `--history-max-age-days`) 내 같은 주소는 API를 부르지 않고 재사용하고, 이전 실행과 판정이 달라진 주소(예: 지난달 1건 → 이번달 0건)는 검증상세에 `⚠ 판정 변경`으로 표시합니다. 행정구역 개편·건물 멸실 신호이므로 사람이 확인해야 합니다.
- `--corrections-out`: 원문 그대로는 검색이 안 되고 골격/지번 변형으로만 통한 주소 쌍을 JSON 리포트로 수확합니다. 사람이 검토해 원주소를 고치거나 `--typo-rules` 규칙으로 승격하면, 오타 사전이 데이터에서 스스로 자랍니다.

### 1-1) 파워쉘 실패 환류 리포트

후단 자동입력(파워쉘)이 처리결과를 기록한 엑셀에서 실패 행만 모아 재정제 리포트를 만듭니다.

```bash
address-cleaner feedback result.xlsx -o feedback.json
```

- M열이 `실패(...)`인 행을 실패 유형별로 집계하고, 원주소를 **현재 규칙으로 다시 정제**해 봅니다.
- `requeryChanged: true`인 행은 규칙이 그동안 개선돼 I열만 갱신하면 되는 행입니다.
- 재정제해도 같은 검색어가 나오는 행이 새 정제 규칙이 필요한 사례이며, `psDetail`(전산 반응)과 함께 규칙 보강의 입력이 됩니다.

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

## 후단 자동입력(HUG 차세대 파워쉘)과의 연동 열 규약

이 레포의 출력 엑셀은 사내 파워쉘 자동입력 스크립트(hug_auto)가 그대로 읽습니다. 열 배치가 암묵 규약이므로 변경 시 양쪽을 함께 맞춰야 합니다.

| 열 | 쓰는 쪽 | 내용 |
|---|---|---|
| H | (원본) | 원주소 |
| I | 이 레포 | 정제된 검색어 (파워쉘이 전산에 입력) |
| M | 양쪽 공유 | 이 레포: `검색주소없음`/`2건이상검색` → 파워쉘이 이 값을 보면 자동입력 제외. 파워쉘: `완료`/`실패(...)`/`스킵(...)` 처리결과 |
| N | 파워쉘 | 처리상세 (전산 팝업 문구 등) |
| O | 이 레포 | 검증상세 (`--detail-col O` 권장: 검색 경로/건수/표준주소/우편번호) |

## 실패 사례 환류

파워쉘 N열에 남은 전산 실패(예: `실패(우편번호미확정)`)는 정제 규칙의 구멍을 뜻합니다. 환류 루틴:

1. 실패 행의 원주소(H)와 전산 반응(N)을 확보한다.
2. 원인 패턴을 `tests/`에 실패 케이스로 먼저 추가한다 (예: `제1(상층하층)층` 복층 표기).
3. `normalizer.py`(일반)·`registry/normalize.py`(등기) 규칙을 보강해 테스트를 통과시킨다.
4. 시/도 명칭 등 공용 상수는 `regions.py`가 단일 출처이므로 그쪽만 고친다.

## 보안 원칙

- API 키는 `.env` 또는 쉘 환경변수로만 둡니다.
- 원본 legacy 스크립트에 있던 하드코딩 키는 레포에 커밋하지 않습니다.
- `.env`, 로그, Excel 산출물은 기본적으로 Git 추적에서 제외합니다.
- 우정사업본부(ePost) 엔드포인트는 현재 `http://`입니다. 운영망에서 `https://openapi.epost.go.kr` 응답이 확인되면 `clients.py`의 엔드포인트를 https로 전환하세요.
