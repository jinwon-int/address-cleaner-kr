"""실전 실패 사례(우편번호미확정) 환류로 추가된 정제 규칙 테스트.

각 케이스는 파워쉘 주소찾기 팝업에서 '실패(우편번호미확정)'로 끝난 실제
원주소의 표기 구조(오타·동/층/호 형태·지번 형식·건물명 토큰)만 유지하고,
법정동·지번·도로명·건물명은 가상 값(샘플/테스트/견본/모의…)으로 바꾼 것이다.
"""

import openpyxl

from address_cleaner.clients import SearchResult
from address_cleaner.excel import _verify, process_workbook
from address_cleaner.normalizer import (
    base_for_search,
    normalize_for_search,
    normalize_unit_dong,
)


def test_road_bungil_number_stays_glued_to_road_name():
    # '샘플로123번길'을 '샘플로123 번길'로 가르면 도로명이 깨져 검색이 실패한다.
    result = normalize_for_search(
        "경기 수원시 팔달구 샘플로123번길45 (테스트동,SAMPLE TEST 1)--408"
    )
    assert result.kind == "road"
    assert result.query.startswith("경기 수원시 팔달구 샘플로123번길 45")
    assert "--" not in result.query
    assert result.query.endswith("408호")


def test_numbered_dong_ga_stays_one_token():
    # '샘플동4가'는 서수 법정동 한 토큰이다. '샘플동 4가'로 가르면 안 된다.
    result = normalize_for_search(
        "서울 영등포구 샘플동4가 1-9 샘플52 501호 (도로명 : 테스트로12길 3)"
    )
    assert result.kind == "lot"
    assert result.query == "서울 영등포구 샘플동4가 1-9 샘플52 501호"
    assert base_for_search(result.query, result.kind) == "서울 영등포구 샘플동4가 1-9"


def test_et_al_household_marker_is_removed():
    result = normalize_for_search(
        "서울 강서구 테스트동 4321 샘플153 105동 203호 외 3세대"
    )
    assert result.query == "서울 강서구 테스트동 4321 샘플153 105동 203호"


def test_legal_area_description_is_cut_and_unit_recovered():
    result = normalize_for_search(
        "경기 용인시 기흥구 샘플동 123-45외 1필지 테스트빌 3층 151.45㎡ 중 도면표시 "
        "ㄱ,ㄴ,ㄷ,ㄹ,ㅁ,ㅂ,ㅅ,ㄱ의 각 점을 순차로 연결한 선 내 부분 302호 전체 67.74㎡"
    )
    assert result.query == "경기 용인시 기흥구 샘플동 123-45 테스트빌 302호"
    assert result.kind == "lot"


def test_legal_drawing_description_without_unit_leaves_skeleton():
    result = normalize_for_search(
        "경기도 수원시 권선구 견본동 321-100외 1필지 2층 194.49㎡ 중 도면표시 "
        "가,나,다,라,가의 각 점을 순차로 연결한 선내부분 24㎡"
    )
    assert result.query == "경기도 수원시 권선구 견본동 321-100"


def test_dev_district_block_lot_tokens_are_removed():
    result = normalize_for_search(
        "인천 남동구 샘플동 357-1번지 샘플테스트구역씨10블럭 샘플메트로3차테스트타워 A동 1802호"
    )
    assert result.query == "인천 남동구 샘플동 357-1 샘플메트로3차테스트타워 A동 1802호"


def test_dev_district_prefix_glued_to_building_is_removed():
    result = normalize_for_search(
        "인천 남동구 샘플동 357-1 테스트구역씨10블록샘플메트로3차테스트타워 제씨동 제4503호"
    )
    assert (
        result.query
        == "인천 남동구 샘플동 357-1 샘플메트로3차테스트타워 제씨동 제4503호"
    )


def test_apartment_complex_name_with_sindosi_is_not_split():
    result = normalize_for_search(
        "경기도 화성시 동탄구 샘플동 975 동탄2신도시샘플파크33단지 제2468동 제6층 제601호"
    )
    assert (
        result.query
        == "경기도 화성시 동탄구 샘플동 975 동탄2신도시샘플파크33단지 제2468동 제601호"
    )


def test_dong_echo_paren_is_removed():
    result = normalize_for_search("인천 미추홀구 테스트동 246-80 1104 (테스트동)")
    assert result.query == "인천 미추홀구 테스트동 246-80 1104"


def test_dash_dong_orphan_is_removed():
    cases = [
        ("서울 강서구 샘플로 321-45-동 902호", "서울 강서구 샘플로 321-45 902호"),
        ("서울 강서구 샘플로 321-45 -동 802호", "서울 강서구 샘플로 321-45 802호"),
    ]
    for raw, expected in cases:
        result = normalize_for_search(raw)
        assert result.query == expected
        assert result.kind == "road"


def test_letter_dash_unit_becomes_identified_dong():
    result = normalize_for_search("서울 강서구 견본동 357 제비-508호")
    assert result.query == "서울 강서구 견본동 357 비동 508호"


def test_basement_floor_and_dong_typo_are_normalized():
    # '제2등'은 '제2동' 오타, '제지하층'은 층 정보라 제거된다.
    result = normalize_for_search(
        "인천 미추홀구 모의동 135-246 샘플빌라 제2등 제지하층 제비02호"
    )
    assert result.query == "인천 미추홀구 모의동 135-246 샘플빌라 제2동 제비02호"


def test_letter_glued_floor_keeps_dong_and_drops_floor():
    result = normalize_for_search(
        "서울 강서구 테스트동 246-135 샘플하임 비제2층 제204호"
    )
    assert result.query == "서울 강서구 테스트동 246-135 샘플하임 비동 제204호"


def test_officetel_word_glued_to_unit_is_removed():
    result = normalize_for_search(
        "인천 미추홀구 테스트동 98-3 샘플어스테이트 102동 오피스텔410호"
    )
    assert result.query == "인천 미추홀구 테스트동 98-3 샘플어스테이트 102동 410호"


def test_building_name_ending_in_officetel_is_kept():
    result = normalize_for_search("경기 김포시 샘플동 579 샘플오피스텔 205호")
    assert result.query == "경기 김포시 샘플동 579 샘플오피스텔 205호"


def test_duplicated_unit_phrases_are_collapsed():
    result = normalize_for_search(
        "인천 미추홀구 견본동 864-21 제비동 402호 (견본동, 테스트아파트) 제비동동402호"
    )
    assert result.query == "인천 미추홀구 견본동 864-21 제비동 402호 테스트아파트"


def test_equivalent_dong_and_ho_variants_are_kept_once():
    result = normalize_for_search(
        "서울특별시 샘플구 테스트동 123-4 샘플빌라 제101동 제402호 별관 101동 402호"
    )

    assert result.query == (
        "서울특별시 샘플구 테스트동 123-4 샘플빌라 제101동 제402호 별관"
    )
    assert result.kind == "lot"


def test_letter_and_zero_padded_unit_variants_are_kept_once():
    cases = [
        (
            "에이동 B01호 A동 B1호",
            "에이동 B01호",
        ),
        (
            "비동 b01호 B동 B1호",
            "비동 b01호",
        ),
    ]
    for units, expected_units in cases:
        result = normalize_for_search(
            f"서울특별시 샘플구 테스트동 123-4 샘플빌라 {units}"
        )

        assert result.query == (
            f"서울특별시 샘플구 테스트동 123-4 샘플빌라 {expected_units}"
        )


def test_repeated_standalone_legal_dong_is_kept_once():
    result = normalize_for_search(
        "서울특별시 샘플구 테스트동 테스트동 123-4 샘플빌라 101호"
    )

    assert result.query == "서울특별시 샘플구 테스트동 123-4 샘플빌라 101호"
    assert result.kind == "lot"


def test_unit_dedup_does_not_confuse_legal_dong_names_or_embedded_tokens():
    raw = "제기동 기동 영등포동2가 테스트동아파트 101호"

    assert normalize_unit_dong(raw) == raw


def test_unit_dedup_removes_parentheses_left_empty_by_duplicate():
    result = normalize_for_search(
        "서울특별시 샘플구 테스트동 123-4 샘플빌라 101호 (101호)"
    )

    assert result.query == "서울특별시 샘플구 테스트동 123-4 샘플빌라 101호"


def test_bare_unit_then_identified_unit_keeps_identified_form():
    result = normalize_for_search(
        "서울 영등포구 테스트로 9, 604호 (샘플동2가) 1동 604호"
    )
    assert result.query == "서울 영등포구 테스트로 9 1동 604호"


def test_vowel_jamo_residue_after_digit_is_removed():
    result = normalize_for_search("인천 남동구 샘플동 357-1 제1ㅣ동 제1305호")
    assert result.query == "인천 남동구 샘플동 357-1 제1동 제1305호"


def test_sido_typo_gwangyeosit_is_corrected():
    result = normalize_for_search(
        "인천광여깃 남동구 모의동 2468-15 테스트빌라 제1층 제101호"
    )
    assert result.query == "인천광역시 남동구 모의동 2468-15 테스트빌라 제101호"


# ---- 건물명 검색어 (kind="building") ----


def test_building_query_without_lot_number():
    result = normalize_for_search("서울 강북구 샘플동 샘플아이빌 비동 402호")
    assert result.kind == "building"
    assert result.searchable
    assert result.query == "서울 강북구 샘플동 샘플아이빌 비동 402호"
    assert result.detail == "비동 402호"


def test_building_name_is_adopted_from_annotation_paren():
    result = normalize_for_search("서울 강서구 테스트동 B동 503호 (테스트동, 샘플캐슬)")
    assert result.kind == "building"
    assert result.query == "서울 강서구 테스트동 샘플캐슬 B동 503호"


def test_building_name_glued_to_dong_after_sigungu_is_split():
    result = normalize_for_search("인천광역시 미추홀구 테스트동샘플파크빌 306호")
    assert result.kind == "building"
    assert result.query == "인천광역시 미추홀구 테스트동 샘플파크빌 306호"


def test_building_query_keeps_dev_district_out(  # 지구/블럭/로트 잡음 제거와의 결합
):
    result = normalize_for_search(
        "경기 김포시 샘플읍 테스트리 샘플테스트6지구 도시개발사업구역 에이2블럭 "
        "샘플앤테스트시티2단지 213동 405호"
    )
    assert result.kind == "building"
    assert (
        result.query == "경기 김포시 샘플읍 테스트리 샘플앤테스트시티2단지 213동 405호"
    )


def test_unit_only_addresses_stay_unsearchable():
    for raw in [
        "인천 남동구 간석동 503호",
        "경기 고양시 일산동구 풍동 102동 202호",
        "인천 부평구 부평동 5층 504호",
        "경기 김포시 통진읍 303호",
    ]:
        result = normalize_for_search(raw)
        assert result.query == ""
        assert result.kind == "invalid"
        assert result.status == "unrecognized"


def test_base_for_search_building_is_district_and_name():
    assert (
        base_for_search("서울 강북구 샘플동 샘플아이빌 비동 402호", "building")
        == "서울 강북구 샘플동 샘플아이빌"
    )


def test_verify_building_falls_back_to_district_and_name_base():
    class _BaseOnlyJuso:
        key = "test-key"

        def __init__(self):
            self.queries = []

        def search(self, query: str, count: int = 5):
            self.queries.append(query)
            if query == "서울 강북구 샘플동 샘플아이빌":
                return SearchResult(
                    "juso",
                    1,
                    {"roadAddr": "서울특별시 강북구 샘플로 100", "zipNo": "00100"},
                    {},
                )
            return SearchResult("juso", 0, {}, {})

    juso = _BaseOnlyJuso()
    status, detail, correction = _verify(
        "서울 강북구 샘플동 샘플아이빌 비동 402호", "building", juso, None
    )

    assert status == "verified"
    assert juso.queries == [
        "서울 강북구 샘플동 샘플아이빌 비동 402호",
        "서울 강북구 샘플동 샘플아이빌",
    ]
    assert correction["type"] == "상세부제거"
    assert correction["working"] == "서울 강북구 샘플동 샘플아이빌"


# ---- 검증 확정 골격으로 검색어 교체 ----


def _base_only_workbook(tmp_path):
    input_path = tmp_path / "input.xlsx"
    wb = openpyxl.Workbook()
    ws = wb.active
    ws["H1"] = "원주소"
    ws["H2"] = "경기도 고양시 일산동구 견본동 98-7 테스트타운 제비동 제101호"
    wb.save(input_path)
    return input_path


class _BaseOnlySkeletonJuso:
    """상세 포함 검색은 0건, 골격 검색만 1건을 돌려주는 가짜 Juso."""

    key = "test-key"

    def search(self, query: str, count: int = 5):
        if query == "경기도 고양시 일산동구 견본동 98-7":
            return SearchResult(
                "juso",
                1,
                {"roadAddr": "경기도 고양시 일산동구 샘플로 123", "zipNo": "00401"},
                {},
            )
        return SearchResult("juso", 0, {}, {})


def test_process_workbook_rewrites_query_with_verified_skeleton(tmp_path, monkeypatch):
    monkeypatch.setattr("address_cleaner.excel.JusoClient", _BaseOnlySkeletonJuso)
    input_path = _base_only_workbook(tmp_path)
    output_path = tmp_path / "out.xlsx"

    stats = process_workbook(
        input_path,
        output_path,
        source_col="H",
        target_col="I",
        status_col="M",
        detail_col="N",
        provider="juso",
        mark_missing=True,
    )

    result_ws = openpyxl.load_workbook(output_path).active
    # 상세 포함 검색어는 팝업에서 0건이 확정이므로, 통한 골격으로 교체된다.
    assert result_ws["I2"].value == "경기도 고양시 일산동구 견본동 98-7"
    assert result_ws["M2"].value in (None, "")
    assert "골격으로 교체" in result_ws["N2"].value
    assert stats["query_rewritten"] == 1
    assert stats["verified"] == 1


def test_process_workbook_keeps_detail_query_when_disabled(tmp_path, monkeypatch):
    monkeypatch.setattr("address_cleaner.excel.JusoClient", _BaseOnlySkeletonJuso)
    input_path = _base_only_workbook(tmp_path)
    output_path = tmp_path / "out.xlsx"

    stats = process_workbook(
        input_path,
        output_path,
        source_col="H",
        target_col="I",
        status_col="M",
        provider="juso",
        mark_missing=True,
        rewrite_working_query=False,
    )

    result_ws = openpyxl.load_workbook(output_path).active
    assert (
        result_ws["I2"].value
        == "경기도 고양시 일산동구 견본동 98-7 테스트타운 제비동 제101호"
    )
    assert stats["query_rewritten"] == 0
