import openpyxl

from address_cleaner.clients import SearchResult
from address_cleaner.cli import main
from address_cleaner.excel import STATUS_AMBIGUOUS, STATUS_NOT_FOUND, _verify, process_workbook
from address_cleaner.normalizer import normalize_for_search, preprocess_raw_address


def test_lot_address_strips_building_detail():
    result = normalize_for_search("경기도 파주시 야당동 57-17 정우펠리스 제303동 제1층 제101호")
    assert result.query == "경기도 파주시 야당동 57-17"
    assert result.kind == "lot"
    assert result.searchable


def test_lot_address_removes_extra_lot_marker():
    result = normalize_for_search("경기도 파주시 동패동 7외 1필지 노블타운 제102동 제4층 제401호")
    assert result.query == "경기도 파주시 동패동 7"
    assert result.kind == "lot"


def test_road_address_keeps_road_and_building_number():
    result = normalize_for_search("서울특별시 강남구 테헤란로 152 강남파이낸스센터 10층")
    assert result.query == "서울특별시 강남구 테헤란로 152"
    assert result.kind == "road"


def test_preprocess_repairs_missing_spaces():
    assert preprocess_raw_address("경기도파주시야당동57-17") == "경기도 파주시 야당동 57-17"


def test_malformed_address_is_not_searchable():
    result = normalize_for_search("주소 미상")
    assert result.query == ""
    assert result.kind == "invalid"
    assert not result.searchable


def test_unrecognized_address_does_not_emit_fallback_query():
    result = normalize_for_search("경기도 파주시 정우펠리스 제303동 제1층 제101호")
    assert result.query == ""
    assert result.kind == "invalid"
    assert result.status == "unrecognized"


def test_excel_marks_invalid_source_as_not_found(tmp_path):
    input_path = tmp_path / "input.xlsx"
    output_path = tmp_path / "output.xlsx"
    wb = openpyxl.Workbook()
    ws = wb.active
    ws["H1"] = "원주소"
    ws["H2"] = "주소 미상"
    wb.save(input_path)

    stats = process_workbook(input_path, output_path, source_col="H", target_col="I", status_col="N")

    result_wb = openpyxl.load_workbook(output_path)
    result_ws = result_wb.active
    assert result_ws["I2"].value is None
    assert result_ws["N2"].value == STATUS_NOT_FOUND
    assert stats["missing"] == 1


def test_mark_missing_without_api_keys_raises(tmp_path):
    input_path = tmp_path / "input.xlsx"
    output_path = tmp_path / "output.xlsx"
    wb = openpyxl.Workbook()
    ws = wb.active
    ws["H1"] = "원주소"
    ws["H2"] = "경기도 파주시 야당동 57-17 정우펠리스 제303동 제1층 제101호"
    wb.save(input_path)

    try:
        process_workbook(
            input_path,
            output_path,
            source_col="H",
            target_col="I",
            status_col="N",
            mark_missing=True,
        )
    except RuntimeError as exc:
        assert "API key" in str(exc)
    else:
        raise AssertionError("expected RuntimeError")


def test_provider_none_with_mark_missing_does_not_mark_searchable_rows(tmp_path):
    input_path = tmp_path / "input.xlsx"
    output_path = tmp_path / "output.xlsx"
    wb = openpyxl.Workbook()
    ws = wb.active
    ws["H1"] = "원주소"
    ws["H2"] = "경기도 파주시 야당동 57-17 정우펠리스 제303동 제1층 제101호"
    wb.save(input_path)

    stats = process_workbook(
        input_path,
        output_path,
        source_col="H",
        target_col="I",
        status_col="N",
        provider="none",
        mark_missing=True,
    )

    result_wb = openpyxl.load_workbook(output_path)
    result_ws = result_wb.active
    assert result_ws["I2"].value == "경기도 파주시 야당동 57-17"
    assert result_ws["N2"].value is None
    assert stats["missing"] == 0


class _FakeJuso:
    def __init__(self, total_count: int, error: bool = False):
        self.total_count = total_count
        self.error = error

    def search(self, query: str, count: int = 5):
        first = {"errorCode": "E"} if self.error else {}
        return SearchResult("juso", self.total_count, first, {})


class _FakeEpost:
    def __init__(self, total_count: int, error: bool = False):
        self.total_count = total_count
        self.error = error

    def search(self, query: str, search_se: str = "road", count: int = 5):
        first = {"returnCode": "30"} if self.error else {}
        return SearchResult("epost", self.total_count, first, "")


def test_verify_marks_ambiguous_when_multiple_results():
    assert _verify("서울특별시 강남구 테헤란로 152", "road", _FakeJuso(2), None) == "ambiguous"


def test_verify_marks_ambiguous_before_single_match():
    assert _verify("서울특별시 강남구 테헤란로 152", "road", _FakeJuso(1), _FakeEpost(2)) == "ambiguous"


def test_verify_ignores_one_provider_error_when_another_provider_works():
    assert _verify("서울특별시 강남구 테헤란로 152", "road", _FakeJuso(1), _FakeEpost(0, error=True)) == "verified"


def test_verify_raises_when_all_configured_providers_error():
    try:
        _verify("서울특별시 강남구 테헤란로 152", "road", _FakeJuso(0, error=True), None)
    except RuntimeError as exc:
        assert "API errors" in str(exc)
    else:
        raise AssertionError("expected RuntimeError")


def test_status_labels_are_user_facing_korean():
    assert STATUS_NOT_FOUND == "검색주소없음"
    assert STATUS_AMBIGUOUS == "2건이상검색"


def test_cli_prints_clean_error_for_missing_api_keys(tmp_path, capsys):
    input_path = tmp_path / "input.xlsx"
    output_path = tmp_path / "output.xlsx"
    wb = openpyxl.Workbook()
    ws = wb.active
    ws["H1"] = "원주소"
    ws["H2"] = "경기도 파주시 야당동 57-17"
    wb.save(input_path)

    exit_code = main([
        "excel",
        str(input_path),
        "-o",
        str(output_path),
        "--source-col",
        "H",
        "--target-col",
        "I",
        "--status-col",
        "N",
        "--mark-missing",
    ])

    captured = capsys.readouterr()
    assert exit_code == 2
    assert captured.err.startswith("error: ")
