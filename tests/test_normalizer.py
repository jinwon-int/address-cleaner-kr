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

