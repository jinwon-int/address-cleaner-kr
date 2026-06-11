from __future__ import annotations

import json
import tempfile
import unittest
from pathlib import Path

from address_cleaner.registry.juso import make_queries
from address_cleaner.registry.normalize import (
    clean_raw,
    has_extra_parcels,
    load_typo_rules,
    original_is_under_specified,
    parse_lot_addr,
    set_extra_typo_rules,
    strip_detail,
    strip_unit,
    typo_fix,
)


FIXTURES = Path(__file__).parent / "fixtures" / "anonymized_addresses.json"


class AddressRefinementTest(unittest.TestCase):
    @classmethod
    def setUpClass(cls) -> None:
        cls.fixtures = json.loads(FIXTURES.read_text(encoding="utf-8"))

    def test_clean_and_parse_anonymized_lot_addresses(self) -> None:
        for case in self.fixtures:
            with self.subTest(case=case["name"]):
                cleaned = clean_raw(case["raw"])

                self.assertEqual(cleaned, case["expected_clean"])
                self.assertEqual(strip_unit(cleaned), case["expected_unitless"])
                self.assertEqual(parse_lot_addr(cleaned)["addr"], case["expected_lot_addr"])

    def test_make_queries_keeps_lot_lookup_candidate(self) -> None:
        for case in self.fixtures:
            with self.subTest(case=case["name"]):
                queries = make_queries(case["raw"], "", "", "")

                self.assertIn(case["expected_lot_addr"], queries)
                self.assertLessEqual(len(queries), 12)


class TypoFixTest(unittest.TestCase):
    def test_keeps_road_names_starting_with_si(self) -> None:
        self.assertEqual(typo_fix("인천광역시 시민로 100"), "인천광역시 시민로 100")

    def test_collapses_duplicated_si_token(self) -> None:
        self.assertEqual(typo_fix("인천광역시 시 부평구 부평동 100-1"), "인천광역시 부평구 부평동 100-1")

    def test_expands_sido_abbreviation_only_at_start(self) -> None:
        self.assertEqual(typo_fix("서울 강남구 역삼동 123-4"), "서울특별시 강남구 역삼동 123-4")
        self.assertEqual(
            typo_fix("인천광역시 연수구 옥련동 12-3 서울 빌라 101호"),
            "인천광역시 연수구 옥련동 12-3 서울 빌라 101호",
        )

    def test_strips_leading_zip_before_expanding_abbreviation(self) -> None:
        self.assertEqual(typo_fix("12345 경기 김포시 사우동 256-1"), "경기도 김포시 사우동 256-1")


class ExtraTypoRulesTest(unittest.TestCase):
    def tearDown(self) -> None:
        set_extra_typo_rules(None)

    def test_extra_rules_apply_after_loading(self) -> None:
        self.assertEqual(typo_fix("서울특별시 강남구 역삼동 12-3 헬리오시티이"), "서울특별시 강남구 역삼동 12-3 헬리오시티이")

        set_extra_typo_rules([("헬리오시티이", "헬리오시티")])

        self.assertEqual(typo_fix("서울특별시 강남구 역삼동 12-3 헬리오시티이"), "서울특별시 강남구 역삼동 12-3 헬리오시티")

    def test_load_typo_rules_accepts_list_and_dict_forms(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            list_path = Path(tmp) / "rules_list.json"
            list_path.write_text('[["가나다", "가나도"]]', encoding="utf-8")
            dict_path = Path(tmp) / "rules_dict.json"
            dict_path.write_text('{"replacements": [["가나다", "가나도"]]}', encoding="utf-8")

            self.assertEqual(load_typo_rules(list_path), [("가나다", "가나도")])
            self.assertEqual(load_typo_rules(dict_path), [("가나다", "가나도")])

    def test_load_typo_rules_rejects_malformed_entries(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            bad_path = Path(tmp) / "rules_bad.json"
            bad_path.write_text('[["하나만"]]', encoding="utf-8")

            with self.assertRaises(ValueError):
                load_typo_rules(bad_path)


class ParseLotAddrTest(unittest.TestCase):
    def test_mountain_lot_number(self) -> None:
        parsed = parse_lot_addr("서울특별시 관악구 신림동 산 101-1 가나빌라 201호")

        self.assertEqual(parsed["lot"], "산101-1")
        self.assertEqual(parsed["addr"], "서울특별시 관악구 신림동 산101-1")

    def test_sido_omitted_address(self) -> None:
        self.assertEqual(parse_lot_addr("수원시 팔달구 인계동 123-4")["addr"], "수원시 팔달구 인계동 123-4")

    def test_unit_number_is_not_a_lot(self) -> None:
        self.assertEqual(parse_lot_addr("서울특별시 성동구 마장동 801호")["addr"], "")

    def test_dong_without_district_is_not_an_address(self) -> None:
        self.assertEqual(parse_lot_addr("마장동 801")["addr"], "")


class StripDetailTest(unittest.TestCase):
    def test_keeps_lot_only_address(self) -> None:
        self.assertEqual(strip_detail("서울특별시 강남구 역삼동 123"), "서울특별시 강남구 역삼동 123")

    def test_strips_bare_trailing_unit_number(self) -> None:
        self.assertEqual(strip_detail("서울특별시 강남구 역삼동 123-4 1202"), "서울특별시 강남구 역삼동 123-4")

    def test_keeps_road_building_number(self) -> None:
        self.assertEqual(strip_detail("서울특별시 강남구 테헤란로 123 1202"), "서울특별시 강남구 테헤란로 123")


class ExtraParcelsTest(unittest.TestCase):
    def test_detects_extra_parcels(self) -> None:
        self.assertTrue(has_extra_parcels("인천광역시 연수구 옥련동 123-4 외 3필지 101호"))

    def test_ignores_oe_inside_building_name(self) -> None:
        self.assertFalse(has_extra_parcels("인천광역시 연수구 옥련동 123-4 외동마을 101호"))


class UnderSpecifiedTest(unittest.TestCase):
    def test_dong_and_unit_only_is_under_specified(self) -> None:
        self.assertTrue(original_is_under_specified("서울특별시 성동구 마장동 801호"))

    def test_mountain_lot_is_specified(self) -> None:
        self.assertFalse(original_is_under_specified("서울특별시 관악구 신림동 산 101-1 201호"))

    def test_building_name_is_specified(self) -> None:
        self.assertFalse(original_is_under_specified("서울특별시 성동구 마장동 가람빌리지 801호"))


class BackwardCompatImportTest(unittest.TestCase):
    def test_cli_still_re_exports_refactored_functions(self) -> None:
        from address_cleaner.registry.cli import (  # noqa: F401
            clean_raw,
            make_queries,
            parse_lot_addr,
            refine,
            strip_unit,
            typo_fix,
        )


if __name__ == "__main__":
    unittest.main()
