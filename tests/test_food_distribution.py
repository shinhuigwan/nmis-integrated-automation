import unittest
import random
from datetime import date, timedelta
from unittest.mock import MagicMock, patch

from food_distribution import choose_rows, validate_request, verify_prepared, save_distribution, saved_groups_plan, prepare_saved_distribution


class DistributionTests(unittest.TestCase):
    def rows(self):
        return [dict(memberId=str(i), memberTypeName="정회원" if i % 2 else "비회원",
                     distributionYn="N", distributionDate=None) for i in range(301)]

    def test_selects_exact_count_for_each_member_filter(self):
        for group in ("전체", "정회원", "비회원"):
            rows = self.rows()
            ids = choose_rows(rows, group, 100)
            self.assertEqual(len(ids), 100)
            if group != "전체":
                self.assertTrue(all(rows[int(key)]["memberTypeName"] == group for key in ids))
            self.assertTrue(all(r["distributionYn"] == "N" for r in rows))

    def test_excludes_already_distributed_or_dated_rows(self):
        rows = self.rows()
        rows[0]["distributionYn"] = "Y"
        rows[1]["distributionDate"] = "20261104"
        ids = choose_rows(rows, "전체", 100)
        self.assertNotIn("0", ids)
        self.assertNotIn("1", ids)

    def test_random_sample_uses_entire_eligible_population_without_duplicates(self):
        rows = self.rows()
        expected_pool = [str(i) for i in range(301) if i % 2]
        ids = choose_rows(rows, "정회원", 100, rng=random.Random(7))
        self.assertEqual(ids, random.Random(7).sample(expected_pool, 100))
        self.assertEqual(len(set(ids)), 100)
        self.assertTrue(any(int(key) > 200 for key in ids))
        self.assertNotEqual(set(ids), set(expected_pool[:100]))

    def test_shortage_and_duplicate_ids_fail_before_editing(self):
        with self.assertRaisesRegex(ValueError, "부족"):
            choose_rows(self.rows(), "정회원", 200)
        rows = self.rows()
        rows[1]["memberId"] = rows[0]["memberId"]
        with self.assertRaisesRegex(ValueError, "중복"):
            choose_rows(rows, "전체", 2)

    def test_validation(self):
        self.assertEqual(validate_request("전체", "100", "2026-11-04", "2026-11"),
                         (100, "20261104", "202611"))
        for count in ("0", "-1", "1.5", "50001"):
            with self.assertRaises(ValueError):
                validate_request("전체", count, "20261104", "202611")
        with self.assertRaises(ValueError):
            validate_request("전체", 1, "20260230", "202602")
        with self.assertRaisesRegex(ValueError, "미래"):
            validate_request("전체", 1, (date.today() + timedelta(days=1)).strftime("%Y%m%d"), "202611")

    def test_stale_or_extra_selection_blocks_save(self):
        state = dict(month="202611", memberType="", rows=[dict(memberId="1", selected="true", distributionYn="Y", distributionDate="20261104")])
        verify_prepared(state, ["1"], "20261104", "202611", "")
        state["rows"].append(dict(state["rows"][0], memberId="2"))
        with self.assertRaises(RuntimeError):
            verify_prepared(state, ["1"], "20261104", "202611", "")
        page = MagicMock()
        page.evaluate.return_value = state
        with self.assertRaises(RuntimeError):
            save_distribution(page, dict(ids=["1"], raw_date="20261104", raw_month="202611", member_code=""))
        page.locator.assert_not_called()

    def test_selected_groups_keep_their_own_dates_and_ignore_unselected_groups(self):
        groups = [{"m1": {}}, {"m2": {}}, {"m3": {}}, {}, {}]
        dates = ["20261101", "invalid-unselected-date", "20261103", "", ""]
        plan = saved_groups_plan(groups, dates, [0, 2], "전체", "202611")
        self.assertEqual(plan["dates_by_id"], {"m1": "20261101", "m3": "20261103"})
        self.assertEqual([group["index"] for group in plan["details"]], [0, 2])
        with self.assertRaises(ValueError):
            saved_groups_plan(groups, dates, [], "전체", "202611")
        with self.assertRaisesRegex(ValueError, "배포4"):
            saved_groups_plan(groups, dates, [3], "전체", "202611")
        with self.assertRaisesRegex(ValueError, "중복"):
            saved_groups_plan([{"same": {}}, {}, {"same": {}}, {}, {}], dates, [0, 2], "전체", "202611")

    def test_multi_group_prepare_applies_dates_in_one_batch_and_reports_skips(self):
        groups = [{"m1": {}, "m4": {}}, {"m2": {}}, {"m3": {}}, {}, {}]
        dates = ["20261101", "20261102", "20261103", "", ""]
        before = dict(rows=[dict(memberId=f"m{i}", memberTypeName="정회원", distributionYn="Y" if i == 4 else "N", distributionDate=None) for i in range(1, 5)])
        after = dict(month="202611", memberType="", rows=[dict(memberId="m1", selected="true", distributionYn="Y", distributionDate="20261101"), dict(memberId="m3", selected="true", distributionYn="Y", distributionDate="20261103")])
        page = MagicMock()
        with patch("food_distribution.lookup_distribution", return_value=before), patch("food_distribution.read_state", return_value=after):
            result = prepare_saved_distribution(page, "전체", groups, dates, [0, 2], "202611", lambda message: None)
        self.assertEqual(result["dates_by_id"], {"m1": "20261101", "m3": "20261103"})
        self.assertEqual(page.evaluate.call_args.args[1]["datesById"], result["dates_by_id"])
        self.assertIn("배포1: 1곳", result["summary"])
        self.assertIn("제외 1곳", result["summary"])
        self.assertNotIn("배포2", result["summary"])

    def test_multi_date_mismatch_blocks_save(self):
        state = dict(month="202611", memberType="", rows=[dict(memberId="m1", selected="true", distributionYn="Y", distributionDate="20261101"), dict(memberId="m3", selected="true", distributionYn="Y", distributionDate="20261103")])
        mapping = {"m1": "20261101", "m3": "20261103"}
        verify_prepared(state, ["m1", "m3"], mapping, "202611", "")
        state["rows"][1]["distributionDate"] = "20261105"
        page = MagicMock()
        page.evaluate.return_value = state
        with self.assertRaises(RuntimeError):
            save_distribution(page, dict(ids=["m1", "m3"], dates_by_id=mapping, raw_date="", raw_month="202611", member_code=""))
        page.locator.assert_not_called()

    def test_invalid_saved_date_prevents_any_query_or_edit(self):
        page = MagicMock()
        with patch("food_distribution.lookup_distribution") as query:
            with self.assertRaisesRegex(ValueError, "배포1"):
                prepare_saved_distribution(page, "전체", [{"m1": {}}, {}, {}, {}, {}], [""] * 5, [0], "202611")
        query.assert_not_called()
        page.evaluate.assert_not_called()
