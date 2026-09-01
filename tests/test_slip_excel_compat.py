import unittest
from datetime import datetime
from pathlib import Path
from unittest.mock import patch

from nmis_slip_automation import (
    Transaction,
    _click_step1_save_button,
    _save_and_confirm,
    classify_transaction_type,
    describe_slip_settings,
    parse_excel,
    slip_settings_ready,
    suggest_slip_settings,
    transaction_date_range,
)


class SlipExcelCompatibilityTests(unittest.TestCase):
    def test_transaction_exposes_ui_compatibility_values(self):
        incoming = Transaction(
            source_row=27,
            transacted_at=datetime(2026, 9, 1, 15, 59, 41),
            content="회비",
            withdrawal=None,
            deposit=13000,
            memo="",
            note="",
        )
        outgoing = Transaction(
            source_row=28,
            transacted_at=datetime(2026, 9, 1, 16, 0, 0),
            content="급여",
            withdrawal=100000,
            deposit=None,
            memo="",
            note="",
        )

        self.assertEqual(incoming.raw_id, "tx-27-20260901155941")
        self.assertEqual(incoming.date_str, "2026-09-01 15:59:41")
        self.assertEqual(incoming.direction, "입금")
        self.assertEqual(outgoing.direction, "출금")

    def test_parse_excel_returns_transactions_from_existing_reader(self):
        expected = [object(), object()]
        with patch(
            "nmis_slip_automation.read_excel",
            return_value=(datetime(2026, 9, 1).date(), datetime(2026, 9, 1).date(), expected),
        ) as mocked_reader:
            result = parse_excel(Path("sample.xls"))

        self.assertIs(result, expected)
        mocked_reader.assert_called_once_with(Path("sample.xls"))

    def test_date_range_uses_earliest_and_latest_transaction_dates(self):
        transactions = [
            Transaction(2, datetime(2026, 9, 1, 9), "", None, 1, "", ""),
            Transaction(3, datetime(2026, 8, 7, 18), "", 1, None, "", ""),
            Transaction(4, datetime(2026, 8, 20, 12), "", None, 1, "", ""),
        ]

        earliest, latest = transaction_date_range(transactions)

        self.assertEqual(earliest.isoformat(), "2026-08-07")
        self.assertEqual(latest.isoformat(), "2026-09-01")

    def test_cms_from_central_association_is_automatically_member_fee(self):
        transaction = Transaction(
            27,
            datetime(2026, 8, 7, 2),
            "한국외식업중앙회",
            None,
            2380940,
            "CMS",
            "4913-001",
        )

        result = classify_transaction_type(transaction, {"회비": "member_fee"})

        self.assertEqual(result, "member_fee")

    def test_direct_collection_deposit_is_automatically_member_fee(self):
        transaction = Transaction(
            18,
            datetime(2026, 8, 31, 13, 35, 26),
            "직접수금 33건",
            None,
            396000,
            "현금",
            "4913-001",
        )

        result = classify_transaction_type(transaction, {})
        settings = suggest_slip_settings(transaction, result)

        self.assertEqual(result, "member_fee")
        self.assertEqual(settings["account_code"], "5141")
        self.assertTrue(slip_settings_ready(transaction, settings))

    def test_keywords_search_content_memo_and_note(self):
        salary = Transaction(1, datetime(2026, 8, 1), "8월 지급", 100, None, "급여", "")
        join_fee = Transaction(2, datetime(2026, 8, 2), "입금", None, 100, "", "신규 가입금")
        rules = {"급여": "salary", "가입금": "join_fee"}

        self.assertEqual(classify_transaction_type(salary, rules), "salary")
        self.assertEqual(classify_transaction_type(join_fee, rules), "join_fee")

    def test_member_and_join_fee_account_codes_are_restored(self):
        member = Transaction(1, datetime(2026, 8, 1), "월회비", None, 13000, "", "")
        joining = Transaction(2, datetime(2026, 8, 2), "가입금", None, 100000, "", "")

        member_settings = suggest_slip_settings(member, "member_fee")
        join_settings = suggest_slip_settings(joining, "join_fee")

        self.assertEqual(member_settings["account_code"], "5141")
        self.assertEqual(join_settings["account_code"], "5102")
        self.assertTrue(slip_settings_ready(member, member_settings))
        self.assertTrue(slip_settings_ready(joining, join_settings))

    def test_cms_requires_count_and_fee_before_registration(self):
        cms = Transaction(
            3,
            datetime(2026, 8, 7),
            "한국외식업중앙회",
            None,
            2380940,
            "CMS",
            "",
        )
        settings = suggest_slip_settings(cms, "member_fee")

        self.assertEqual(settings["mode"], "cms_bundle")
        self.assertFalse(slip_settings_ready(cms, settings))
        self.assertIn("5141 회비 + 4385 잡비", describe_slip_settings(cms, settings))

        settings.update(cms_count=190, cms_fee=19000)
        self.assertTrue(slip_settings_ready(cms, settings))

    def test_salary_is_split_into_known_account_codes(self):
        salary = Transaction(4, datetime(2026, 8, 25), "급여", 1360000, None, "", "")

        settings = suggest_slip_settings(salary, "salary")

        self.assertEqual(settings["mode"], "salary_bundle")
        self.assertEqual(settings["basic_account"], "4326")
        self.assertEqual(settings["bonus_account"], "4403")
        self.assertEqual(settings["basic_pay"], 860000)
        self.assertEqual(settings["bonus_pay"], 500000)
        self.assertTrue(slip_settings_ready(salary, settings))

    def test_general_keyword_recovers_communication_account(self):
        transaction = Transaction(5, datetime(2026, 8, 10), "KT 통신요금", 55000, None, "", "")

        settings = suggest_slip_settings(transaction, "기타")

        self.assertEqual(settings["account_code"], "4408")
        self.assertEqual(settings["account_name"], "통신운반비")
        self.assertTrue(slip_settings_ready(transaction, settings))

    def test_unknown_transaction_is_blocked_until_account_is_set(self):
        transaction = Transaction(6, datetime(2026, 8, 11), "미분류 거래", 1000, None, "", "")

        settings = suggest_slip_settings(transaction, "기타")

        self.assertFalse(slip_settings_ready(transaction, settings))

    def test_slip_save_click_prefers_modal_fnsave_button(self):
        page = unittest.mock.MagicMock()
        save_button = unittest.mock.MagicMock()
        save_button.is_visible.return_value = True
        save_locator = unittest.mock.MagicMock()
        save_locator.count.return_value = 1
        save_locator.nth.return_value = save_button
        page.locator.return_value = save_locator

        clicked = _click_step1_save_button(page, lambda _message: None)

        self.assertTrue(clicked)
        first_selector = page.locator.call_args_list[0].args[0]
        self.assertIn(".modal:visible", first_selector)
        self.assertIn("fnSave", first_selector)
        save_button.click.assert_called_once_with(force=True)

    def test_save_flow_stops_when_modal_register_button_is_missing(self):
        page = unittest.mock.MagicMock()
        with patch("nmis_slip_automation._click_step1_save_button", return_value=False):
            with self.assertRaisesRegex(RuntimeError, "등록창의 '등록\\(fnSave\\)'"):
                _save_and_confirm(page, lambda _message: None)


if __name__ == "__main__":
    unittest.main()
