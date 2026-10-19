import unittest
from datetime import datetime
from unittest.mock import MagicMock, patch

from nmis_slip_automation import (
    ACCOUNT_CODES,
    Transaction,
    classify_transaction_type,
    register_transaction_on_nmis,
    slip_settings_ready,
    suggest_slip_settings,
)


class CondolenceAccountTests(unittest.TestCase):
    def test_withdrawal_content_selects_condolence_account_for_registration(self):
        self.assertEqual(ACCOUNT_CODES["4311"], "경조비")
        for content in ("경조, 홍길동", "회원 경조비", "경조 회비", "직원 경조 급여"):
            with self.subTest(content=content):
                tx = Transaction(1, datetime(2026, 10, 19), content, 50000, None, "", "")
                kind = classify_transaction_type(tx, {"회비": "member_fee", "급여": "salary"})
                settings = suggest_slip_settings(tx, kind)
                self.assertEqual(settings["account_code"], "4311")
                self.assertEqual(settings["account_name"], "경조비")
                self.assertEqual(settings["mode"], "single")
                self.assertTrue(slip_settings_ready(tx, settings))
                with patch("nmis_slip_automation.register_single_slip_on_page") as register:
                    register_transaction_on_nmis(MagicMock(), tx, kind, settings)
                self.assertEqual(register.call_args.kwargs["account_code"], "4311")
                self.assertEqual(register.call_args.kwargs["slip_type_label"], "출금전표")

    def test_incoming_condolence_payment_is_not_assigned_expense_account(self):
        tx = Transaction(2, datetime(2026, 10, 19), "경조 입금", None, 50000, "", "")
        self.assertNotEqual(suggest_slip_settings(tx, "기타")["account_code"], "4311")


if __name__ == "__main__":
    unittest.main()
