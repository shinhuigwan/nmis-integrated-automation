import tempfile
import unittest
from datetime import date
from pathlib import Path
from unittest.mock import MagicMock, patch

from openpyxl import Workbook

from membership_fee_register import (
    _can_fill_input,
    _month_value_matches,
    FeePaymentItem,
    ensure_nmis_logged_in,
    find_monthly_fee_row_index,
    find_member_monthly_fee_row_index,
    navigate_to_fee_management,
    parse_fee_payment_excel,
    register_fee_payment_on_nmis,
)


class _AttributeLocator:
    def __init__(self, **attributes):
        self.attributes = attributes

    def get_attribute(self, name):
        return self.attributes.get(name)


class _FakeLocator:
    def __init__(self, page, kind):
        self.page = page
        self.kind = kind

    @property
    def first(self):
        return self

    def count(self):
        if self.kind in {"login_id", "login_password"}:
            return 1 if self.page.login_visible else 0
        if self.kind == "marker":
            return 1 if self.page.ready else 0
        if self.kind == "submenu":
            return 1 if self.page.submenu_visible else 0
        if self.kind == "top":
            return 0 if self.page.login_visible else 1
        return 0

    def click(self, force=False):
        if self.kind == "top":
            self.page.top_clicked = True
            self.page.submenu_visible = True
        elif self.kind == "submenu":
            self.page.submenu_clicked = True
            self.page.ready = True

    def fill(self, value):
        self.page.filled[self.kind] = value

    def press(self, key):
        if self.kind == "login_password" and key == "Enter":
            self.page.login_visible = False

    def wait_for(self, state="visible", timeout=0):
        if self.count() == 0:
            raise AssertionError(f"{self.kind} was not visible")


class _FakeFeePage:
    def __init__(self):
        self.ready = False
        self.login_visible = False
        self.state_route_supported = True
        self.state_route_called = False
        self.submenu_visible = False
        self.top_clicked = False
        self.submenu_clicked = False
        self.filled = {}

    def locator(self, selector):
        if selector in {"#feecontrollist", "#feecontrollist:visible"}:
            return _FakeLocator(self, "submenu")
        if selector == "#userId:visible":
            return _FakeLocator(self, "login_id")
        if selector == "#userPass:visible":
            return _FakeLocator(self, "login_password")
        if selector == "#menuKFB_BH:visible":
            return _FakeLocator(self, "top")
        if selector.startswith("input[name='stdYearMonth']") or selector.startswith(
            "input[name='ceoMemberName']"
        ):
            return _FakeLocator(self, "marker")
        return _FakeLocator(self, "missing")

    def get_by_role(self, role, name, exact=False):
        if role == "link" and name == "회비" and exact:
            return _FakeLocator(self, "top")
        if role == "link" and name == "회비 및 가입금관리" and exact:
            return _FakeLocator(self, "submenu")
        return _FakeLocator(self, "missing")

    def wait_for_timeout(self, milliseconds):
        pass

    def evaluate(self, script):
        self.state_route_called = True
        if self.state_route_supported:
            self.ready = True
            return {"route": "fee/control/list"}
        return {"route": "", "error": "not found"}


class MembershipFeeRegisterTests(unittest.TestCase):
    def test_fee_search_clicks_query_button_after_month_and_representative(self):
        page = MagicMock()
        month = MagicMock()
        representative = MagicMock()
        missing = MagicMock()
        form = MagicMock()
        search = MagicMock()
        month.first = month
        representative.first = representative
        missing.first = missing
        search.first = search
        month.count.return_value = 1
        representative.count.return_value = 1
        missing.count.return_value = 0
        search.count.return_value = 1
        month.locator.return_value = form
        month.input_value.return_value = "2026-01"
        form.get_by_role.return_value = search

        def locate(selector):
            if "stdYearMonth" in selector:
                return month
            if "ceoMemberName" in selector:
                return representative
            return missing

        page.locator.side_effect = locate
        item = FeePaymentItem(2, date(2026, 1, 10), "김보라", 13000)

        with patch("membership_fee_register._visible_grid_rows_with_amount", return_value=[]):
            result = register_fee_payment_on_nmis(page, item)

        month.fill.assert_called_once_with("202601")
        representative.click.assert_called_once_with(force=True)
        representative.fill.assert_called_once_with("김보라")
        representative.press.assert_called_once_with("Tab")
        search.click.assert_called_once_with(force=True)
        self.assertTrue(result["missing_month_fee"])

    def test_month_value_verification_accepts_formatting_but_rejects_wrong_month(self):
        expected = date(2026, 1, 10)
        self.assertTrue(_month_value_matches("2026-01", expected))
        self.assertTrue(_month_value_matches("202601", expected))
        self.assertFalse(_month_value_matches("2026-02", expected))
        self.assertFalse(_month_value_matches("", expected))

    def test_readonly_or_disabled_search_inputs_are_not_filled(self):
        self.assertFalse(_can_fill_input(_AttributeLocator(readonly="readonly")))
        self.assertFalse(_can_fill_input(_AttributeLocator(disabled="disabled")))
        self.assertFalse(_can_fill_input(_AttributeLocator(**{"aria-disabled": "true"})))
        self.assertTrue(_can_fill_input(_AttributeLocator()))

    def test_expired_nmis_session_is_logged_in_before_menu_navigation(self):
        page = _FakeFeePage()
        page.login_visible = True
        logs = []

        logged_in = ensure_nmis_logged_in(page, "saved-id", "saved-password", logs.append)

        self.assertTrue(logged_in)
        self.assertFalse(page.login_visible)
        self.assertEqual(page.filled["login_id"], "saved-id")
        self.assertEqual(page.filled["login_password"], "saved-password")
        self.assertFalse(any("saved-password" in message for message in logs))

    def test_navigation_uses_exact_fee_state_route(self):
        page = _FakeFeePage()
        logs = []

        navigate_to_fee_management(page, logs.append)

        self.assertTrue(page.state_route_called)
        self.assertFalse(page.top_clicked)
        self.assertTrue(page.ready)
        self.assertIn("회비 및 가입금관리 화면 이동 완료", logs[-1])

    def test_navigation_falls_back_to_exact_menu_id(self):
        page = _FakeFeePage()
        page.state_route_supported = False

        navigate_to_fee_management(page)

        self.assertTrue(page.top_clicked)
        self.assertTrue(page.submenu_clicked)
        self.assertTrue(page.ready)

    def test_parse_fee_excel_reads_dates_names_amounts_and_sorts(self):
        with tempfile.TemporaryDirectory() as tmp:
            path = Path(tmp) / "회비입금.xlsx"
            workbook = Workbook()
            sheet = workbook.active
            sheet.title = "거래내역"
            sheet.append(["거래일자", "성명/상호", "금액"])
            sheet.append([date(2026, 8, 25), "홍길동", "15,000원"])
            sheet.append([date(2026, 1, 10), "김보라", 13000])
            workbook.save(path)
            workbook.close()

            result = parse_fee_payment_excel(path)

            self.assertEqual(result.sheet_name, "거래내역")
            self.assertEqual([item.member_name for item in result.items], ["김보라", "홍길동"])
            self.assertEqual(result.items[0].month_text, "2026-01")
            self.assertEqual(result.items[0].month_raw, "202601")
            self.assertEqual(result.items[0].receipt_date_raw, "20260110")
            self.assertEqual(result.items[0].amount, 13000)
            self.assertEqual(result.items[0].excel_row, 3)
            self.assertEqual(result.skipped_rows, [])

    def test_parse_fee_excel_collects_invalid_rows_without_stopping(self):
        with tempfile.TemporaryDirectory() as tmp:
            path = Path(tmp) / "회비입금.xlsx"
            workbook = Workbook()
            sheet = workbook.active
            sheet.title = "거래내역"
            sheet.append(["거래일자", "성명/상호", "금액"])
            sheet.append([date(2026, 1, 10), "", 13000])
            sheet.append([date(2026, 1, 11), "정상회원", 10000])
            workbook.save(path)
            workbook.close()

            result = parse_fee_payment_excel(path)

            self.assertEqual(len(result.items), 1)
            self.assertEqual(result.items[0].member_name, "정상회원")
            self.assertEqual(len(result.skipped_rows), 1)
            self.assertIn("엑셀 2행", result.skipped_rows[0])

    def test_find_monthly_fee_row_ignores_join_fee_only_result(self):
        self.assertIsNone(find_monthly_fee_row_index(["김보라 가입금 0 0 직접 Y"]))
        self.assertEqual(
            find_monthly_fee_row_index(
                ["김보라 가입금 0 0 직접 Y", "김보라 월회비 10,000 0 직접 Y"]
            ),
            1,
        )

    def test_monthly_fee_row_must_belong_to_searched_member(self):
        rows = [
            "김보라 01096015220 가입금 0 0 직접 Y",
            "다른회원 01000000000 월회비 10,000 0 직접 Y",
        ]
        self.assertIsNone(find_member_monthly_fee_row_index(rows, "김보라"))
        rows.append("김보라 01096015220 월회비 10,000 0 직접 Y")
        self.assertEqual(find_member_monthly_fee_row_index(rows, "김보라"), 2)

    def test_fee_type_control_must_confirm_monthly_row(self):
        rows = [
            "김보라 가입금 월회비처럼 보이는 다른 텍스트",
            "김보라 월회비 10,000 0 직접 Y",
        ]
        self.assertEqual(
            find_member_monthly_fee_row_index(rows, "김보라", ["가입금", "월회비"]),
            1,
        )
        self.assertIsNone(
            find_member_monthly_fee_row_index(rows, "김보라", ["가입금", "가입금"])
        )


if __name__ == "__main__":
    unittest.main()
