from __future__ import annotations

import unittest
from unittest.mock import patch

import nmis_slip_automation as automation


def valid_report_data() -> dict:
    return {
        "study_cnt": 314,
        "real_cnt": 310,
        "member_cnt": 225,
        "fee_revenue": 0,
        "join_fee_revenue": 0,
        "cms_cnt": 0,
        "direct_cnt": 0,
        "member_source_loaded": True,
        "revenue_source_loaded": True,
        "journal_source_loaded": True,
    }


class MonthlyReportGuardrailTests(unittest.TestCase):
    def test_validate_monthly_report_data_allows_real_zero_amounts(self) -> None:
        automation.validate_monthly_report_data(valid_report_data())

    def test_validate_monthly_report_data_rejects_missing_source(self) -> None:
        data = valid_report_data()
        data["journal_source_loaded"] = False

        with self.assertRaisesRegex(RuntimeError, "금전출납부"):
            automation.validate_monthly_report_data(data)

    def test_validate_monthly_report_data_rejects_zero_core_member_value(self) -> None:
        data = valid_report_data()
        data["member_cnt"] = 0

        with self.assertRaisesRegex(RuntimeError, "회원수"):
            automation.validate_monthly_report_data(data)

    def test_four_stage_runner_reports_exact_failed_stage(self) -> None:
        calls: list[str] = []

        def pass_step1(*args, **kwargs):
            calls.append("step1")
            return {"ok": True}

        def fail_step2(*args, **kwargs):
            calls.append("step2")
            raise RuntimeError("조회 결과 없음")

        def unexpected_step3(*args, **kwargs):
            calls.append("step3")
            return valid_report_data()

        with (
            patch.object(automation, "fill_monthly_report_from_nmis", pass_step1),
            patch.object(automation, "fill_member_status_from_nmis", fail_step2),
            patch.object(automation, "fetch_sheet4_data_only", unexpected_step3),
            self.assertRaisesRegex(RuntimeError, "2/4단계 회원현황 실패"),
        ):
            automation.fill_all_monthly_reports_sequentially(
                object(),
                "unused.xls",
                target_year_month="2026-09",
            )

        self.assertEqual(calls, ["step1", "step2"])


if __name__ == "__main__":
    unittest.main()
