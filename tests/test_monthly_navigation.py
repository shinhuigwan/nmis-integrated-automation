import tempfile
import unittest
from pathlib import Path
from unittest.mock import MagicMock, patch

import nmis_slip_automation as automation
import nmis_slip_ui as ui


class MonthlyNavigationTests(unittest.TestCase):
    def test_direct_route_opens_form_without_top_menu_click(self):
        page = MagicMock()
        with patch.object(automation, "go_to_angular_state") as navigate:
            automation.navigate_to_settlement_statement(page)
        navigate.assert_called_once_with(page, "account/close/reandex/book/list")
        page.locator.return_value.first.wait_for.assert_called_once_with(state="visible", timeout=15000)
        page.locator.return_value.first.click.assert_not_called()

    def test_failed_direct_route_uses_menu_and_verifies_destination(self):
        page = MagicMock()
        page.locator.return_value.count.return_value = 1
        with patch.object(automation, "go_to_angular_state", side_effect=RuntimeError("unavailable")), patch.object(automation, "wait_for_angular_state") as wait:
            automation.navigate_to_settlement_statement(page)
        wait.assert_called_once_with(page, "account/close/reandex/book/list")
        page.locator.return_value.first.click.assert_called()

    def test_cancel_file_selection_starts_no_work(self):
        app = MagicMock()
        app.monthly_excel_var.get.return_value = ""
        with patch.object(ui.filedialog, "askopenfilename", return_value=""), patch.object(ui.threading, "Thread") as thread:
            ui.ModernSlipUI.start_run_monthly_fill(app)
        thread.assert_not_called()

    def test_selected_workbook_is_passed_to_four_stage_runner(self):
        with tempfile.TemporaryDirectory() as directory:
            workbook = Path(directory) / "chosen.xls"
            workbook.touch()
            app = MagicMock()
            app.monthly_excel_var.get.return_value = ""
            app.member_ym_var.get.return_value = "2026-10"
            with patch.object(ui.filedialog, "askopenfilename", return_value=str(workbook)), patch.object(ui, "cdp_is_ready", return_value=True), patch.object(ui.threading, "Thread") as thread, patch.object(ui, "sync_playwright"), patch.object(ui, "find_nmis_page", return_value=object()), patch.object(ui, "fill_all_monthly_reports_sequentially") as fill:
                ui.ModernSlipUI.start_run_monthly_fill(app)
                thread.call_args.kwargs["target"]()
            self.assertEqual(fill.call_args.kwargs["excel_path"], workbook.resolve())
            app.monthly_excel_var.set.assert_called_once_with(str(workbook.resolve()))
