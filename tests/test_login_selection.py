import tempfile
import unittest
from pathlib import Path
from unittest.mock import MagicMock, patch

import nmis_slip_ui as ui
from local_profiles import ProfileStore


class LoginSelectionTests(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory()
        self.addCleanup(self.temp.cleanup)
        self.store = ProfileStore(Path(self.temp.name) / "profiles.json", lambda s: "enc:" + s, lambda s: s[4:])
        self.first = self.store.save_account(0, "first", "pw1")
        self.second = self.store.save_account(1, "second", "pw2")
        self.store.activate(self.first["id"])
        self.store.set_checked(self.second["id"], True)
        self.store.save_group(self.first["id"], 0, {"first-member": {}})
        self.store.save_group(self.second["id"], 0, {"second-member": {}})
        for name, value in (("PROFILE_STORE", self.store), ("NMIS_USER_ID", "first"), ("NMIS_PASSWORD", "pw1"), ("CONNECTION_BUSY", False)):
            patcher = patch.object(ui, name, value)
            patcher.start()
            self.addCleanup(patcher.stop)
        self.app = MagicMock()
        self.app.distribution_busy = False
        self.app.account_switch_busy = False
        self.app._choose_login_profile.return_value = self.second
        self.app.after.side_effect = lambda delay, callback: callback()

    def run_connection(self, error=None):
        with patch.object(ui, "cdp_is_ready", return_value=True), patch.object(ui.threading, "Thread") as thread, patch.object(ui, "sync_playwright"), patch.object(ui, "_find_nmis_page", return_value=object()), patch.object(ui, "require_session_account", side_effect=error) as login, patch.object(ui, "save_settings"), patch.object(ui.time, "sleep"), patch.object(ui.messagebox, "showerror"):
            ui.ModernSlipUI.start_attachable_chrome(self.app)
            self.assertTrue(ui.CONNECTION_BUSY)
            thread.call_args.kwargs["target"]()
        return login

    def test_verified_choice_becomes_account_for_all_operations_and_groups(self):
        login = self.run_connection()
        self.assertEqual(login.call_args.args[1:3], ("second", "pw2"))
        self.assertTrue(login.call_args.kwargs["switch"])
        self.assertEqual((ui.NMIS_USER_ID, ui.NMIS_PASSWORD), ("second", "pw2"))
        self.assertEqual(set(self.store.groups(self.store.active["id"])[0]), {"second-member"})
        self.assertEqual(len(self.store.login_choices), 2)
        self.assertFalse(ui.CONNECTION_BUSY)

    def test_failed_connection_does_not_rebind_existing_account(self):
        self.run_connection(RuntimeError("login failed"))
        self.assertEqual(self.store.active["id"], self.first["id"])
        self.assertEqual(ui.NMIS_USER_ID, "first")
        self.assertFalse(ui.CONNECTION_BUSY)

    def test_cancel_does_not_connect_or_change_account(self):
        self.app._choose_login_profile.return_value = None
        with patch.object(ui, "sync_playwright") as connect:
            ui.ModernSlipUI.start_attachable_chrome(self.app)
        connect.assert_not_called()
        self.assertEqual(self.store.active["id"], self.first["id"])

    def test_one_checked_account_is_chosen_without_dialog(self):
        self.store.set_checked(self.second["id"], False)
        self.assertEqual(ui.ModernSlipUI._choose_login_profile(self.app)["id"], self.first["id"])
