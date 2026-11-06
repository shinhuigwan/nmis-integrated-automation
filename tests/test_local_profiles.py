import json
import random
import tempfile
import unittest
from pathlib import Path
from unittest.mock import MagicMock, patch

from local_profiles import ProfileStore, require_session_account, assign_distribution_members
from food_distribution import choose_rows


class ProfilesTests(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory()
        self.addCleanup(self.temp.cleanup)
        self.path = Path(self.temp.name) / "profiles.json"
        self.protect = lambda value: "encrypted:" + value[::-1]
        self.unprotect = lambda value: value.removeprefix("encrypted:")[::-1]
        self.store = ProfileStore(self.path, self.protect, self.unprotect)

    def reopen(self):
        return ProfileStore(self.path, self.protect, self.unprotect)

    def test_three_accounts_single_activation_and_password_persistence(self):
        records = [self.store.save_account(i, f"user{i}", f"example-password-{i}") for i in range(3)]
        with self.assertRaises(ValueError):
            self.store.save_account(3, "fourth", "password")
        with self.assertRaises(ValueError):
            self.store.save_account(1, "user0", "password")
        self.store.activate(records[0]["id"])
        self.store.activate(records[1]["id"])
        self.assertEqual(self.reopen().credentials(), ("user1", "example-password-1"))
        self.assertNotIn("example-password", self.path.read_text())
        self.store.deactivate(records[0]["id"])
        self.assertEqual(self.store.active["id"], records[1]["id"])
        self.store.deactivate(records[1]["id"])
        self.assertEqual(self.reopen().credentials(), ("", ""))

    def test_migration_does_not_reactivate_disabled_account(self):
        self.store.migrate("existing", "password")
        record = self.store.active
        self.assertEqual(self.store.credentials(), ("existing", "password"))
        self.store.deactivate(record["id"])
        self.store.migrate("existing", "password")
        self.assertIsNone(self.store.active)
        self.assertEqual(len(self.store.accounts), 1)

    def test_account_edit_retains_password_but_new_identity_has_separate_groups(self):
        record = self.store.save_account(0, "old", "password")
        self.store.save_group(record["id"], 0, {"member-1": {"memberName": "업소"}})
        self.store.save_account(0, "old", "")
        self.store.activate(record["id"])
        self.assertEqual(self.store.credentials(), ("old", "password"))
        changed = self.store.save_account(0, "new", "new-password")
        self.assertIsNone(self.store.active)
        self.assertNotEqual(changed["id"], record["id"])
        self.assertEqual(self.store.groups(changed["id"]), [{} for _ in range(5)])

    def test_five_groups_persist_and_are_disjoint_only_within_account(self):
        first = self.store.save_account(0, "one", "password")
        second = self.store.save_account(1, "two", "password")
        self.store.save_group(first["id"], 0, {"m1": {"memberName": "가게1"}})
        with self.assertRaisesRegex(ValueError, "다른 배포"):
            self.store.save_group(first["id"], 1, {"m1": {}})
        self.store.save_group(second["id"], 0, {"m1": {}})
        self.store.save_group(first["id"], 4, {"m5": {}})
        with self.assertRaises(ValueError):
            self.store.save_group(first["id"], 5, {})
        self.assertEqual(len(self.reopen().groups(first["id"])), 5)
        self.assertIn("m1", self.reopen().groups(first["id"])[0])
        self.store.save_group(first["id"], 0, {})
        self.store.save_group(first["id"], 1, {"m1": {}})
        self.assertIn("m1", self.reopen().groups(first["id"])[1])

    def test_random_selection_never_escapes_saved_group(self):
        rows = [dict(memberId=str(i), memberTypeName="정회원", distributionYn="N") for i in range(100)]
        ids = choose_rows(rows, "정회원", 20, allowed_ids=[str(i) for i in range(50)], rng=random.Random(8))
        self.assertEqual(len(set(ids)), 20)
        self.assertTrue(set(ids) <= {str(i) for i in range(50)})
        with self.assertRaises(ValueError):
            choose_rows(rows, "정회원", 1, allowed_ids=[])
        rows[0]["distributionYn"] = "Y"
        with self.assertRaisesRegex(ValueError, "부족"):
            choose_rows(rows, "정회원", 50, allowed_ids=[str(i) for i in range(50)])

    def test_session_mismatch_blocks_operations_without_logging_out(self):
        page = MagicMock()
        page.url = "http://nmis.foodservice.or.kr/"
        page.locator.return_value.count.return_value = 0
        page.evaluate.return_value = "another-account"
        with self.assertRaisesRegex(RuntimeError, "로그인 계정이 변경"):
            require_session_account(page, "active-account", "password")
        page.locator.return_value.click.assert_not_called()
        with self.assertRaisesRegex(RuntimeError, "체크"):
            require_session_account(page, "", "")

    def test_explicit_switch_uses_saved_credentials_and_verifies_identity(self):
        page = MagicMock()
        page.url = "http://nmis.foodservice.or.kr/"
        page.locator.return_value.count.side_effect = [0, 1]
        page.evaluate.side_effect = ["old", "new"]
        with patch("membership_fee_register.ensure_nmis_logged_in") as login:
            require_session_account(page, "new", "password", switch=True)
        login.assert_called_once()
        self.assertEqual(login.call_args.args[1:3], ("new", "password"))

    def test_login_identity_ignores_case_and_padding(self):
        page = MagicMock()
        page.url = "http://nmis.foodservice.or.kr/"
        page.locator.return_value.count.return_value = 0
        page.evaluate.return_value = "  DEMO-ACCOUNT  "
        require_session_account(page, "demo-account", "password", switch=True)
        page.locator.return_value.click.assert_not_called()

    def test_slow_initial_login_form_is_ready_before_credentials_are_checked(self):
        page = MagicMock()
        page.url = "http://nmis.foodservice.or.kr/"
        page.locator.return_value.count.return_value = 0
        page.evaluate.return_value = "selected"

        def ready(*args, **kwargs):
            page.locator.return_value.count.return_value = 1

        page.wait_for_function.side_effect = ready
        with patch("membership_fee_register.ensure_nmis_logged_in") as login:
            require_session_account(page, "selected", "password", switch=True)
        login.assert_called_once()
        self.assertEqual(page.wait_for_function.call_count, 2)
        page.locator.return_value.click.assert_not_called()

    def test_loading_timeout_does_not_login_or_logout(self):
        from playwright.sync_api import TimeoutError as PlaywrightTimeoutError
        page = MagicMock()
        page.url = "http://nmis.foodservice.or.kr/"
        page.wait_for_function.side_effect = PlaywrightTimeoutError("still loading")
        with patch("membership_fee_register.ensure_nmis_logged_in") as login:
            with self.assertRaisesRegex(RuntimeError, "30초"):
                require_session_account(page, "selected", "password", switch=True)
        login.assert_not_called()
        page.locator.assert_not_called()

    def test_other_site_is_rejected_before_waiting_or_using_credentials(self):
        page = MagicMock()
        page.url = "https://example.com/"
        with patch("membership_fee_register.ensure_nmis_logged_in") as login:
            with self.assertRaisesRegex(RuntimeError, "NMIS 주소"):
                require_session_account(page, "selected", "password", switch=True)
        page.wait_for_function.assert_not_called()
        login.assert_not_called()

    def test_navigation_away_during_loading_does_not_use_credentials(self):
        page = MagicMock()
        page.url = "http://nmis.foodservice.or.kr/"
        page.wait_for_function.side_effect = lambda *args, **kwargs: setattr(page, "url", "https://example.com/")
        with patch("membership_fee_register.ensure_nmis_logged_in") as login:
            with self.assertRaisesRegex(RuntimeError, "NMIS 주소"):
                require_session_account(page, "selected", "password", switch=True)
        login.assert_not_called()

    def test_checked_candidates_are_distinct_from_connected_account(self):
        first = self.store.save_account(0, "first", "password")
        second = self.store.save_account(1, "second", "password")
        self.store.set_checked(first["id"], True)
        self.store.set_checked(second["id"], True)
        self.assertIsNone(self.store.active)
        self.assertEqual(len(self.reopen().login_choices), 2)
        self.store.activate(first["id"])
        self.assertEqual(self.store.credentials()[0], "first")
        self.store.save_account(0, "FIRST", "")
        self.assertEqual(self.store.active["id"], first["id"])
        self.store.set_checked(first["id"], False)
        self.assertIsNone(self.store.active)
        self.assertEqual([a["id"] for a in self.reopen().login_choices], [second["id"]])

    def test_legacy_active_account_becomes_checked_candidate(self):
        self.store.migrate("old", "password")
        legacy = json.loads(self.path.read_text())
        legacy.pop("checked_ids")
        self.path.write_text(json.dumps(legacy))
        reopened = self.reopen()
        self.assertEqual([a["id"] for a in reopened.login_choices], [reopened.active["id"]])

    def test_five_group_dates_and_members_restore_across_accounts(self):
        first = self.store.save_account(0, "first", "password")
        second = self.store.save_account(1, "second", "password")
        groups = [{str(i): {"memberName": f"업소{i}"}} for i in range(5)]
        dates = [f"2026-11-{i+1:02d}" for i in range(5)]
        self.store.save_distribution_data(first["id"], groups, dates)
        self.store.save_distribution_data(second["id"], [{"other": {}}] + [{} for _ in range(4)], ["20261201"] * 5)
        reopened = self.reopen()
        restored, restored_dates = reopened.reload_distribution_data(first["id"])
        self.assertEqual(restored, groups)
        self.assertEqual(restored_dates, [f"202611{i+1:02d}" for i in range(5)])
        self.assertEqual(reopened.groups(second["id"])[0], {"other": {}})
        self.assertEqual(reopened.group_dates(second["id"]), ["20261201"] * 5)

    def test_invalid_dates_or_overlap_do_not_partially_save_data(self):
        record = self.store.save_account(0, "test", "password")
        groups = [{"one": {}}] + [{} for _ in range(4)]
        self.store.save_distribution_data(record["id"], groups, ["20261101"] * 5)
        before = self.path.read_bytes()
        with self.assertRaisesRegex(ValueError, "날짜"):
            self.store.save_distribution_data(record["id"], groups, ["20260230"] + [""] * 4)
        with self.assertRaisesRegex(ValueError, "중복"):
            self.store.save_distribution_data(record["id"], [{"one": {}}, {"one": {}}] + [{}] * 3, [""] * 5)
        self.assertEqual(self.path.read_bytes(), before)
        self.assertEqual(self.store.group_dates(record["id"]), ["20261101"] * 5)

    def test_reload_reads_latest_disk_snapshot_and_legacy_groups(self):
        record = self.store.save_account(0, "test", "password")
        self.store.save_group(record["id"], 0, {"old": {}})
        self.assertEqual(self.store.group_dates(record["id"]), [""] * 5)
        other_instance = self.reopen()
        other_instance.save_distribution_data(record["id"], [{"new": {}}] + [{}] * 4, ["20261201"] * 5)
        groups, dates = self.store.reload_distribution_data(record["id"])
        self.assertEqual(groups[0], {"new": {}})
        self.assertEqual(dates, ["20261201"] * 5)

    def test_checked_members_move_to_explicit_target_without_replacing_others(self):
        groups = [{"m1": {"memberName": "첫번째"}}, {"m2": {"memberName": "기존"}}, {"m3": {}}, {}, {}]
        moved = assign_distribution_members(groups, 1, {"m1": groups[0]["m1"], "m4": {}})
        self.assertEqual(set(moved[0]), set())
        self.assertEqual(set(moved[1]), {"m1", "m2", "m4"})
        self.assertEqual(moved[2], groups[2])
        self.assertIn("m1", groups[0])
        self.assertEqual(sum("m1" in group for group in moved), 1)

    def test_saving_to_same_tab_keeps_current_checks(self):
        groups = [{"m1": {}}, {}, {}, {}, {}]
        self.assertEqual(assign_distribution_members(groups, 0, groups[0]), groups)
        with self.assertRaises(ValueError):
            assign_distribution_members(groups, 5, {})
