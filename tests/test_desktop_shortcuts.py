import tempfile
import unittest
from pathlib import Path
from unittest.mock import MagicMock

from desktop_shortcuts import ensure_desktop_shortcuts


class DesktopShortcutTests(unittest.TestCase):
    def test_creates_links_then_leaves_user_removed_icons_alone(self):
        with tempfile.TemporaryDirectory() as folder:
            root = Path(folder)
            (root / "NMIS-Automation.exe").touch()
            shell = MagicMock()
            self.assertTrue(ensure_desktop_shortcuts(root, desktop=root, shell=shell))
            self.assertEqual(shell.CreateShortcut.call_count, 2)
            self.assertEqual(shell.CreateShortcut.return_value.TargetPath, str(root / "NMIS-Automation.exe"))
            shell.reset_mock()
            self.assertFalse(ensure_desktop_shortcuts(root, desktop=root, shell=shell))
            shell.CreateShortcut.assert_not_called()
            self.assertTrue(ensure_desktop_shortcuts(root, desktop=root, shell=shell, force=True))

    def test_moved_folder_retargets_shortcuts_without_touching_profiles(self):
        with tempfile.TemporaryDirectory() as folder:
            root = Path(folder) / "original"
            root.mkdir()
            (root / "NMIS-Automation.exe").touch()
            ensure_desktop_shortcuts(root, desktop=root, shell=MagicMock())
            (root / "data/account_profiles.json").write_text("keep")
            moved = Path(folder) / "moved"
            root.rename(moved)
            shell = MagicMock()
            self.assertTrue(ensure_desktop_shortcuts(moved, desktop=moved, shell=shell))
            self.assertEqual(shell.CreateShortcut.return_value.TargetPath, str(moved / "NMIS-Automation.exe"))
            self.assertEqual((moved / "data/account_profiles.json").read_text(), "keep")
