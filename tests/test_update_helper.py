import ctypes
import os
import shutil
import subprocess
import tempfile
import unittest
from pathlib import Path


@unittest.skipUnless(os.name == "nt", "Windows updater helper")
class UpdateHelperTests(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory()
        self.addCleanup(self.temp.cleanup)
        self.root = Path(self.temp.name)
        self.app = self.root / "app with spaces"
        self.stage = self.root / "stage with spaces"
        self.payload = self.stage / "payload" / "NMIS-Automation"
        for root in (self.app, self.payload):
            (root / "_internal").mkdir(parents=True)
        self.stage.mkdir(exist_ok=True)
        self.helper = self.stage / "apply_update.ps1"
        shutil.copy2(Path(__file__).resolve().parents[1] / "packaging/apply_update.ps1", self.helper)
        (self.app / "NMIS-Automation.exe").write_text("old exe")
        (self.app / "_internal/runtime.dll").write_text("old runtime")
        (self.app / "Setup.ps1").write_text("old setup")
        (self.app / "settings.json").write_text('{"keep":"settings"}')
        (self.app / "data").mkdir()
        (self.app / "data/account_profiles.json").write_text('{"keep":"accounts and groups"}')
        (self.payload / "NMIS-Automation.exe").write_text("new exe")
        (self.payload / "_internal/runtime.dll").write_text("new runtime")
        (self.payload / "Setup.ps1").write_text("new setup")

    def execute(self):
        return subprocess.run(["powershell.exe", "-NoProfile", "-ExecutionPolicy", "Bypass", "-File", str(self.helper),
                               "-AppRoot", str(self.app), "-PayloadRoot", str(self.payload), "-NoRestart"],
                              capture_output=True, timeout=45)

    def assert_user_files_preserved(self):
        self.assertEqual((self.app / "settings.json").read_text(), '{"keep":"settings"}')
        self.assertEqual((self.app / "data/account_profiles.json").read_text(), '{"keep":"accounts and groups"}')

    def test_install_preserves_user_data(self):
        result = self.execute()
        self.assertEqual(result.returncode, 0, result.stderr.decode(errors="replace"))
        self.assertEqual((self.app / "NMIS-Automation.exe").read_text(), "new exe")
        self.assertEqual((self.app / "_internal/runtime.dll").read_text(), "new runtime")
        self.assertEqual((self.stage / "backup/NMIS-Automation.exe").read_text(), "old exe")
        self.assert_user_files_preserved()

    def test_locked_payload_rolls_back_already_replaced_exe_and_runtime(self):
        create = ctypes.windll.kernel32.CreateFileW
        create.restype = ctypes.c_void_p
        handle = create(str(self.payload / "Setup.ps1"), 0x80000000, 0, None, 3, 0x80, None)
        self.assertNotEqual(handle, ctypes.c_void_p(-1).value)
        try:
            result = self.execute()
        finally:
            ctypes.windll.kernel32.CloseHandle(ctypes.c_void_p(handle))
        self.assertNotEqual(result.returncode, 0)
        self.assertEqual((self.app / "NMIS-Automation.exe").read_text(), "old exe")
        self.assertEqual((self.app / "_internal/runtime.dll").read_text(), "old runtime")
        self.assertEqual((self.app / "Setup.ps1").read_text(), "old setup")
        self.assert_user_files_preserved()
