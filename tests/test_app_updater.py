import io
import json
import tempfile
import unittest
import zipfile
from pathlib import Path
from unittest.mock import patch

from app_updater import ASSET_NAME, check_latest, extract_payload, stage_update, version_tuple


class Response(io.BytesIO):
    headers = {}


class AppUpdaterTests(unittest.TestCase):
    def test_numeric_version_comparison(self):
        self.assertGreater(version_tuple("v2026.11.05.10"), version_tuple("2026.11.05.9"))
        with self.assertRaises(ValueError):
            version_tuple("arbitrary-tag")

    def test_latest_release_validation_and_no_update(self):
        base = "https://github.com/shinhuigwan/nmis-integrated-automation/releases/download/v2026.11.05.1/"
        release = dict(tag_name="v2026.11.05.1", body="notes", assets=[dict(name=name, browser_download_url=base + name) for name in (ASSET_NAME, ASSET_NAME + ".sha256")])
        with patch("app_updater.request", return_value=Response(json.dumps(release).encode())):
            self.assertEqual(check_latest("2026.11.05.0")["version"], "2026.11.05.1")
        with patch("app_updater.request", return_value=Response(json.dumps(release).encode())):
            self.assertIsNone(check_latest("2026.11.05.1"))
        release["assets"][0]["browser_download_url"] = "https://other.invalid/update.zip"
        with patch("app_updater.request", return_value=Response(json.dumps(release).encode())):
            with self.assertRaises(ValueError):
                check_latest("2026.11.05.0")

    def test_extract_rejects_paths_and_user_data(self):
        for bad in ("../outside", "/outside", "NMIS-Automation/../outside", "NMIS-Automation/settings.json", "NMIS-Automation/data/accounts.json", "NMIS-Automation/_internal/x:stream"):
            with self.subTest(path=bad), tempfile.TemporaryDirectory() as folder:
                archive = Path(folder) / "test.zip"
                with zipfile.ZipFile(archive, "w") as z:
                    z.writestr(bad, "data")
                with self.assertRaises(ValueError):
                    extract_payload(archive, Path(folder) / "extract")
                self.assertFalse((Path(folder) / "outside").exists())

    def test_valid_release_extracts(self):
        with tempfile.TemporaryDirectory() as folder:
            archive = Path(folder) / "test.zip"
            with zipfile.ZipFile(archive, "w") as z:
                z.writestr("NMIS-Automation/NMIS-Automation.exe", "exe")
                z.writestr("NMIS-Automation/_internal/runtime.dll", "runtime")
            root = extract_payload(archive, Path(folder) / "extract")
            self.assertTrue((root / "NMIS-Automation.exe").is_file())

    def test_bad_checksum_never_runs_downloaded_executable(self):
        with patch("app_updater.request", side_effect=[Response(("0" * 64).encode()), Response(b"corrupt payload")]), patch("app_updater.subprocess.run") as execute:
            with self.assertRaisesRegex(ValueError, "검증"):
                stage_update(dict(checksum_url="checksum", url="download", version="2026.11.05.1"))
        execute.assert_not_called()
