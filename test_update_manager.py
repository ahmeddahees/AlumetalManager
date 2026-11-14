import hashlib
import io
import json
import unittest
from urllib.error import HTTPError
from unittest.mock import patch

import update_manager


class UpdateManagerTests(unittest.TestCase):
    def test_missing_public_release_is_reported_as_no_release(self):
        error = HTTPError("https://api.github.com", 404, "Not Found", {}, None)
        with patch("update_manager.urlopen", side_effect=error):
            with self.assertRaises(update_manager.NoReleaseFound):
                update_manager.get_latest_release("1.0.0")

    def test_latest_release_returns_only_newer_versions(self):
        response = {
            "tag_name": "v1.2.0",
            "html_url": "https://github.com/ahmeddahees/AlumetalManager/releases/tag/v1.2.0",
            "assets": [
                {
                    "name": "AlumetalManager.exe",
                    "browser_download_url": "https://github.com/ahmeddahees/AlumetalManager/releases/download/v1.2.0/AlumetalManager.exe",
                    "size": 20,
                },
                {
                    "name": "SHA256SUMS.txt",
                    "browser_download_url": "https://github.com/ahmeddahees/AlumetalManager/releases/download/v1.2.0/SHA256SUMS.txt",
                    "size": 90,
                },
            ],
        }
        with patch(
            "update_manager.urlopen",
            return_value=io.BytesIO(json.dumps(response).encode("utf-8")),
        ):
            release = update_manager.get_latest_release("1.1.0")
        self.assertEqual(release.version, "1.2.0")
        self.assertEqual(release.executable_size, 20)

        with patch(
            "update_manager.urlopen",
            return_value=io.BytesIO(json.dumps(response).encode("utf-8")),
        ):
            self.assertIsNone(update_manager.get_latest_release("1.2.0"))

    def test_latest_release_rejects_download_urls_outside_github(self):
        response = {
            "tag_name": "v1.2.0",
            "html_url": "https://github.com/ahmeddahees/AlumetalManager/releases/tag/v1.2.0",
            "assets": [
                {
                    "name": "AlumetalManager.exe",
                    "browser_download_url": "https://example.com/AlumetalManager.exe",
                    "size": 20,
                }
            ],
        }
        with patch(
            "update_manager.urlopen",
            return_value=io.BytesIO(json.dumps(response).encode("utf-8")),
        ):
            with self.assertRaises(update_manager.UpdateError):
                update_manager.get_latest_release("1.1.0")

    def test_download_checks_size_and_sha256_before_returning_file(self):
        executable = b"test executable bytes"
        checksum = hashlib.sha256(executable).hexdigest()
        release = update_manager.Release(
            version="1.2.0",
            page_url="https://github.com/ahmeddahees/AlumetalManager/releases/tag/v1.2.0",
            executable_url="https://github.com/ahmeddahees/AlumetalManager/releases/download/v1.2.0/AlumetalManager.exe",
            checksum_url="https://github.com/ahmeddahees/AlumetalManager/releases/download/v1.2.0/SHA256SUMS.txt",
            executable_size=len(executable),
        )
        with patch(
            "update_manager.urlopen",
            side_effect=[
                io.BytesIO(f"{checksum}  AlumetalManager.exe\n".encode("utf-8")),
                io.BytesIO(executable),
            ],
        ):
            downloaded = update_manager.download_release(release)
        try:
            self.assertEqual(downloaded.read_bytes(), executable)
        finally:
            downloaded.unlink(missing_ok=True)

    def test_download_rejects_a_checksum_mismatch(self):
        executable = b"untrusted bytes"
        release = update_manager.Release(
            version="1.2.0",
            page_url="https://github.com/ahmeddahees/AlumetalManager/releases/tag/v1.2.0",
            executable_url="https://github.com/ahmeddahees/AlumetalManager/releases/download/v1.2.0/AlumetalManager.exe",
            checksum_url="https://github.com/ahmeddahees/AlumetalManager/releases/download/v1.2.0/SHA256SUMS.txt",
            executable_size=len(executable),
        )
        with patch(
            "update_manager.urlopen",
            side_effect=[
                io.BytesIO((b"0" * 64) + b"  AlumetalManager.exe\n"),
                io.BytesIO(executable),
            ],
        ):
            with self.assertRaises(update_manager.UpdateError):
                update_manager.download_release(release)


if __name__ == "__main__":
    unittest.main()
