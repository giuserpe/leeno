import os
import tempfile
import unittest
from unittest.mock import patch
from tools.version.upload_webdav import upload_file


class TestUploadWebDAV(unittest.TestCase):
    def setUp(self):
        self.tmp_file = tempfile.NamedTemporaryFile(delete=False)
        self.tmp_file.write(b"test content")
        self.tmp_file.close()

    def tearDown(self):
        if os.path.exists(self.tmp_file.name):
            os.remove(self.tmp_file.name)

    @patch("tools.version.upload_webdav._make_request")
    def test_upload_success(self, mock_make_request):
        mock_make_request.return_value = (201, "Created")

        res = upload_file(
            self.tmp_file.name,
            "https://example.com/LeenoNigthlyBuilds",
            "user",
            "pass",
            max_retries=3,
            delays=[0, 0],
        )

        self.assertTrue(res)
        self.assertEqual(mock_make_request.call_count, 1)

    @patch("tools.version.upload_webdav._make_request")
    def test_upload_locked_then_success(self, mock_make_request):
        locked_resp = (
            423,
            "<s:exception>OCA\\DAV\\Connector\\Sabre\\Exception\\FileLocked</s:exception>",
        )
        del_resp = (204, "No Content")
        success_resp = (200, "OK")

        mock_make_request.side_effect = [locked_resp, del_resp, success_resp]

        res = upload_file(
            self.tmp_file.name,
            "https://example.com/LeenoNigthlyBuilds",
            "user",
            "pass",
            max_retries=3,
            delays=[0, 0],
        )

        self.assertTrue(res)
        self.assertEqual(mock_make_request.call_count, 3)

    @patch("tools.version.upload_webdav._make_request")
    def test_upload_unretryable_error(self, mock_make_request):
        mock_make_request.return_value = (401, "Unauthorized")

        res = upload_file(
            self.tmp_file.name,
            "https://example.com/LeenoNigthlyBuilds",
            "user",
            "pass",
            max_retries=3,
            delays=[0, 0],
        )

        self.assertFalse(res)
        self.assertEqual(mock_make_request.call_count, 1)

    @patch("tools.version.upload_webdav._make_request")
    def test_upload_exhaust_retries(self, mock_make_request):
        locked_resp = (
            423,
            "<s:exception>OCA\\DAV\\Connector\\Sabre\\Exception\\FileLocked</s:exception>",
        )
        del_resp = (423, "Locked")

        # Each retry does a PUT attempt and a DELETE attempt
        mock_make_request.side_effect = [
            locked_resp,
            del_resp,
            locked_resp,
            del_resp,
            locked_resp,
        ]

        res = upload_file(
            self.tmp_file.name,
            "https://example.com/LeenoNigthlyBuilds",
            "user",
            "pass",
            max_retries=3,
            delays=[0, 0],
        )

        self.assertFalse(res)

    def test_upload_missing_file(self):
        res = upload_file(
            "/non/existent/file.txt",
            "https://example.com/LeenoNigthlyBuilds",
            "user",
            "pass",
            max_retries=1,
            delays=[0],
        )
        self.assertFalse(res)


if __name__ == "__main__":
    unittest.main()
