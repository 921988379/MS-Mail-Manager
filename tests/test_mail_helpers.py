import unittest
import importlib.util
import base64
from pathlib import Path

spec = importlib.util.spec_from_file_location('app', str(Path(__file__).resolve().parents[1] / 'app.py'))
app = importlib.util.module_from_spec(spec)
spec.loader.exec_module(app)

class MailHelperTests(unittest.TestCase):
    def test_build_xoauth2_payload_matches_official_format(self):
        # imaplib.authenticate() base64-encodes the returned value itself,
        # so our helper must return the raw XOAUTH2 string, not pre-base64 text.
        payload = app.build_xoauth2_payload('user@example.com', 'ACCESS')
        self.assertEqual(payload, 'user=user@example.com\x01auth=Bearer ACCESS\x01\x01')

    def test_html_to_text_preview_strips_tags(self):
        self.assertEqual(app.html_to_text_preview('<p>Hello<br>World</p>', 20), 'Hello World')

if __name__ == '__main__':
    unittest.main()
