import unittest
import importlib.util
from pathlib import Path

spec = importlib.util.spec_from_file_location('app', str(Path(__file__).resolve().parents[1] / 'app.py'))
app = importlib.util.module_from_spec(spec)
spec.loader.exec_module(app)

class BatchFourPartFormatTests(unittest.TestCase):
    def test_parse_email_password_client_refresh_dash_format(self):
        text = 'user@outlook.com----pass123----client-id----refresh-token'
        rows = app.parse_batch_accounts(text)
        self.assertEqual(rows, [{
            'email': 'user@outlook.com',
            'password': 'pass123',
            'client_id': 'client-id',
            'refresh_token': 'refresh-token',
        }])

    def test_parse_old_three_part_format_still_works(self):
        rows = app.parse_batch_accounts('user@outlook.com----client-id----refresh-token')
        self.assertEqual(rows[0]['email'], 'user@outlook.com')
        self.assertEqual(rows[0]['password'], '')
        self.assertEqual(rows[0]['client_id'], 'client-id')
        self.assertEqual(rows[0]['refresh_token'], 'refresh-token')

if __name__ == '__main__':
    unittest.main()
