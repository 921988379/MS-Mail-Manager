import unittest
import importlib.util
from pathlib import Path

spec = importlib.util.spec_from_file_location('app', str(Path(__file__).resolve().parents[1] / 'app.py'))
app = importlib.util.module_from_spec(spec)
spec.loader.exec_module(app)

class SavedAccountParsingTests(unittest.TestCase):
    def test_parse_batch_line_pipe(self):
        rows = app.parse_batch_accounts('a@hotmail.com|cid|rtoken')
        self.assertEqual(rows, [{'email': 'a@hotmail.com', 'password': '', 'client_id': 'cid', 'refresh_token': 'rtoken'}])

    def test_parse_batch_line_comma(self):
        rows = app.parse_batch_accounts('a@hotmail.com,cid,rtoken')
        self.assertEqual(rows[0]['email'], 'a@hotmail.com')
        self.assertEqual(rows[0]['client_id'], 'cid')
        self.assertEqual(rows[0]['refresh_token'], 'rtoken')

    def test_parse_batch_ignores_bad_lines(self):
        rows = app.parse_batch_accounts('bad\na@hotmail.com----cid----rtoken')
        self.assertEqual(len(rows), 1)
        self.assertEqual(rows[0]['client_id'], 'cid')

if __name__ == '__main__':
    unittest.main()
