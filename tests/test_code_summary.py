import unittest
import importlib.util
from pathlib import Path

spec = importlib.util.spec_from_file_location('app', str(Path(__file__).resolve().parents[1] / 'app.py'))
app = importlib.util.module_from_spec(spec)
spec.loader.exec_module(app)

class CodeSummaryTests(unittest.TestCase):
    def test_extracts_code_from_subject(self):
        mail = {'from': 'GitHub <noreply@github.com>', 'subject': '123456 is your GitHub code', 'date': 'today', 'preview': 'ignore'}
        s = app.extract_verification_summary('a@hotmail.com', mail)
        self.assertEqual(s['code'], '123456')
        self.assertIn('GitHub', s['source'])

    def test_extracts_code_from_preview(self):
        mail = {'from': 'Microsoft account team <x@microsoft.com>', 'subject': 'Security code', 'date': 'today', 'preview': 'Use 987654 as your security code.'}
        s = app.extract_verification_summary('a@hotmail.com', mail)
        self.assertEqual(s['code'], '987654')
        self.assertIn('Microsoft', s['source'])

    def test_ignores_non_code_email(self):
        mail = {'from': 'News <news@example.com>', 'subject': 'hello', 'date': 'today', 'preview': 'no code here'}
        self.assertIsNone(app.extract_verification_summary('a@hotmail.com', mail))

if __name__ == '__main__':
    unittest.main()
