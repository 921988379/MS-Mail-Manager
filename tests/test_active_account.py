import unittest
import importlib.util
from pathlib import Path

spec = importlib.util.spec_from_file_location('app', str(Path(__file__).resolve().parents[1] / 'app.py'))
app = importlib.util.module_from_spec(spec)
spec.loader.exec_module(app)

class ActiveAccountTests(unittest.TestCase):
    def test_make_active_account_cookie(self):
        cookie = app.make_active_account_cookie('12')
        self.assertIn('active_account_id=12', cookie)
        self.assertIn('Path=/', cookie)

    def test_get_active_account_id_from_cookie(self):
        self.assertEqual(app.get_active_account_id('foo=bar; active_account_id=99; x=y'), '99')
        self.assertEqual(app.get_active_account_id(''), '')

if __name__ == '__main__':
    unittest.main()
