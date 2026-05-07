import unittest
import importlib.util
from pathlib import Path

spec = importlib.util.spec_from_file_location('app', str(Path(__file__).resolve().parents[1] / 'app.py'))
app = importlib.util.module_from_spec(spec)
spec.loader.exec_module(app)

class SecurityHelperTests(unittest.TestCase):
    def test_mask_secret_keeps_edges_only(self):
        self.assertEqual(app.mask_secret('abcdef1234567890'), 'abcd…7890')

    def test_mask_short_secret(self):
        self.assertEqual(app.mask_secret('abc'), '***')

    def test_digest_is_stable_and_not_raw(self):
        d1 = app.secret_digest('refresh-token-example')
        d2 = app.secret_digest('refresh-token-example')
        self.assertEqual(d1, d2)
        self.assertNotIn('refresh-token-example', d1)
        self.assertEqual(len(d1), 64)

    def test_constant_time_password(self):
        self.assertTrue(app.check_password('secret', 'secret'))
        self.assertFalse(app.check_password('secret', 'wrong'))

if __name__ == '__main__':
    unittest.main()
