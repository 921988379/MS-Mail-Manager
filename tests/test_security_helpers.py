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

    def test_encrypt_decrypt_without_data_key_is_plain_compatible(self):
        raw = 'secret-refresh-token'
        old = app.DATA_KEY
        try:
            app.DATA_KEY = ''
            self.assertEqual(app.encrypt_secret_value(raw), raw)
            self.assertEqual(app.decrypt_secret_value(raw), raw)
        finally:
            app.DATA_KEY = old

    def test_encrypt_decrypt_with_data_key_roundtrips(self):
        raw = 'secret-refresh-token'
        old = app.DATA_KEY
        try:
            app.DATA_KEY = 'unit-test-data-key'
            encrypted = app.encrypt_secret_value(raw)
            self.assertTrue(encrypted.startswith(app.ENCRYPTED_PREFIX))
            self.assertNotIn(raw, encrypted)
            self.assertEqual(app.decrypt_secret_value(encrypted), raw)
        finally:
            app.DATA_KEY = old

    def test_encrypted_marker_without_key_decrypts_to_empty(self):
        old = app.DATA_KEY
        try:
            app.DATA_KEY = ''
            self.assertEqual(app.decrypt_secret_value('enc:v1:not-a-real-token'), '')
        finally:
            app.DATA_KEY = old

if __name__ == '__main__':
    unittest.main()
