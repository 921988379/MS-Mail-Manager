import unittest
import importlib.util
from pathlib import Path

spec = importlib.util.spec_from_file_location('app', str(Path(__file__).resolve().parents[1] / 'app.py'))
app = importlib.util.module_from_spec(spec)
spec.loader.exec_module(app)

class OAuthEndpointTests(unittest.TestCase):
    def test_default_tenant_is_consumers_for_personal_accounts(self):
        self.assertEqual(app.normalize_tenant(''), 'consumers')
        self.assertEqual(app.normalize_tenant('consumers'), 'consumers')

    def test_blocks_full_url_as_tenant(self):
        self.assertEqual(app.normalize_tenant('https://evil.example'), 'consumers')

if __name__ == '__main__':
    unittest.main()
