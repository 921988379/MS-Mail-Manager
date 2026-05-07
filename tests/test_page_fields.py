import unittest
import importlib.util
from pathlib import Path
from unittest.mock import patch

spec = importlib.util.spec_from_file_location('app', str(Path(__file__).resolve().parents[1] / 'app.py'))
app = importlib.util.module_from_spec(spec)
spec.loader.exec_module(app)

class PageFieldTests(unittest.TestCase):
    def test_home_page_only_exposes_required_inputs(self):
        with patch.object(app, 'recent_records', return_value=[]), patch.object(app, 'saved_accounts', return_value=[]):
            html = app.render_home_page()
        self.assertIn('name="client_id"', html)
        self.assertIn('name="email"', html)
        self.assertIn('name="refresh_token"', html)
        self.assertIn('name="batch"', html)
        self.assertNotIn('name="tenant"', html)
        self.assertNotIn('name="scope"', html)

if __name__ == '__main__':
    unittest.main()
