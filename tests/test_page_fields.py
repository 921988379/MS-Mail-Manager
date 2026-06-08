import unittest
import importlib.util
from pathlib import Path
from unittest.mock import patch

spec = importlib.util.spec_from_file_location('app', str(Path(__file__).resolve().parents[1] / 'app.py'))
app = importlib.util.module_from_spec(spec)
spec.loader.exec_module(app)

class PageFieldTests(unittest.TestCase):
    def test_mailboxes_page_exposes_import_fields(self):
        with patch.object(app, 'saved_accounts', return_value=[]), patch.object(app, 'saved_categories', return_value=[]):
            html = app.render_mailboxes_page()
        self.assertIn('name="client_id"', html)
        self.assertIn('name="email"', html)
        self.assertIn('name="refresh_token"', html)
        self.assertIn('name="batch"', html)
        self.assertNotIn('name="tenant"', html)
        self.assertNotIn('name="scope"', html)

if __name__ == '__main__':
    unittest.main()
