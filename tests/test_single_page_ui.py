import unittest
import importlib.util
from pathlib import Path
from unittest.mock import patch

spec = importlib.util.spec_from_file_location('app', str(Path(__file__).resolve().parents[1] / 'app.py'))
app = importlib.util.module_from_spec(spec)
spec.loader.exec_module(app)

class ToolboxUITests(unittest.TestCase):
    def test_home_page_is_toolbox_dashboard(self):
        with patch.object(app, 'dashboard_summary', return_value={
            'total_accounts': 0,
            'ok_accounts': 0,
            'error_accounts': 0,
            'categories': 0,
            'api_keys': 0,
            'api_calls_today': 0,
            'api_fail_today': 0,
            'recent': [],
        }):
            html = app.render_home_page()
        self.assertIn('控制台总览', html)
        self.assertIn('邮箱管理', html)
        self.assertIn('API 密钥', html)
        self.assertIn('项目管理', html)
        self.assertIn('curl -H "X-API-Key:', html)
        self.assertNotIn('最近记录', html)

if __name__ == '__main__':
    unittest.main()
