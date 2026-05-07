import unittest
import importlib.util
from pathlib import Path
from unittest.mock import patch

spec = importlib.util.spec_from_file_location('app', str(Path(__file__).resolve().parents[1] / 'app.py'))
app = importlib.util.module_from_spec(spec)
spec.loader.exec_module(app)

class SinglePageUITests(unittest.TestCase):
    def test_home_page_has_compact_regions_and_no_recent_records(self):
        with patch.object(app, 'recent_records', return_value=[]), patch.object(app, 'saved_accounts', return_value=[]):
            html = app.render_home_page()
        self.assertIn('class="grid"', html)
        self.assertIn('保存账号', html)
        self.assertIn('批量导入', html)
        self.assertIn('已保存账号', html)
        self.assertIn('最新验证码邮件', html)
        self.assertNotIn('最近记录', html)

if __name__ == '__main__':
    unittest.main()
