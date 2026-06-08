import unittest
import importlib.util
from pathlib import Path
from unittest.mock import patch

spec = importlib.util.spec_from_file_location('app', str(Path(__file__).resolve().parents[1] / 'app.py'))
app = importlib.util.module_from_spec(spec)
spec.loader.exec_module(app)

class RefreshPageTests(unittest.TestCase):
    def test_tokens_page_has_standalone_refresh_tool(self):
        with patch.object(app, 'saved_accounts', return_value=[]):
            html = app.render_tokens_page()
        self.assertIn('刷新令牌工具', html)
        self.assertIn('action="/token_tool"', html)
        self.assertIn('name="tenant"', html)
        self.assertIn('name="scope"', html)
        self.assertIn('save_account', html)
        self.assertIn('默认不保存', html)

    def test_token_tool_route_is_registered(self):
        source = Path(__file__).resolve().parents[1].joinpath('app.py').read_text()
        self.assertIn("post_path == '/token_tool'", source)
        self.assertIn('exchange_refresh_token(client_id, refresh_token, scope=scope, tenant=tenant)', source)

    def test_sidebar_tokens_link_is_not_modal_intercepted(self):
        html = app.app_page('令牌管理', 'tokens', '<p>body</p>')
        self.assertIn('href="/tokens"', html)
        self.assertIn('a[href^="/token?"]', html)
        self.assertNotIn('a[href^="/token"]', html)

if __name__ == '__main__':
    unittest.main()
