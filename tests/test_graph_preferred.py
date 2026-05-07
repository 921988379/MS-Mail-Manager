import unittest
import importlib.util
from pathlib import Path
from unittest.mock import patch

spec = importlib.util.spec_from_file_location('app', str(Path(__file__).resolve().parents[1] / 'app.py'))
app = importlib.util.module_from_spec(spec)
spec.loader.exec_module(app)

class GraphPreferredTests(unittest.TestCase):
    def test_fetch_codes_prefers_graph_preview_over_imap(self):
        account = {'email': 'a@hotmail.com', 'client_id': 'cid', 'refresh_token': 'rt'}
        def fake_exchange(client_id, refresh_token, scope='', tenant='consumers'):
            return True, {'access_token': 'AT'}
        graph_mail = [{'from':'OpenAI <noreply@tm1.openai.com>', 'subject':'你的临时 ChatGPT 登录代码', 'preview':'输入此临时验证码以继续：\n\n446095', 'date':'now', 'folder':'收件箱'}]
        imap_mail = [{'from':'OpenAI <noreply@tm1.openai.com>', 'subject':'你的临时 ChatGPT 登录代码', 'preview':'@font-face css no visible code', 'date':'now', 'folder':'收件箱'}]
        with patch.object(app, 'exchange_refresh_token', side_effect=fake_exchange), \
             patch.object(app, 'fetch_graph_latest_emails', return_value=graph_mail), \
             patch.object(app, 'fetch_latest_emails', return_value=imap_mail), \
             patch.object(app, 'update_saved_status'):
            res = app.fetch_latest_codes_for_account(account, limit=5)
        self.assertEqual(res['source_api'], 'graph')
        self.assertEqual(res['codes'][0]['code'], '446095')

if __name__ == '__main__':
    unittest.main()
