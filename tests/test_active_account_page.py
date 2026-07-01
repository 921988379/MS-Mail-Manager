import unittest
import importlib.util
from pathlib import Path
from unittest.mock import patch

spec = importlib.util.spec_from_file_location('app', str(Path(__file__).resolve().parents[1] / 'app.py'))
app = importlib.util.module_from_spec(spec)
spec.loader.exec_module(app)

class ActiveAccountPageTests(unittest.TestCase):
    def test_mails_page_marks_active_email_and_manual_fetch(self):
        rows = [{'id': 7, 'email': 'a@hotmail.com', 'password_mask': 'p…3', 'client_id': 'cid', 'token_mask': 'tok…mask', 'last_status': 'ok', 'last_error': ''}]
        with patch.object(app, 'saved_accounts', return_value=rows):
            html = app.render_mails_page('7')
        self.assertIn('当前邮箱：<b id="active-email">a@hotmail.com</b>', html)
        self.assertIn('/select?id=7&next=/mails', html)
        self.assertNotIn('/mails-workbench', html)
        self.assertNotIn('高级工作台', html)
        self.assertNotIn('setInterval(pollCodes, 30000)', html)
        self.assertIn('等待手动获取', html)
        self.assertIn('邮箱位置', html)
        self.assertIn('读取邮件', html)
        self.assertIn('最新邮件摘要', html)

if __name__ == '__main__':
    unittest.main()
