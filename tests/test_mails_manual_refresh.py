import unittest
import importlib.util
from pathlib import Path
from unittest.mock import patch

spec = importlib.util.spec_from_file_location('app', str(Path(__file__).resolve().parents[1] / 'app.py'))
app = importlib.util.module_from_spec(spec)
spec.loader.exec_module(app)


class MailsManualRefreshTests(unittest.TestCase):
    def test_mails_page_does_not_auto_poll_codes(self):
        rows = [{'id': 1, 'email': 'a@hotmail.com', 'category': '项目A', 'last_status': 'outlook_rest_ok'}]
        with patch.object(app, 'saved_accounts', return_value=rows):
            html = app.render_mails_page('1')
        self.assertIn('手动获取邮件', html)
        self.assertIn('等待手动获取', html)
        self.assertIn('不会自动刷新', html)
        self.assertNotIn('setInterval(pollCodes', html)
        self.assertNotIn('pollCodes(); setInterval', html)
        self.assertIn('outlook_rest', html)


if __name__ == '__main__':
    unittest.main()
