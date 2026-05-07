import unittest
import importlib.util
from pathlib import Path

spec = importlib.util.spec_from_file_location('app', str(Path(__file__).resolve().parents[1] / 'app.py'))
app = importlib.util.module_from_spec(spec)
spec.loader.exec_module(app)

class SpamFolderTests(unittest.TestCase):
    def test_mail_folders_include_inbox_and_junk(self):
        self.assertIn(('INBOX', '收件箱'), app.MAIL_FOLDERS)
        self.assertIn(('Junk', '垃圾邮箱'), app.MAIL_FOLDERS)

    def test_summary_keeps_folder_label(self):
        mail = {'from':'OpenAI <noreply@tm.openai.com>','subject':'验证码 123456','preview':'','date':'now','folder':'垃圾邮箱'}
        s = app.extract_verification_summary('a@hotmail.com', mail)
        self.assertEqual(s['folder'], '垃圾邮箱')

if __name__ == '__main__':
    unittest.main()
