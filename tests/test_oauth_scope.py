import unittest
import importlib.util
from pathlib import Path

spec = importlib.util.spec_from_file_location('app', str(Path(__file__).resolve().parents[1] / 'app.py'))
app = importlib.util.module_from_spec(spec)
spec.loader.exec_module(app)

class OAuthScopeTests(unittest.TestCase):
    def test_default_scope_uses_consumer_outlook_scope(self):
        scope = app.default_oauth_scope()
        self.assertIn('offline_access', scope.split())
        self.assertIn('https://outlook.office.com/IMAP.AccessAsUser.All', scope.split())
        self.assertIn('https://outlook.office.com/SMTP.Send', scope.split())
        self.assertNotIn('https://graph.microsoft.com/User.Read', scope.split())
        self.assertNotIn('openid', scope.split())
        self.assertNotIn('profile', scope.split())
        self.assertNotIn('email', scope.split())

if __name__ == '__main__':
    unittest.main()
