import unittest
from unittest.mock import patch
import importlib.util
from pathlib import Path

spec = importlib.util.spec_from_file_location('app', str(Path(__file__).resolve().parents[1] / 'app.py'))
app = importlib.util.module_from_spec(spec)
spec.loader.exec_module(app)

class OAuthScopeTests(unittest.TestCase):
    def test_default_scope_uses_graph_mail_read_scope(self):
        scope = app.default_oauth_scope()
        self.assertIn('offline_access', scope.split())
        self.assertIn('https://graph.microsoft.com/Mail.Read', scope.split())
        self.assertNotIn('https://graph.microsoft.com/User.Read', scope.split())
        self.assertNotIn('openid', scope.split())
        self.assertNotIn('profile', scope.split())
        self.assertNotIn('email', scope.split())


    def test_compatible_exchange_retries_without_scope_on_aadsts70000(self):
        calls = []
        def fake_exchange(client_id, refresh_token, scope='', tenant='consumers'):
            calls.append(scope)
            if scope:
                return False, {'error': 'invalid_grant', 'error_codes': [70000], 'error_description': 'AADSTS70000 scope unauthorized'}
            return True, {'access_token': 'at', 'refresh_token': 'rt2', 'scope': 'https://graph.microsoft.com/Mail.ReadWrite'}
        with patch.object(app, 'exchange_refresh_token', fake_exchange):
            ok, payload = app.exchange_refresh_token_compatible('cid', 'rt', scope=app.default_oauth_scope(), tenant='consumers')
        self.assertTrue(ok)
        self.assertEqual(calls, [app.default_oauth_scope(), ''])
        self.assertEqual(payload['_scope_mode'], 'original_scope_fallback')

if __name__ == '__main__':
    unittest.main()
