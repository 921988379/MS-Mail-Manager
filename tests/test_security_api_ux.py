import unittest
import importlib.util
from pathlib import Path

spec = importlib.util.spec_from_file_location('app', str(Path(__file__).resolve().parents[1] / 'app.py'))
app = importlib.util.module_from_spec(spec)
spec.loader.exec_module(app)

class SecurityAndApiUXTests(unittest.TestCase):
    def test_csrf_cookie_roundtrip_and_tamper_rejected(self):
        cookie = app.make_csrf_cookie()
        raw = cookie.split(';', 1)[0].split('=', 1)[1]
        self.assertTrue(app.verify_csrf_cookie_value(raw))
        self.assertTrue(app.verify_csrf_form('rtweb_csrf=' + raw, {'csrf_token': [raw]}))
        self.assertFalse(app.verify_csrf_form('rtweb_csrf=' + raw, {'csrf_token': [raw + 'x']}))

    def test_summary_includes_preview_for_body_rule(self):
        summary = app.extract_verification_summary('a@hotmail.com', {
            'from': 'Microsoft account team',
            'subject': 'Your code',
            'preview': 'Use 123456 to sign in',
            'folder': '收件箱',
            'date': 'today',
        })
        self.assertIsNotNone(summary)
        self.assertEqual(summary['code'], '123456')
        self.assertIn('Use 123456', summary['preview'])
        self.assertTrue(app.mail_matches_rule({'from': summary['source'], 'subject': summary['subject'], 'preview': summary['preview']}, {'sender_keywords': '', 'subject_keywords': '', 'body_keywords': 'sign in'}))



    def test_help_documents_token_and_imap_tools(self):
        html = app.render_help_page()
        self.assertIn('仓库里已有的工具能力', html)
        self.assertIn('/token_tool', html)
        self.assertIn('/api/v1/account-status', html)
        self.assertIn('密码 IMAP 可读', html)
        self.assertIn('AADSTS70000', html)

    def test_api_manage_documents_account_status_endpoint(self):
        html = app.render_api_manage_page()
        self.assertIn('/api/v1/account-status', html)
        self.assertIn('account-status', html)

    def test_account_status_endpoint_requires_two_scopes(self):
        source = Path(app.__file__).read_text()
        self.assertIn("'/api/v1/account-status': ['accounts', 'latest_code']", source)
        self.assertIn("missing_scopes", source)

    def test_api_manage_documents_header_and_projects_endpoint(self):
        html = app.render_api_manage_page()
        self.assertIn('X-API-Key', html)
        self.assertIn('/api/v1/projects', html)
        self.assertNotIn('/api/v1/health?key=', html)

    def test_api_scope_helpers_and_query_switch_constant(self):
        self.assertEqual(app.normalize_api_scopes('health, latest-code, bad, projects'), 'health,latest_code,projects')
        row = {'scopes': 'health,projects'}
        self.assertTrue(app.api_key_has_scope(row, 'health'))
        self.assertFalse(app.api_key_has_scope(row, 'accounts'))
        self.assertIsInstance(app.API_ALLOW_QUERY_KEY, bool)

    def test_api_key_page_has_scopes_and_query_switch_notice(self):
        html = app.render_api_key_page()
        self.assertIn('name="scopes" value="latest_code"', html)
        self.assertIn('accounts 会暴露邮箱地址元数据', html)
        self.assertIn('Query key 开关', html)

    def test_security_headers_include_csp(self):
        self.assertIn('Content-Security-Policy', Path(__file__).resolve().parents[1].joinpath('app.py').read_text())
        self.assertIn("frame-ancestors 'none'", Path(__file__).resolve().parents[1].joinpath('app.py').read_text())

if __name__ == '__main__':
    unittest.main()
