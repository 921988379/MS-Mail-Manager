import unittest
import importlib.util
from pathlib import Path
from unittest.mock import patch

spec = importlib.util.spec_from_file_location('app', str(Path(__file__).resolve().parents[1] / 'app.py'))
app = importlib.util.module_from_spec(spec)
spec.loader.exec_module(app)

class AccountInspectTests(unittest.TestCase):
    def test_mailboxes_page_has_inspect_button_and_modal_title(self):
        rows = [{'id': 1, 'email': 'a@hotmail.com', 'password_mask': 'p…s', 'category': '项目A', 'client_id': 'cid', 'token_mask': 'tok…mask', 'last_status': '', 'last_error': ''}]
        with patch.object(app, 'saved_accounts', return_value=rows), patch.object(app, 'saved_categories', return_value=['项目A']):
            html = app.render_mailboxes_page()
        self.assertIn('action="/inspect_account"', html)
        self.assertIn('综合检测', html)
        self.assertIn("'/inspect_account': '账号综合检测'", app.app_page('t', 'mailboxes', '<p>x</p>'))

    def test_inspect_falls_back_to_password_imap_when_token_fails(self):
        account = {'id': 1, 'email': 'a@hotmail.com', 'password': 'secret', 'client_id': 'cid', 'refresh_token': 'rt'}
        with patch.object(app, 'check_saved_account_token', return_value={'ok': False, 'status': '失效/错误', 'error': 'invalid_grant'}), \
             patch.object(app, 'exchange_refresh_token', return_value=(False, {'error': 'invalid_grant'})), \
             patch.object(app, 'fetch_password_latest_emails', return_value=[{'from': 'Microsoft', 'subject': 'Code 123456', 'preview': '', 'date': 'now', 'folder': '收件箱'}]), \
             patch.object(app, 'update_saved_status'):
            result = app.inspect_saved_account(account)
        self.assertFalse(result['token']['ok'])
        self.assertTrue(result['password_login']['ok'])
        self.assertEqual(result['best_source'], 'imap_password')
        self.assertEqual(result['best_codes'][0]['code'], '123456')

    def test_render_inspect_result_does_not_expose_secrets(self):
        result = {
            'email': 'a@hotmail.com',
            'password_login': {'configured': True, 'ok': True, 'error': '', 'mail_count': 1, 'codes': []},
            'token': {'configured': True, 'ok': False, 'status': '失效/错误', 'expires_in': '', 'scope': '', 'rotated': False, 'error': 'invalid_grant'},
            'graph_mail': {'ok': False, 'error': 'invalid_grant', 'mail_count': 0, 'codes': []},
            'best_codes': [{'source_api': 'imap_password', 'folder': '收件箱', 'code': '123456', 'source': 'Microsoft', 'subject': 'Code', 'date': 'now'}],
            'best_source': 'imap_password',
        }
        html = app.render_account_inspect_result(result)
        self.assertIn('账号综合检测', html)
        self.assertIn('账号密码 IMAP', html)
        self.assertIn('123456', html)
        self.assertNotIn('refresh_token', html.lower())
        self.assertNotIn('access_token', html.lower())

    def test_latest_codes_uses_password_imap_when_refresh_token_fails(self):
        account = {'id': 1, 'email': 'a@hotmail.com', 'password': 'secret', 'client_id': 'cid', 'refresh_token': 'rt'}
        with patch.object(app, 'exchange_refresh_token', return_value=(False, {'error': 'invalid_grant'})), \
             patch.object(app, 'fetch_password_latest_emails', return_value=[{'from': 'Microsoft', 'subject': 'Code 654321', 'preview': '', 'date': 'now', 'folder': '收件箱'}]), \
             patch.object(app, 'update_saved_status'):
            result = app.fetch_latest_codes_for_account(account)
        self.assertEqual(result['status'], 'ok')
        self.assertEqual(result['source_api'], 'imap_password')
        self.assertEqual(result['codes'][0]['code'], '654321')

    def test_batch_inspect_button_route_and_status_labels(self):
        rows = [{'id': 1, 'email': 'a@hotmail.com', 'password_mask': 'p…s', 'category': '项目A', 'client_id': 'cid', 'token_mask': 'tok…mask', 'last_status': 'imap_password_ok', 'last_error': ''}]
        with patch.object(app, 'saved_accounts', return_value=rows), patch.object(app, 'saved_categories', return_value=['项目A']):
            html = app.render_mailboxes_page()
        self.assertIn('action="/inspect_selected"', html)
        self.assertIn('密码 IMAP 可读', html)
        self.assertIn("'/inspect_selected': '批量综合检测'", app.app_page('t', 'mailboxes', '<p>x</p>'))

    def test_batch_inspect_result_summary(self):
        result = {
            'email': 'a@hotmail.com',
            'password_login': {'configured': True, 'ok': True, 'error': '', 'mail_count': 1, 'codes': []},
            'token': {'configured': True, 'ok': False, 'status': '失效/错误', 'expires_in': '', 'scope': '', 'rotated': False, 'error': 'invalid_grant'},
            'graph_mail': {'ok': False, 'error': 'invalid_grant', 'mail_count': 0, 'codes': []},
            'best_codes': [{'source_api': 'imap_password', 'folder': '收件箱', 'code': '123456', 'source': 'Microsoft', 'subject': 'Code', 'date': 'now'}],
            'best_source': 'imap_password',
        }
        html = app.render_batch_inspect_result([result])
        self.assertIn('批量综合检测结果', html)
        self.assertIn('密码 IMAP 可读', html)
        self.assertIn('验证码摘要', html)

    def test_status_label_maps_new_statuses(self):
        self.assertEqual(app.status_label('imap_password_ok'), '密码 IMAP 可读')
        self.assertEqual(app.status_label('all_failed'), '全部失败')

if __name__ == '__main__':
    unittest.main()
