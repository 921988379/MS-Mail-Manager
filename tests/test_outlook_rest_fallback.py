import unittest
import importlib.util
from pathlib import Path
from unittest.mock import patch

spec = importlib.util.spec_from_file_location('app', str(Path(__file__).resolve().parents[1] / 'app.py'))
app = importlib.util.module_from_spec(spec)
spec.loader.exec_module(app)


class OutlookRestFallbackTests(unittest.TestCase):
    def test_outlook_rest_message_to_mail(self):
        msg = {
            'ReceivedDateTime': '2026-06-17T10:00:00Z',
            'Subject': 'Your code is 123456',
            'BodyPreview': 'Use 123456 to sign in',
            'From': {'EmailAddress': {'Name': 'Microsoft', 'Address': 'account-security-noreply@accountprotection.microsoft.com'}},
        }
        mail = app.outlook_rest_message_to_mail(msg)
        self.assertEqual(mail['date'], '2026-06-17T10:00:00Z')
        self.assertIn('Microsoft', mail['from'])
        self.assertEqual(mail['subject'], 'Your code is 123456')
        self.assertIn('123456', mail['preview'])

    def test_fetch_codes_uses_outlook_rest_when_graph_fails(self):
        account = {'email': 'a@hotmail.com', 'client_id': 'cid', 'refresh_token': 'rt'}
        rest_mail = [{'from': 'Microsoft', 'subject': 'Code 654321', 'preview': '', 'date': 'now', 'folder': '收件箱'}]
        with patch.object(app, 'exchange_refresh_token', return_value=(True, {'access_token': 'AT'})), \
             patch.object(app, 'fetch_graph_latest_emails', side_effect=Exception('HTTP Error 401: Unauthorized')), \
             patch.object(app, 'fetch_outlook_rest_latest_emails', return_value=rest_mail), \
             patch.object(app, 'fetch_latest_emails') as imap_mock, \
             patch.object(app, 'update_saved_status') as status_mock:
            result = app.fetch_latest_codes_for_account(account)
        self.assertEqual(result['status'], 'ok')
        self.assertEqual(result['source_api'], 'outlook_rest')
        self.assertEqual(result['codes'][0]['code'], '654321')
        imap_mock.assert_not_called()
        status_mock.assert_called_with('a@hotmail.com', 'outlook_rest_ok')

    def test_inspect_reports_outlook_rest_channel(self):
        account = {'id': 1, 'email': 'a@hotmail.com', 'password': '', 'client_id': 'cid', 'refresh_token': 'rt'}
        rest_mail = [{'from': 'Microsoft', 'subject': 'Code 111222', 'preview': '', 'date': 'now', 'folder': '收件箱'}]
        with patch.object(app, 'check_saved_account_token', return_value={'ok': True, 'status': '正常'}), \
             patch.object(app, 'exchange_refresh_token', return_value=(True, {'access_token': 'AT'})), \
             patch.object(app, 'fetch_graph_latest_emails', side_effect=Exception('HTTP Error 401: Unauthorized')), \
             patch.object(app, 'fetch_outlook_rest_latest_emails', return_value=rest_mail), \
             patch.object(app, 'update_saved_status'):
            result = app.inspect_saved_account(account)
        self.assertFalse(result['graph_mail']['ok'])
        self.assertTrue(result['outlook_rest_mail']['ok'])
        self.assertEqual(result['best_source'], 'outlook_rest')
        self.assertEqual(result['best_codes'][0]['code'], '111222')


if __name__ == '__main__':
    unittest.main()
