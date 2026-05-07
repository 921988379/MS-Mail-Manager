import unittest
import importlib.util
from pathlib import Path

spec = importlib.util.spec_from_file_location('app', str(Path(__file__).resolve().parents[1] / 'app.py'))
app = importlib.util.module_from_spec(spec)
spec.loader.exec_module(app)

class GraphMailHelperTests(unittest.TestCase):
    def test_graph_message_to_mail(self):
        msg = {
            'receivedDateTime': '2026-05-06T10:00:00Z',
            'subject': 'Your code is 123456',
            'bodyPreview': 'Use 123456 to sign in',
            'from': {'emailAddress': {'name': 'Microsoft', 'address': 'account-security-noreply@accountprotection.microsoft.com'}}
        }
        mail = app.graph_message_to_mail(msg)
        self.assertEqual(mail['date'], '2026-05-06T10:00:00Z')
        self.assertIn('Microsoft', mail['from'])
        self.assertEqual(mail['subject'], 'Your code is 123456')
        self.assertIn('123456', mail['preview'])

    def test_imap_not_connected_error_is_detected(self):
        self.assertTrue(app.is_imap_not_connected_error('User is authenticated but not connected.'))
        self.assertFalse(app.is_imap_not_connected_error('bad password'))

if __name__ == '__main__':
    unittest.main()
