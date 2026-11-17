import socket
import unittest
from unittest.mock import Mock, patch
from urllib.error import HTTPError, URLError

from network_fetch import is_dns_error, open_with_dns_retry


class NetworkFetchTests(unittest.TestCase):
    def test_retries_dns_failure_then_returns_response(self):
        response = object()
        opener = Mock(side_effect=[
            URLError(socket.gaierror(socket.EAI_AGAIN, "temporary lookup failure")),
            response,
        ])
        with patch("network_fetch.time.sleep") as sleep:
            result = open_with_dns_retry(opener, "request", timeout=10)
        self.assertIs(result, response)
        self.assertEqual(opener.call_count, 2)
        sleep.assert_called_once_with(1)

    def test_persistent_dns_failure_is_bounded_and_classified(self):
        error = URLError(socket.gaierror(socket.EAI_NONAME, "no address"))
        opener = Mock(side_effect=error)
        with patch("network_fetch.time.sleep") as sleep:
            with self.assertRaises(URLError):
                open_with_dns_retry(opener, "request", timeout=10)
        self.assertEqual(opener.call_count, 3)
        self.assertEqual([call.args[0] for call in sleep.call_args_list], [1, 2])
        self.assertTrue(is_dns_error(error))

    def test_http_errors_are_not_retried(self):
        error = HTTPError("https://example.test/", 403, "Forbidden", {}, None)
        opener = Mock(side_effect=error)
        with patch("network_fetch.time.sleep") as sleep:
            with self.assertRaises(HTTPError):
                open_with_dns_retry(opener, "request", timeout=10)
        opener.assert_called_once()
        sleep.assert_not_called()
        self.assertFalse(is_dns_error(error))


if __name__ == "__main__":
    unittest.main()
