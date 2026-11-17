"""Bounded retries for transient DNS failures in review-only network fetches."""
import socket
import sys
import time
from urllib.error import URLError


def is_dns_error(error):
    reason = error.reason if isinstance(error, URLError) else error
    return isinstance(reason, socket.gaierror)


def open_with_dns_retry(open_function, request, timeout, *, attempts=3):
    if attempts < 1:
        raise ValueError("attempts must be at least one")
    for attempt in range(attempts):
        try:
            return open_function(request, timeout=timeout)
        except (URLError, socket.gaierror) as exc:
            if not is_dns_error(exc) or attempt + 1 == attempts:
                raise
            delay = 2 ** attempt
            print(
                f"DNS lookup failed; retrying in {delay}s "
                f"({attempt + 1}/{attempts - 1})",
                file=sys.stderr,
                flush=True,
            )
            time.sleep(delay)
