from __future__ import annotations

from pathlib import Path
import hmac
import hashlib

# Secret key used to sign license files. Replace with your own secret in production.
SECRET_KEY = b"demo-secret"

def load_license(path: Path) -> str | None:
    """Return the license string from *path* or ``None`` if not found."""
    try:
        return path.read_text(encoding="utf-8").strip()
    except OSError:
        return None

def validate_license(path: Path) -> bool:
    """Validate the license file using HMAC/SHA256.

    The license must contain the hex digest of ``HMAC(SECRET_KEY, b"MEOS-Extract")``.
    """
    key = load_license(path)
    if not key:
        return False
    expected = hmac.new(SECRET_KEY, b"MEOS-Extract", hashlib.sha256).hexdigest()
    return hmac.compare_digest(key, expected)
