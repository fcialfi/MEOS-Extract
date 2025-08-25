from __future__ import annotations

from pathlib import Path
import hmac
import hashlib
import json
from datetime import datetime, date

# Secret key used to sign license files. Replace with your own secret in production.
SECRET_KEY = b"demo-secret"

def load_license(path: Path) -> dict[str, str] | None:
    """Return the license information from *path* or ``None`` if not found."""
    try:
        return json.loads(path.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError):
        return None

def validate_license(path: Path) -> bool:
    """Validate the license file using HMAC/SHA256.

    The license file must contain ``client``, ``expires`` and ``signature`` fields.
    The signature is expected to be the hex digest of
    ``HMAC(SECRET_KEY, f"{client}|{expires}")``.
    """
    lic = load_license(path)
    if not lic:
        return False

    client = lic.get("client")
    expires = lic.get("expires")
    signature = lic.get("signature")
    if not (client and expires and signature):
        return False

    payload = f"{client}|{expires}".encode("utf-8")
    expected = hmac.new(SECRET_KEY, payload, hashlib.sha256).hexdigest()
    if not hmac.compare_digest(signature, expected):
        return False

    try:
        exp_date = datetime.strptime(expires, "%Y-%m-%d").date()
    except ValueError:
        return False

    return date.today() <= exp_date
