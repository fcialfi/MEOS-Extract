from __future__ import annotations

from pathlib import Path
import hmac
import hashlib
import json
from datetime import datetime, date
import sys

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

    When running from a PyInstaller ``--onefile`` bundle, ``license.key`` may be
    packaged inside the executable. In that case ``sys._MEIPASS`` points to the
    temporary extraction directory. This function will look for the license file
    there if it is not found at the provided path.
    """
    lic_path = path
    if not lic_path.exists():
        meipass = getattr(sys, "_MEIPASS", None)
        if meipass:
            alt = Path(meipass) / path.name
            if alt.exists():
                lic_path = alt

    lic = load_license(lic_path)
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


_VALIDATED = False


def ensure_valid_license() -> None:
    """Exit with ``RuntimeError`` if ``license.key`` is missing or invalid."""
    global _VALIDATED
    if _VALIDATED:
        return
    if not validate_license(Path("license.key")):
        raise RuntimeError("Invalid or missing license.")
    _VALIDATED = True
