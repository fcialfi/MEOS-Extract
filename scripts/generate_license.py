#!/usr/bin/env python3
"""Generate a signed license file.

This script is intended for internal use to produce ``license.key`` files.
It reads a secret key from a file, signs the provided license metadata using
HMAC/SHA256, and writes the result to ``license.key``.

The secret key file **must** be stored securely and should not be distributed
with the application or committed to source control.
"""
from __future__ import annotations

import argparse
import hmac
import hashlib
import json
from pathlib import Path
from datetime import datetime

def load_secret(path: Path) -> bytes:
    try:
        return path.read_bytes().strip()
    except OSError as exc:
        raise SystemExit(f"Unable to read secret key: {exc}")

def main() -> None:
    parser = argparse.ArgumentParser(description="Generate a signed license file")
    parser.add_argument("--client", required=True, help="Client name")
    parser.add_argument(
        "--expires", required=True, help="Expiration date in YYYY-MM-DD format"
    )
    parser.add_argument(
        "--secret-file", required=True, help="Path to file containing HMAC secret"
    )
    parser.add_argument(
        "--output", default="license.key", help="Output path for license file"
    )
    args = parser.parse_args()

    # Validate expiration date format
    try:
        datetime.strptime(args.expires, "%Y-%m-%d")
    except ValueError as exc:
        raise SystemExit(f"Invalid --expires value: {exc}")

    secret = load_secret(Path(args.secret_file))

    payload = f"{args.client}|{args.expires}".encode("utf-8")
    signature = hmac.new(secret, payload, hashlib.sha256).hexdigest()

    license_data = {"client": args.client, "expires": args.expires, "signature": signature}
    Path(args.output).write_text(json.dumps(license_data), encoding="utf-8")
    print(f"License written to {args.output}")

if __name__ == "__main__":
    main()
