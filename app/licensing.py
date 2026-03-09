"""
Hardware-locked licensing system.
Generates a machine fingerprint and validates license keys offline.
"""

import hashlib
import hmac
import json
import os
import platform
import sys
import uuid
from typing import Optional, Tuple

# Embedded in the compiled binary (Nuitka obfuscates this in the .exe)
_LICENSE_SECRET = b"FM_EMBROIDERY_COMBINER_2026_SECRET_KEY"

LICENSE_FILE = "license.dat"


def _get_app_dir() -> str:
    if getattr(sys, 'frozen', False):
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.dirname(os.path.abspath(__file__)))


def get_machine_id() -> str:
    """Generate a unique machine fingerprint from hardware identifiers."""
    mac = uuid.getnode()
    cpu = platform.processor()
    system = platform.system()
    raw = f"{mac}-{cpu}-{system}".encode()
    return hashlib.sha256(raw).hexdigest()[:16].upper()


def generate_license_key(machine_id: str, secret: bytes = _LICENSE_SECRET) -> str:
    """Generate a license key for a given machine ID (developer use only)."""
    sig = hmac.new(secret, machine_id.encode(), hashlib.sha256).hexdigest()
    key = sig[:16].upper()
    return f"{key[:4]}-{key[4:8]}-{key[8:12]}-{key[12:16]}"


def validate_license_key(machine_id: str, license_key: str, secret: bytes = _LICENSE_SECRET) -> bool:
    """Check if a license key is valid for this machine."""
    expected = generate_license_key(machine_id, secret)
    return hmac.compare_digest(expected, license_key.strip().upper())


def get_license_path() -> str:
    return os.path.join(_get_app_dir(), LICENSE_FILE)


def save_license(machine_id: str, license_key: str) -> None:
    data = {"machine_id": machine_id, "license_key": license_key}
    try:
        with open(get_license_path(), 'w') as f:
            json.dump(data, f)
    except OSError:
        pass


def load_license() -> Tuple[Optional[str], Optional[str]]:
    path = get_license_path()
    if not os.path.exists(path):
        return None, None
    try:
        with open(path, 'r') as f:
            data = json.load(f)
        return data.get("machine_id"), data.get("license_key")
    except (json.JSONDecodeError, KeyError, OSError):
        return None, None


def check_license() -> bool:
    """Check if the current machine has a valid license."""
    try:
        current_id = get_machine_id()
        saved_id, saved_key = load_license()
        if saved_id and saved_key:
            if saved_id == current_id and validate_license_key(current_id, saved_key):
                return True
    except Exception:
        pass
    return False


def activate_license(license_key: str) -> bool:
    """Attempt to activate with the given license key."""
    machine_id = get_machine_id()
    if validate_license_key(machine_id, license_key):
        save_license(machine_id, license_key)
        return True
    return False
