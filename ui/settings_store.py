"""Securely store/retrieve the user's Groq settings via Windows Credential Manager.

Key resolution for the engine:
    user key (keyring)  ->  bundled default key (config)  ->  none (offline).
No config files are ever touched by the user.
"""
from __future__ import annotations

import sys
import os

# Allow running both as part of the package and frozen exe.
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

import config  # noqa: E402

try:
    import keyring  # noqa: E402
    _KEYRING_OK = True
except Exception:                       # pragma: no cover
    _KEYRING_OK = False


def get_user_key() -> str:
    if not _KEYRING_OK:
        return ""
    try:
        return keyring.get_password(config.KEYRING_SERVICE, config.KEYRING_USER_KEY) or ""
    except Exception:
        return ""


def set_user_key(key: str) -> None:
    if not _KEYRING_OK:
        return
    try:
        if key:
            keyring.set_password(config.KEYRING_SERVICE, config.KEYRING_USER_KEY, key)
        else:
            keyring.delete_password(config.KEYRING_SERVICE, config.KEYRING_USER_KEY)
    except Exception:
        pass


def get_model() -> str:
    if _KEYRING_OK:
        try:
            m = keyring.get_password(config.KEYRING_SERVICE, config.KEYRING_USER_MODEL)
            if m:
                return m
        except Exception:
            pass
    return config.DEFAULT_GROQ_MODEL


def set_model(model: str) -> None:
    if not _KEYRING_OK:
        return
    try:
        keyring.set_password(config.KEYRING_SERVICE, config.KEYRING_USER_MODEL, model)
    except Exception:
        pass


def effective_key() -> str:
    """The key the engine should use: user key first, else bundled default."""
    return get_user_key() or config.bundled_groq_key()
