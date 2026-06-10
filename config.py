"""Application-wide constants and the bundled default Groq key.

NOTE on the bundled key: it is stored lightly obfuscated (base64) as deterrence
only -- this is NOT real secrecy. Anyone can extract it from the .exe. Users are
encouraged to supply their own free Groq key via the in-app Settings dialog,
which is stored securely in the Windows Credential Manager.
"""
from __future__ import annotations

import base64

APP_NAME = "Excel Intelligence Agent"
ORG_NAME = "SAHEL GENERAL HOSPITAL"   # shown in the app window header
APP_VERSION = "1.7.0"          # <-- bump this for every release (1.0.1, 1.1.0, ...)
EXE_NAME = "ExcelIntelligenceAgent"

# BUILD_DATE is stamped automatically at build time by tools/stamp_build.py.
# In a dev run (python main.py) the generated file may be absent -> fallback.
try:
    from buildinfo import BUILD_DATE       # type: ignore
except Exception:
    BUILD_DATE = "development build"

# URL of the JSON update manifest (see README). Leave "" to disable update checks.
# Example manifest:
#   { "version": "1.1.0", "release_date": "2026-07-01",
#     "url": "https://.../ExcelIntelligenceAgent.exe",
#     "notes": "What changed in this release." }
UPDATE_MANIFEST_URL = ""

# Keyring service / username used to store the user-supplied Groq key.
KEYRING_SERVICE = "ExcelIntelligenceAgent"
KEYRING_USER_KEY = "groq_api_key"
KEYRING_USER_MODEL = "groq_model"

# Groq (OpenAI-compatible) endpoint.
GROQ_BASE_URL = "https://api.groq.com/openai/v1/chat/completions"
DEFAULT_GROQ_MODEL = "llama-3.3-70b-versatile"
AVAILABLE_GROQ_MODELS = [
    "llama-3.3-70b-versatile",
    "llama-3.1-8b-instant",
]

# Names of the sheets the engine appends. The originals are never touched.
OUTPUT_SHEETS = {
    "dashboard": "Dashboard",
    "pivot": "Pivot Analysis",
    "kpi": "KPI Analysis",
    "summary": "Executive Summary",
}

# Bundled default Groq key. The key is NOT stored in this file (so it can be
# committed to a public repo). It is read from, in order:
#   1. local_secrets.py  (GROQ_API_KEY)  -- gitignored; bundled into the exe
#   2. the GROQ_API_KEY environment variable
# Each value may be a raw 'gsk_...' key or a base64-obfuscated blob.
# (Per-user keys entered in the app's Settings always take precedence at runtime.)
def _raw_bundled_key() -> str:
    try:
        import local_secrets  # noqa: PLC0415 -- gitignored, optional
        k = (getattr(local_secrets, "GROQ_API_KEY", "") or "").strip()
        if k:
            return k
    except Exception:
        pass
    import os  # noqa: PLC0415
    return (os.environ.get("GROQ_API_KEY", "") or "").strip()


def bundled_groq_key() -> str:
    """Return the bundled key. Accepts a raw 'gsk_' key or a base64 blob."""
    raw = _raw_bundled_key()
    if not raw:
        return ""
    if raw.startswith("gsk_"):
        return raw
    try:
        return base64.b64decode(raw.encode()).decode()
    except Exception:
        return ""
