"""Template for local_secrets.py (which is gitignored).

To bundle a default Groq key into the build:
  1. Copy this file to  local_secrets.py
  2. Paste your key below (raw 'gsk_...' or a base64 blob)

If local_secrets.py is absent, the app falls back to the GROQ_API_KEY
environment variable, then to per-user keys entered in the app's Settings.
"""

GROQ_API_KEY = ""
