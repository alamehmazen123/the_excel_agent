"""Stamp the build date AND the bundled Groq key into buildinfo.py.

Run by build.bat before packaging. buildinfo.py is gitignored (so the key never
reaches a public repo) and is imported normally by config.py, which guarantees
PyInstaller bundles it into the exe -- unlike a conditionally-imported module.
"""
from __future__ import annotations

import datetime
import os

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))


def _resolve_key() -> str:
    """The bundled Groq key from local_secrets.py, else the GROQ_API_KEY env."""
    import sys
    sys.path.insert(0, ROOT)
    try:
        import local_secrets  # type: ignore  # gitignored
        k = (getattr(local_secrets, "GROQ_API_KEY", "") or "").strip()
        if k:
            return k
    except Exception:
        pass
    return (os.environ.get("GROQ_API_KEY", "") or "").strip()


def main() -> None:
    today = datetime.date.today().isoformat()
    key = _resolve_key()
    path = os.path.join(ROOT, "buildinfo.py")
    with open(path, "w", encoding="utf-8") as fh:
        fh.write(
            "# Auto-generated at build time by tools/stamp_build.py -- do not edit.\n"
            "# GITIGNORED: holds the bundled key; never commit this file.\n"
            f'BUILD_DATE = "{today}"\n'
            f'BUNDLED_GROQ_KEY = {key!r}\n'
        )
    print(f"Stamped BUILD_DATE = {today} | key bundled: {bool(key)}")


if __name__ == "__main__":
    main()
