"""Stamp the current build date into buildinfo.py. Run by build.bat before packaging."""
from __future__ import annotations

import datetime
import os

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))


def main() -> None:
    today = datetime.date.today().isoformat()
    path = os.path.join(ROOT, "buildinfo.py")
    with open(path, "w", encoding="utf-8") as fh:
        fh.write(
            "# Auto-generated at build time by tools/stamp_build.py -- do not edit.\n"
            f'BUILD_DATE = "{today}"\n'
        )
    print(f"Stamped BUILD_DATE = {today}")


if __name__ == "__main__":
    main()
