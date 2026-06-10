"""Self-update for the frozen Windows exe.

How updating a *running* onefile exe works on Windows:
  Windows refuses to DELETE or OVERWRITE a running .exe, but it DOES allow it to
  be RENAMED. So we:
    1. download the new exe next to the current one,
    2. rename the running exe  App.exe -> App.exe.old   (allowed while running),
    3. move the downloaded file into place as  App.exe,
    4. launch the new App.exe and exit,
    5. on the next startup, delete the leftover App.exe.old.

This module is UI-free so the engine stays reusable. The desktop UI drives it.
"""
from __future__ import annotations

import os
import re
import subprocess
import sys
import tempfile
from dataclasses import dataclass
from typing import Callable, Optional

import requests

DownloadProgress = Callable[[int, int], None]   # (bytes_done, bytes_total)


@dataclass
class UpdateInfo:
    version: str
    release_date: str
    url: str
    notes: str = ""


def is_frozen() -> bool:
    """True when running as the packaged exe (not `python main.py`)."""
    return bool(getattr(sys, "frozen", False))


def current_exe_path() -> str:
    return os.path.realpath(sys.executable)


def _parse_version(v: str) -> tuple[int, ...]:
    """'1.2.10' -> (1, 2, 10); tolerant of junk and missing parts."""
    nums = re.findall(r"\d+", v or "")
    return tuple(int(n) for n in nums) or (0,)


def is_newer(remote: str, local: str) -> bool:
    a, b = _parse_version(remote), _parse_version(local)
    # pad to equal length for a fair tuple comparison
    n = max(len(a), len(b))
    a += (0,) * (n - len(a))
    b += (0,) * (n - len(b))
    return a > b


def check_for_update(manifest_url: str, current_version: str,
                     timeout: float = 10.0) -> Optional[UpdateInfo]:
    """Fetch the manifest and return UpdateInfo if a newer version exists.

    Fails soft: any network/parse error returns None (never raises to caller).
    """
    if not manifest_url:
        return None
    try:
        resp = requests.get(manifest_url, timeout=timeout)
        if resp.status_code != 200:
            return None
        data = resp.json()
        version = str(data.get("version", "")).strip()
        url = str(data.get("url", "")).strip()
        if not version or not url:
            return None
        if not is_newer(version, current_version):
            return None
        return UpdateInfo(
            version=version,
            release_date=str(data.get("release_date", "")).strip(),
            url=url,
            notes=str(data.get("notes", "")).strip(),
        )
    except Exception:
        return None


def download_update(info: UpdateInfo, progress: Optional[DownloadProgress] = None,
                    timeout: float = 60.0) -> str:
    """Download the new installer (Setup.exe) to a temp file. Returns its path."""
    folder = tempfile.gettempdir()
    fd, tmp_path = tempfile.mkstemp(prefix="ExcelIntelligenceAgent-Setup-",
                                    suffix=".exe", dir=folder)
    os.close(fd)
    try:
        with requests.get(info.url, stream=True, timeout=timeout) as r:
            r.raise_for_status()
            total = int(r.headers.get("Content-Length", 0))
            done = 0
            with open(tmp_path, "wb") as fh:
                for chunk in r.iter_content(chunk_size=256 * 1024):
                    if not chunk:
                        continue
                    fh.write(chunk)
                    done += len(chunk)
                    if progress:
                        progress(done, total)
        if os.path.getsize(tmp_path) == 0:
            raise ValueError("Downloaded file was empty.")
        return tmp_path
    except Exception:
        if os.path.exists(tmp_path):
            os.remove(tmp_path)
        raise


def apply_update_and_restart(installer_path: str) -> None:
    """Run the downloaded installer silently, then exit. Does not return.

    The installer (Inno Setup, per-user) closes this app if needed, replaces the
    installed files, and relaunches the app via its [Run] postinstall entry.
    Only valid when frozen.
    """
    if not is_frozen():
        raise RuntimeError("Self-update is only available in the installed app.")

    # /VERYSILENT  : no UI; /SUPPRESSMSGBOXES : auto-answer prompts;
    # /CLOSEAPPLICATIONS : close us if a file is locked; /NORESTART : never reboot.
    args = [installer_path, "/VERYSILENT", "/SUPPRESSMSGBOXES",
            "/CLOSEAPPLICATIONS", "/NORESTART"]
    DETACHED = 0x00000008 | 0x00000200     # DETACHED_PROCESS | CREATE_NEW_GROUP
    subprocess.Popen(args, close_fds=True, creationflags=DETACHED)
    # Exit now so no installed file is locked; the installer relaunches the app.
    os._exit(0)


def cleanup_old() -> None:
    """No-op retained for backward compatibility (installer model needs none)."""
    return
