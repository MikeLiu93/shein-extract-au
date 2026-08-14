"""Locate chrome.exe.

Windows registers Chrome under the "App Paths" registry key when it installs —
that's the canonical source and covers per-user installs done by any account,
not just LOCALAPPDATA of whoever's logged in right now. We fall through to
`shutil.which` (PATH) and a union of hardcoded install paths as safety nets,
and allow a manual override via the SHEIN_CHROME_PATH env var for portable /
custom installs.

Order (first match wins):
    1. SHEIN_CHROME_PATH env var
    2. Registry: HKCU + HKLM App Paths\\chrome.exe (32-bit view too)
    3. shutil.which("chrome.exe" / "chrome")
    4. Hardcoded install paths (Program Files, LOCALAPPDATA, /Applications)

`find_chrome()` returns the path string or None. `searched_locations()`
returns the list of places checked, for user-visible error messages.
"""
import os
import shutil
import sys
from pathlib import Path


_APP_PATHS_SUBKEY = r"SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\chrome.exe"

_KNOWN_PATHS = [
    r"C:\Program Files\Google\Chrome\Application\chrome.exe",
    r"C:\Program Files (x86)\Google\Chrome\Application\chrome.exe",
    os.path.expandvars(r"%LOCALAPPDATA%\Google\Chrome\Application\chrome.exe"),
    os.path.expandvars(r"%PROGRAMFILES%\Google\Chrome\Application\chrome.exe"),
    os.path.expandvars(r"%PROGRAMFILES(X86)%\Google\Chrome\Application\chrome.exe"),
    "/Applications/Google Chrome.app/Contents/MacOS/Google Chrome",
]


def _from_env() -> str | None:
    p = os.environ.get("SHEIN_CHROME_PATH", "").strip().strip('"')
    return p if p and Path(p).is_file() else None


def _from_registry() -> str | None:
    if sys.platform != "win32":
        return None
    try:
        import winreg  # type: ignore[import-not-found]
    except ImportError:
        return None
    hives = [
        (winreg.HKEY_CURRENT_USER, 0),
        (winreg.HKEY_LOCAL_MACHINE, 0),
        (winreg.HKEY_LOCAL_MACHINE, winreg.KEY_WOW64_32KEY),
    ]
    for hive, flags in hives:
        try:
            with winreg.OpenKey(hive, _APP_PATHS_SUBKEY, 0,
                                winreg.KEY_READ | flags) as k:
                val, _ = winreg.QueryValueEx(k, None)  # Default value = exe path
        except OSError:
            continue
        val = os.path.expandvars(str(val).strip().strip('"'))
        if val and Path(val).is_file():
            return val
    return None


def _from_path() -> str | None:
    return shutil.which("chrome.exe") or shutil.which("chrome")


def _from_known_paths() -> str | None:
    for p in _KNOWN_PATHS:
        if p and Path(p).is_file():
            return p
    return None


def find_chrome() -> str | None:
    """Return chrome.exe path, or None if not found."""
    for finder in (_from_env, _from_registry, _from_path, _from_known_paths):
        found = finder()
        if found:
            return found
    return None


def searched_locations() -> list[str]:
    """Human-readable list of everywhere we looked. Used in error messages."""
    out = ["SHEIN_CHROME_PATH env var"]
    if sys.platform == "win32":
        out += [
            r"HKCU\SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\chrome.exe",
            r"HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\chrome.exe",
            r"HKLM\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\App Paths\chrome.exe",
        ]
    out.append("PATH (chrome.exe / chrome)")
    out += _KNOWN_PATHS
    return out
