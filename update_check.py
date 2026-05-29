"""
Auto-update checker — queries GitHub Releases API, prompts user, swaps the .exe.

Behavior (Mike's Q2/Q3 decisions):
  - Skip if last check was < 24h ago (limit GitHub API calls).
  - If new version exists: Tkinter dialog "立即更新 / 本次跳过"
    (no "skip this version" option — always re-prompt next time).
  - On update accept: download new .exe to %TEMP%, spawn a tiny .bat
    that waits 2s + replaces the old .exe + relaunches it; current
    process exits.

Module is a no-op when:
  - Running from source (sys.frozen not set) — devs use git pull instead
  - Network down / GitHub unreachable
  - last_update_check.json says we just checked
"""

import json
import os
import subprocess
import sys
import time
from datetime import datetime, timedelta
from pathlib import Path

import requests

from version import VERSION

GITHUB_API = "https://api.github.com/repos/MikeLiu93/shein-extract-au/releases/latest"
CHECK_INTERVAL_HOURS = 24
USER_DATA_DIR = Path(os.environ.get("APPDATA", str(Path.home()))) / "shein-extract-au"
LAST_CHECK_FILE = USER_DATA_DIR / "last_update_check.json"


def _load_last_check() -> dict:
    if not LAST_CHECK_FILE.exists():
        return {}
    try:
        return json.loads(LAST_CHECK_FILE.read_text(encoding="utf-8"))
    except Exception:
        return {}


def _save_last_check(data: dict) -> None:
    USER_DATA_DIR.mkdir(parents=True, exist_ok=True)
    LAST_CHECK_FILE.write_text(json.dumps(data, indent=2), encoding="utf-8")


def _should_check() -> bool:
    last = _load_last_check()
    last_iso = last.get("last_check_iso")
    if not last_iso:
        return True
    try:
        last_dt = datetime.fromisoformat(last_iso)
    except ValueError:
        return True
    return datetime.now() - last_dt > timedelta(hours=CHECK_INTERVAL_HOURS)


def _parse_version_tuple(s: str) -> tuple:
    """'v3.5.1' or '3.5.1' → (3, 5, 1). Trailing -beta.X → (3,5,1, -1, X)."""
    s = (s or "").lstrip("vV")
    parts = []
    for chunk in s.split("."):
        m = ""
        for c in chunk:
            if c.isdigit():
                m += c
            else:
                break
        if m:
            parts.append(int(m))
    return tuple(parts) if parts else (0,)


def _is_newer(remote_tag: str, current: str) -> bool:
    return _parse_version_tuple(remote_tag) > _parse_version_tuple(current)


def _fetch_latest_release(timeout: int = 5) -> dict | None:
    try:
        r = requests.get(GITHUB_API, timeout=timeout,
                         headers={"Accept": "application/vnd.github+json"})
        if r.status_code != 200:
            return None
        return r.json()
    except Exception:
        return None


def _show_dialog(latest_tag: str, body: str) -> bool:
    """Returns True if user wants to update, False to skip."""
    try:
        import tkinter as tk
        from tkinter import messagebox
    except Exception:
        # Headless / no Tk — fall back to console prompt
        print(f"\n[更新] 发现新版本 {latest_tag}（当前 v{VERSION}）。")
        ans = input("立即更新? [y/N]: ").strip().lower()
        return ans in ("y", "yes")

    root = tk.Tk()
    root.withdraw()
    title = f"发现新版本 {latest_tag}"
    msg = (
        f"当前版本: v{VERSION}\n"
        f"最新版本: {latest_tag}\n\n"
        f"{body or ''}\n\n"
        "是否立即更新？"
    )
    answer = messagebox.askyesno(title, msg)
    root.destroy()
    return bool(answer)


def _find_exe_asset(release: dict) -> str | None:
    """Look for an installer .exe in the release assets."""
    for asset in release.get("assets", []) or []:
        name = (asset.get("name") or "").lower()
        if name.endswith(".exe") and ("setup" in name or "installer" in name):
            return asset.get("browser_download_url")
    # fallback: any .exe asset
    for asset in release.get("assets", []) or []:
        if (asset.get("name") or "").lower().endswith(".exe"):
            return asset.get("browser_download_url")
    return None


def _download(url: str, dest: Path) -> bool:
    try:
        with requests.get(url, stream=True, timeout=60) as r:
            r.raise_for_status()
            with dest.open("wb") as f:
                for chunk in r.iter_content(chunk_size=1024 * 256):
                    if chunk:
                        f.write(chunk)
        return True
    except Exception as e:
        print(f"[更新] 下载失败: {e}")
        return False


def _spawn_replace_helper(new_exe: Path, current_exe: Path) -> None:
    """
    Spawn a .bat that waits 2s, replaces current_exe with new_exe, then runs it.
    The .bat self-deletes after launch.
    """
    bat = Path(os.environ.get("TEMP", ".")) / "shein_extract_au_update.bat"
    script = (
        "@echo off\r\n"
        "timeout /t 2 /nobreak >nul\r\n"
        f'copy /y "{new_exe}" "{current_exe}" >nul\r\n'
        f'del "{new_exe}" 2>nul\r\n'
        f'start "" "{current_exe}"\r\n'
        '(goto) 2>nul & del "%~f0"\r\n'
    )
    bat.write_text(script, encoding="ascii")
    subprocess.Popen(
        ["cmd", "/c", str(bat)],
        creationflags=getattr(subprocess, "DETACHED_PROCESS", 0)
                      | getattr(subprocess, "CREATE_NEW_PROCESS_GROUP", 0),
        close_fds=True,
    )


def check_for_update(blocking: bool = False) -> None:
    """
    Called once per app launch. blocking=True actually shows the dialog;
    blocking=False just checks but doesn't prompt (we still want the dialog
    by default — keeping this param for future "auto download silently" mode).

    No-op if running from a Python source checkout (not a frozen .exe).
    """
    # Only meaningful when we're a frozen .exe. From source, devs git pull.
    if not getattr(sys, "frozen", False):
        return

    if not _should_check():
        return

    release = _fetch_latest_release()
    # Always update last_check_iso even on failure — don't hammer the API
    _save_last_check({
        "last_check_iso": datetime.now().isoformat(timespec="seconds"),
        "current_version": VERSION,
    })

    if not release:
        return

    tag = release.get("tag_name") or ""
    if not tag or not _is_newer(tag, VERSION):
        return

    body = (release.get("body") or "").strip()[:500]
    notes = body or "（无更新说明）"

    if not _show_dialog(tag, notes):
        return  # User said "本次跳过", continue with current version

    download_url = _find_exe_asset(release)
    if not download_url:
        print("[更新] release 里找不到 .exe 资产，跳过。")
        return

    print(f"[更新] 下载 {tag} ...")
    new_exe = Path(os.environ.get("TEMP", ".")) / f"SheinExtractAU-{tag.lstrip('v')}.exe"
    if not _download(download_url, new_exe):
        return

    print(f"[更新] 下载完成，准备替换并重启...")
    current_exe = Path(sys.executable).resolve()
    _spawn_replace_helper(new_exe, current_exe)
    time.sleep(0.5)
    sys.exit(0)
