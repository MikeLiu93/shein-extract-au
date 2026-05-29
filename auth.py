"""
Password gate with central revocation.

How it works
============
1. On launch we fetch a small JSON allowlist from a URL the owner controls
   (current default: raw.githubusercontent.com/<owner>/<repo>/main/auth.json).
2. The user types a password; we SHA-256 it and check membership.
3. We cache the verified hash + timestamp under %APPDATA%\\shein-extract\\.
   For the next 24h, launches are silent — we only re-prompt if the cache
   expires OR the cached hash gets removed/disabled in the upstream JSON.
4. Offline behavior: if the fetch fails but the local cache is still within
   its 24h window, we WARN and let the user proceed. If the cache is also
   expired or missing, we block.

Revoking access
===============
Edit auth.json in the repo, push commit. Within 24h every employee's exe
will detect the change on next launch and force a re-prompt with the new
password (or fail outright if their hash was removed).

auth.json schema (forward-compatible — supports per-team passwords later)
============================================================
{
  "version": 1,
  "active_passwords": [
    {"label": "shared",   "sha256": "<hex>", "active": true},
    {"label": "team-au",  "sha256": "<hex>", "active": true},
    {"label": "team-old", "sha256": "<hex>", "active": false}
  ]
}

The "label" is purely a human-readable tag for the owner — it does not
affect verification. To disable a password without deleting the row,
set "active": false (treated identically to removal).
"""

import getpass
import hashlib
import json
import os
import sys
import time
from pathlib import Path

import requests


AUTH_URL = os.environ.get(
    "SHEIN_AUTH_URL",
    "https://raw.githubusercontent.com/MikeLiu93/shein-extract-au/main/auth.json",
)
CACHE_FILE = (
    Path(os.environ.get("APPDATA", str(Path.home())))
    / "shein-extract-au"
    / "auth_cache.json"
)
CACHE_TTL_HOURS = 24
FETCH_TIMEOUT_SEC = 8
MAX_PROMPT_ATTEMPTS = 3


def _hash_password(pw: str) -> str:
    return hashlib.sha256(pw.encode("utf-8")).hexdigest().lower()


def _fetch_active_hashes() -> list[str]:
    """Returns lowercase hex digests of all currently-active passwords.
    Raises on any network/parse failure."""
    resp = requests.get(AUTH_URL, timeout=FETCH_TIMEOUT_SEC)
    resp.raise_for_status()
    data = resp.json()
    out = []
    for entry in data.get("active_passwords", []):
        if entry.get("active", True) is False:
            continue
        h = entry.get("sha256")
        if isinstance(h, str) and h:
            out.append(h.lower())
    if not out:
        raise RuntimeError("auth.json contains no active passwords")
    return out


def _load_cache() -> dict | None:
    if not CACHE_FILE.exists():
        return None
    try:
        return json.loads(CACHE_FILE.read_text(encoding="utf-8"))
    except Exception:
        return None


def _save_cache(hash_hex: str) -> None:
    CACHE_FILE.parent.mkdir(parents=True, exist_ok=True)
    CACHE_FILE.write_text(
        json.dumps({"verified_at": int(time.time()), "hash": hash_hex}),
        encoding="utf-8",
    )


def _hours_since(ts: int) -> float:
    return (time.time() - ts) / 3600.0


def gate() -> bool:
    """Returns True if the user is authorized to proceed, False otherwise.
    Prints user-facing messages on stdout for any required interaction."""
    cache = _load_cache()
    cache_hash = cache.get("hash") if cache else None
    cache_age = _hours_since(cache.get("verified_at", 0)) if cache else None

    # Try to fetch upstream allowlist
    try:
        live_hashes = _fetch_active_hashes()
        online = True
    except Exception as e:
        live_hashes = None
        online = False
        offline_reason = f"{e.__class__.__name__}: {e}"

    # Quick path: cached, fresh, and still in upstream list (or upstream unreachable but cache fresh)
    if cache_hash and cache_age is not None and cache_age < CACHE_TTL_HOURS:
        if online and cache_hash in live_hashes:
            return True
        if not online:
            remaining = CACHE_TTL_HOURS - cache_age
            print(f"[验证] 警告: 无法连接验证服务器（{offline_reason}）")
            print(f"       使用本地缓存（剩余 {remaining:.1f} 小时）")
            return True
        # online but cache_hash revoked → fall through to prompt
        print("[验证] 你之前的密码已被管理员撤销，请重新输入。")

    # No usable cache
    if not online:
        print("[验证] 错误: 无法连接验证服务器，且本地无有效缓存。")
        print(f"       原因: {offline_reason}")
        print(f"       请确保此电脑能访问 {AUTH_URL} 后重试。")
        return False

    # Online — prompt up to N times
    for attempt in range(MAX_PROMPT_ATTEMPTS):
        try:
            pw = getpass.getpass("请输入访问密码（联系管理员获取）: ")
        except (EOFError, KeyboardInterrupt):
            print("\n[验证] 用户取消")
            return False
        if not pw:
            continue
        h = _hash_password(pw)
        if h in live_hashes:
            _save_cache(h)
            print("[验证] 通过 ✓")
            return True
        remaining = MAX_PROMPT_ATTEMPTS - attempt - 1
        if remaining > 0:
            print(f"[验证] 密码错误，还剩 {remaining} 次")

    print("[验证] 失败次数过多，退出")
    return False


if __name__ == "__main__":
    sys.exit(0 if gate() else 1)
