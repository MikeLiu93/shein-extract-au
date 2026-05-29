# AU Installer Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Package `shein-extract-au` as a single-icon Windows installable app (`SheinExtractAU.exe`) where every launch shows a console menu: [1] configure paths / [2] run pipeline / [Q] quit. Coexists cleanly with `shein-extract`.

**Architecture:** Mirror 母项目's PyInstaller + Inno Setup chain. App entry (`app_main.py`) is console-based menu driving Tkinter wizard (`setup_wizard.py`) or `run_excel.main()`. Shared bits (`auth.py`, `update_check.py`, `make_password_hash.py`) ported verbatim with AU-specific URLs/dirs. Coexistence via separate install dir, AppId, config dir, Chrome profile (`shein-cdp-profile-au`), CDP port (`9223`), release feed.

**Tech Stack:** Python 3.10+ (Anaconda), Tkinter (wizard + update dialog), `requests`, `openpyxl`, PyInstaller, Inno Setup 6.x.

**Source spec:** `docs/superpowers/specs/2026-05-29-au-installer-design.md`
**Reference repo (port source):** `C:\Users\ak\Desktop\Claude\shein-extract\` (read-only; copy files from here)

---

## File Structure

| File | Action | Responsibility |
|---|---|---|
| `version.py` | Create | Single source of truth for `VERSION` |
| `auth.py` | Create (port) | Password gate + central revocation |
| `auth.json` | Create | Active password hash list (committed; SHA-256 of `Ace2025`) |
| `make_password_hash.py` | Create (port verbatim) | Admin util — generate hash for `auth.json` |
| `update_check.py` | Create (port) | Check GitHub Releases, prompt, download, swap exe |
| `setup_wizard.py` | Create | Tkinter 5-step wizard, writes `%APPDATA%\shein-extract-au\config.env` |
| `app_main.py` | Create | PyInstaller entry; password gate → update check → console menu → wizard or run_excel |
| `pyinstaller.spec` | Create | PyInstaller build spec for `SheinExtractAU.exe` |
| `build.bat` | Create | One-shot build: PyInstaller + Inno Setup |
| `installer.iss` | Create | Inno Setup script — 1 program icon, separate AppId/install dir |
| `INSTALL_GUIDE_CN.md` | Create | End-user install instructions |
| `release.bat` | Create | Semi-automated release: bump version → build → tag → push → open GitHub Release URL |
| `test_setup_wizard.py` | Create | Tests for wizard's pure helpers (config I/O, validation) |
| `test_app_main_menu.py` | Create | Tests for menu logic and config validity check |
| `shein_scraper.py` | Modify | Change `PERSISTENT_PROFILE_DIR` → `shein-cdp-profile-au`, `CDP_PORT` → `9223` |

---

## Task 1: Create `version.py`

**Files:**
- Create: `version.py`

- [ ] **Step 1: Write the file**

```python
"""
Single source of truth for the AU app version.
Bump here, rebuild, tag git as v{VERSION}, push, create GitHub release.
update_check.py compares this against GitHub's latest release tag_name.
"""

VERSION = "0.1.0"
```

- [ ] **Step 2: Verify import works**

Run from `shein-extract-au`:
```
C:\Users\ak\anaconda3\python.exe -c "from version import VERSION; print(VERSION)"
```
Expected: `0.1.0`

- [ ] **Step 3: Commit**

```bash
git add version.py
git commit -m "feat(installer): add version.py at 0.1.0"
```

---

## Task 2: Port `make_password_hash.py` and generate `auth.json`

**Files:**
- Create: `make_password_hash.py` (verbatim copy from 母项目)
- Create: `auth.json` (with SHA-256 of `Ace2025`)

- [ ] **Step 1: Copy `make_password_hash.py` verbatim**

```bash
cp ../shein-extract/make_password_hash.py make_password_hash.py
```

No edits — the file is project-agnostic.

- [ ] **Step 2: Verify it runs**

```
C:\Users\ak\anaconda3\python.exe make_password_hash.py --plain
```
Type `test` twice. Expected output: SHA-256 hex string.

- [ ] **Step 3: Compute hash of `Ace2025`**

```
C:\Users\ak\anaconda3\python.exe -c "import hashlib; print(hashlib.sha256('Ace2025'.encode()).hexdigest())"
```
Record the resulting hex (call it `<HASH>`).

- [ ] **Step 4: Write `auth.json`**

Create `auth.json` with this exact structure (replace `<HASH>` with the value from Step 3):

```json
{
  "version": 1,
  "active_passwords": [
    {"label": "shared-au", "sha256": "<HASH>", "active": true}
  ]
}
```

- [ ] **Step 5: Verify hash matches what `auth.py` will compute (sanity check)**

```
C:\Users\ak\anaconda3\python.exe -c "
import json, hashlib
data = json.load(open('auth.json'))
expected = hashlib.sha256('Ace2025'.encode()).hexdigest()
got = data['active_passwords'][0]['sha256']
assert got == expected, (got, expected)
print('OK')
"
```
Expected: `OK`

- [ ] **Step 6: Commit**

```bash
git add make_password_hash.py auth.json
git commit -m "feat(installer): port make_password_hash.py and seed auth.json with Ace2025 hash"
```

---

## Task 3: Port `auth.py` with AU URL and cache dir

**Files:**
- Create: `auth.py` (copy from 母项目 with 2 string changes)

- [ ] **Step 1: Copy from 母项目**

```bash
cp ../shein-extract/auth.py auth.py
```

- [ ] **Step 2: Edit `auth.py` — change `AUTH_URL` default**

In `auth.py`, find the `AUTH_URL` assignment (around line 49-52):

```python
AUTH_URL = os.environ.get(
    "SHEIN_AUTH_URL",
    "https://raw.githubusercontent.com/MikeLiu93/shein-extract/main/auth.json",
)
```

Change the URL to point to AU's repo:

```python
AUTH_URL = os.environ.get(
    "SHEIN_AUTH_URL",
    "https://raw.githubusercontent.com/MikeLiu93/shein-extract-au/main/auth.json",
)
```

- [ ] **Step 3: Edit `auth.py` — change `CACHE_FILE` subdir**

Find the `CACHE_FILE` assignment (around line 53-57):

```python
CACHE_FILE = (
    Path(os.environ.get("APPDATA", str(Path.home())))
    / "shein-extract"
    / "auth_cache.json"
)
```

Change subdir to `shein-extract-au`:

```python
CACHE_FILE = (
    Path(os.environ.get("APPDATA", str(Path.home())))
    / "shein-extract-au"
    / "auth_cache.json"
)
```

- [ ] **Step 4: Smoke test — local gate against the staged auth.json**

Since `auth.json` is not yet pushed to GitHub, the live fetch will fail. The fallback path is what we test. First wipe any cache, then exercise the fallback by pointing `SHEIN_AUTH_URL` at the local file via `file://` URL — this won't work cleanly with `requests`, so instead just verify the password hash logic standalone:

```
C:\Users\ak\anaconda3\python.exe -c "
from auth import _hash_password
import json
expected = json.load(open('auth.json'))['active_passwords'][0]['sha256']
got = _hash_password('Ace2025')
assert got == expected
print('hash OK')
"
```
Expected: `hash OK`

- [ ] **Step 5: Commit**

```bash
git add auth.py
git commit -m "feat(installer): port auth.py with AU URL and cache dir"
```

---

## Task 4: Port `update_check.py` with AU API URL and naming

**Files:**
- Create: `update_check.py` (copy from 母项目 with 4 string changes)

- [ ] **Step 1: Copy from 母项目**

```bash
cp ../shein-extract/update_check.py update_check.py
```

- [ ] **Step 2: Edit — change `GITHUB_API` URL**

Find (around line 30):
```python
GITHUB_API = "https://api.github.com/repos/MikeLiu93/shein-extract/releases/latest"
```

Change to:
```python
GITHUB_API = "https://api.github.com/repos/MikeLiu93/shein-extract-au/releases/latest"
```

- [ ] **Step 3: Edit — change `USER_DATA_DIR`**

Find (around line 32):
```python
USER_DATA_DIR = Path(os.environ.get("APPDATA", str(Path.home()))) / "shein-extract"
```

Change to:
```python
USER_DATA_DIR = Path(os.environ.get("APPDATA", str(Path.home()))) / "shein-extract-au"
```

- [ ] **Step 4: Edit — change update-helper bat name**

Find (around line 150) inside `_spawn_replace_helper`:
```python
bat = Path(os.environ.get("TEMP", ".")) / "shein_extract_update.bat"
```

Change to:
```python
bat = Path(os.environ.get("TEMP", ".")) / "shein_extract_au_update.bat"
```

- [ ] **Step 5: Edit — change download dest filename**

Find (around line 209):
```python
new_exe = Path(os.environ.get("TEMP", ".")) / f"SheinExtract-{tag.lstrip('v')}.exe"
```

Change to:
```python
new_exe = Path(os.environ.get("TEMP", ".")) / f"SheinExtractAU-{tag.lstrip('v')}.exe"
```

- [ ] **Step 6: Smoke test — import + no-op when not frozen**

```
C:\Users\ak\anaconda3\python.exe -c "
import update_check
update_check.check_for_update()
print('no-op OK')
"
```
Expected: `no-op OK` (returns immediately because `sys.frozen` is False from source)

- [ ] **Step 7: Commit**

```bash
git add update_check.py
git commit -m "feat(installer): port update_check.py with AU API URL and naming"
```

---

## Task 5: Modify `shein_scraper.py` constants for AU coexistence

**Files:**
- Modify: `shein_scraper.py:89` (PERSISTENT_PROFILE_DIR)
- Modify: `shein_scraper.py` (CDP_PORT — locate via grep)

- [ ] **Step 1: Find CDP_PORT line**

```
grep -n "^CDP_PORT" shein_scraper.py
```
Note the line number.

- [ ] **Step 2: Edit `PERSISTENT_PROFILE_DIR`**

Find (line 89):
```python
PERSISTENT_PROFILE_DIR = os.path.join(os.path.expanduser("~"), "shein-cdp-profile")
```

Change to:
```python
PERSISTENT_PROFILE_DIR = os.path.join(os.path.expanduser("~"), "shein-cdp-profile-au")
```

- [ ] **Step 3: Edit `CDP_PORT`**

Find (whatever line — grep showed earlier):
```python
CDP_PORT = 9222
```

Change to:
```python
CDP_PORT = 9223
```

- [ ] **Step 4: Smoke test — run existing test suite (must still pass)**

```
C:\Users\ak\anaconda3\python.exe test_variant_merge.py
```
Expected: `ALL PASS`

- [ ] **Step 5: Compile-check**

```
C:\Users\ak\anaconda3\python.exe -m py_compile shein_scraper.py run_excel.py
```
Expected: silent (no output).

- [ ] **Step 6: Commit**

```bash
git add shein_scraper.py
git commit -m "feat(installer): separate AU Chrome profile dir + CDP port (9223)"
```

---

## Task 6: Write failing tests for `setup_wizard.py` pure helpers

**Files:**
- Create: `test_setup_wizard.py`

The wizard's testable pure helpers (not the Tkinter UI):
- `write_config_env(values: dict) -> Path` — writes KEY=VALUE lines to `config.env`, returns path
- `load_existing_config(path: Path) -> dict` — parses `config.env` back to dict
- `is_first_run_complete(path: Path | None = None) -> bool` — returns True if config exists and has `SHEIN_SUBMITTED_DIR`
- `mask_api_key(key: str) -> str` — `sk-ant-api03-aaa...zzzz` style (first 8 + `...` + last 4); empty → `""`

- [ ] **Step 1: Write the failing test file**

```python
"""Tests for setup_wizard's pure helpers (Tkinter UI not tested here)."""
import tempfile
from pathlib import Path

import setup_wizard as sw


def test_write_then_load_round_trip():
    with tempfile.TemporaryDirectory() as td:
        p = Path(td) / "config.env"
        sw.write_config_env({
            "SHEIN_SUBMITTED_DIR": "D:\\a\\b",
            "SHEIN_INPUT_FILENAME": "希音链接 - LU.xlsx",
        }, path=p)
        loaded = sw.load_existing_config(p)
        assert loaded["SHEIN_SUBMITTED_DIR"] == "D:\\a\\b"
        assert loaded["SHEIN_INPUT_FILENAME"] == "希音链接 - LU.xlsx"


def test_load_missing_file_returns_empty():
    assert sw.load_existing_config(Path("/__nonexistent__")) == {}


def test_is_first_run_complete_false_when_missing():
    assert sw.is_first_run_complete(Path("/__nonexistent__")) is False


def test_is_first_run_complete_true_when_submitted_dir_set():
    with tempfile.TemporaryDirectory() as td:
        p = Path(td) / "config.env"
        p.write_text("SHEIN_SUBMITTED_DIR=D:\\x\n", encoding="utf-8")
        assert sw.is_first_run_complete(p) is True


def test_is_first_run_complete_false_when_only_other_keys():
    with tempfile.TemporaryDirectory() as td:
        p = Path(td) / "config.env"
        p.write_text("ANTHROPIC_API_KEY=sk-foo\n", encoding="utf-8")
        assert sw.is_first_run_complete(p) is False


def test_mask_api_key_short_or_empty():
    assert sw.mask_api_key("") == ""
    assert sw.mask_api_key("short") == "short"  # too short to mask meaningfully


def test_mask_api_key_long():
    k = "sk-ant-api03-abcdefghijklmnopqrstuvwxyz"
    m = sw.mask_api_key(k)
    assert m.startswith("sk-ant-a"), m
    assert m.endswith("wxyz"), m
    assert "..." in m


if __name__ == "__main__":
    test_write_then_load_round_trip()
    test_load_missing_file_returns_empty()
    test_is_first_run_complete_false_when_missing()
    test_is_first_run_complete_true_when_submitted_dir_set()
    test_is_first_run_complete_false_when_only_other_keys()
    test_mask_api_key_short_or_empty()
    test_mask_api_key_long()
    print("ALL PASS")
```

- [ ] **Step 2: Run — verify it fails (import error: no setup_wizard module yet)**

```
C:\Users\ak\anaconda3\python.exe test_setup_wizard.py
```
Expected: `ModuleNotFoundError: No module named 'setup_wizard'` (or import-time error on missing function names — either is "red")

---

## Task 7: Write `setup_wizard.py` to make tests pass + add Tkinter UI

**Files:**
- Create: `setup_wizard.py`

The wizard has two layers:
- **Pure helpers** (testable) — `write_config_env`, `load_existing_config`, `is_first_run_complete`, `mask_api_key`, `validate_paths`
- **Tkinter UI** — `Wizard` class, `run_wizard()`

5 wizard steps (simpler than 母项目, per Spec §6):
1. Welcome
2. System check (Chrome)
3. Path config (3 fields: SUBMITTED_DIR, OUTPUT_DIR, INPUT_FILENAME)
4. API key entry (ANTHROPIC_API_KEY)
5. Done — writes config.env

- [ ] **Step 1: Write the file**

```python
"""
AU setup wizard (Tkinter, 5 steps).

Writes %APPDATA%\\shein-extract-au\\config.env with the user's choices.
Steps:
  1. Welcome
  2. System check (Chrome)
  3. Path config (3 fields)
  4. Anthropic API key
  5. Done (write config.env)
"""
import os
import re
import sys
import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox, ttk

USER_DATA_DIR = Path(os.environ.get("APPDATA", str(Path.home()))) / "shein-extract-au"
CONFIG_FILE = USER_DATA_DIR / "config.env"

CHROME_CANDIDATES = [
    r"C:\Program Files\Google\Chrome\Application\chrome.exe",
    r"C:\Program Files (x86)\Google\Chrome\Application\chrome.exe",
    os.path.expandvars(r"%LOCALAPPDATA%\Google\Chrome\Application\chrome.exe"),
]

# Default starting point for SUBMITTED_DIR — same default config.py uses
DEFAULT_SUBMITTED_DIR = r"D:\共享云端硬盘\02 希音\澳洲站"

# ── Pure helpers ─────────────────────────────────────────────────────────────


def write_config_env(values: dict, path: Path | None = None) -> Path:
    """Write KEY=VALUE lines (UTF-8). Creates parent dirs."""
    p = path or CONFIG_FILE
    p.parent.mkdir(parents=True, exist_ok=True)
    lines = [
        "# Generated by SheinExtractAU setup wizard. Edit by re-running the wizard",
        "# or by deleting this file before launching the app.",
        "",
    ]
    for k, v in values.items():
        lines.append(f"{k}={v}")
    lines.append("")
    p.write_text("\n".join(lines), encoding="utf-8")
    return p


def load_existing_config(path: Path | None = None) -> dict:
    """Parse a KEY=VALUE config.env file back to a dict. Missing file → {}."""
    p = path or CONFIG_FILE
    if not p.exists():
        return {}
    out = {}
    try:
        text = p.read_text(encoding="utf-8")
    except OSError:
        return {}
    for line in text.splitlines():
        line = line.strip()
        if not line or line.startswith("#") or "=" not in line:
            continue
        k, v = line.split("=", 1)
        out[k.strip()] = v.strip()
    return out


def is_first_run_complete(path: Path | None = None) -> bool:
    p = path or CONFIG_FILE
    if not p.exists():
        return False
    try:
        text = p.read_text(encoding="utf-8")
    except OSError:
        return False
    return bool(re.search(r"^\s*SHEIN_SUBMITTED_DIR\s*=\s*\S", text, re.M))


def mask_api_key(key: str) -> str:
    """sk-ant-api03-xxxxxxxx → sk-ant-a...xxxx style. Short/empty unchanged."""
    if not key or len(key) < 16:
        return key
    return f"{key[:8]}...{key[-4:]}"


def find_chrome() -> str | None:
    for p in CHROME_CANDIDATES:
        if Path(p).exists():
            return p
    return None


# ── Wizard (Tkinter) ─────────────────────────────────────────────────────────


class Wizard:
    def __init__(self):
        self.root = tk.Tk()
        self.root.title("SHEIN 上架工具 AU — 配置")
        self.root.geometry("680x500")
        self.root.resizable(False, False)
        try:
            self.root.attributes("-topmost", True)
            self.root.after(500, lambda: self.root.attributes("-topmost", False))
        except Exception:
            pass

        self.completed = False
        self.cancelled = False

        self.values = {
            "submitted_dir": DEFAULT_SUBMITTED_DIR,
            "output_dir": "",
            "input_filename": "希音链接 - LU.xlsx",
            "api_key": "",
        }
        self._load_existing()

        self.frame = ttk.Frame(self.root, padding=20)
        self.frame.pack(fill="both", expand=True)

        self.steps = [
            self._step_welcome,
            self._step_check,
            self._step_paths,
            self._step_api,
            self._step_done,
        ]
        self.step_idx = 0
        self._render()

    def _load_existing(self) -> None:
        cfg = load_existing_config()
        mapping = {
            "SHEIN_SUBMITTED_DIR": "submitted_dir",
            "SHEIN_OUTPUT_DIR": "output_dir",
            "SHEIN_INPUT_FILENAME": "input_filename",
            "ANTHROPIC_API_KEY": "api_key",
        }
        for env_k, state_k in mapping.items():
            if env_k in cfg:
                self.values[state_k] = cfg[env_k]

    def _clear(self):
        for w in self.frame.winfo_children():
            w.destroy()

    def _render(self):
        self._clear()
        self.steps[self.step_idx]()

    def _next(self):
        if self.step_idx < len(self.steps) - 1:
            self.step_idx += 1
            self._render()

    def _back(self):
        if self.step_idx > 0:
            self.step_idx -= 1
            self._render()

    def _heading(self, text):
        ttk.Label(self.frame, text=text,
                  font=("Microsoft YaHei", 16, "bold")).pack(anchor="w", pady=(0, 6))
        ttk.Separator(self.frame, orient="horizontal").pack(fill="x", pady=(0, 12))

    def _para(self, text):
        ttk.Label(self.frame, text=text, wraplength=620, justify="left",
                  font=("Microsoft YaHei", 10)).pack(anchor="w", pady=(0, 8))

    def _nav(self, on_next=None, next_text="下一步", show_back=True):
        bar = ttk.Frame(self.frame)
        bar.pack(side="bottom", fill="x", pady=(20, 0))
        ttk.Button(bar, text="取消", command=self._cancel).pack(side="left")
        if show_back and self.step_idx > 0:
            ttk.Button(bar, text="上一步", command=self._back).pack(side="right", padx=4)
        if on_next:
            ttk.Button(bar, text=next_text, command=on_next).pack(side="right", padx=4)

    def _cancel(self):
        if messagebox.askyesno("确认", "现在退出向导？未保存的修改会丢失。"):
            self.cancelled = True
            self.root.destroy()

    # Steps ───────────────────────────────────────────────────────────────────
    def _step_welcome(self):
        self._heading("SHEIN 上架工具 AU — 配置")
        self._para(
            "这个向导帮你配置输入/输出目录和 API key（约 1-2 分钟）。\n\n"
            "你已经配置过的值会预填进来。点【开始】继续。"
        )
        self._nav(on_next=self._next, next_text="开始", show_back=False)

    def _step_check(self):
        self._heading("系统检查")
        chrome = find_chrome()
        if chrome:
            ttk.Label(self.frame, text=f"✓ 找到 Chrome：{chrome}",
                      foreground="green", font=("Microsoft YaHei", 10),
                      wraplength=620, justify="left").pack(anchor="w", pady=4)
            self._nav(on_next=self._next)
        else:
            ttk.Label(self.frame,
                      text="✗ 未找到 Chrome —— 请先装 https://www.google.com/chrome/ 再重开向导。",
                      foreground="red", font=("Microsoft YaHei", 10),
                      wraplength=620, justify="left").pack(anchor="w", pady=4)
            self._nav(on_next=None)

    def _step_paths(self):
        self._heading("路径设置")
        self._para("确认或修改以下 3 项。")

        rows = [
            ("输入表所在目录", "submitted_dir", True),
            ("输出根目录（留空 = 自动用 输入目录\\上架资料-已完成）", "output_dir", True),
            ("指定输入文件名（可选；留空 = 处理目录下所有 .xlsx）", "input_filename", False),
        ]
        self._entries = {}
        for label, key, browseable in rows:
            row = ttk.Frame(self.frame)
            row.pack(fill="x", pady=4)
            ttk.Label(row, text=label, width=30, anchor="w",
                      font=("Microsoft YaHei", 10)).pack(side="left")
            var = tk.StringVar(value=self.values.get(key, ""))
            ttk.Entry(row, textvariable=var, font=("Consolas", 9)).pack(
                side="left", fill="x", expand=True, padx=4)
            self._entries[key] = var
            if browseable:
                ttk.Button(row, text="浏览...",
                           command=lambda v=var: self._browse(v)).pack(side="left")

        def on_next():
            for k, var in self._entries.items():
                self.values[k] = var.get().strip()
            sd = Path(self.values["submitted_dir"])
            if not sd.is_dir():
                messagebox.showerror("路径错误", f"输入目录不存在:\n{sd}")
                return
            if not self.values["output_dir"]:
                self.values["output_dir"] = str(sd / "上架资料-已完成")
            od = Path(self.values["output_dir"])
            try:
                od.mkdir(parents=True, exist_ok=True)
            except OSError as e:
                messagebox.showerror("无法创建输出目录", str(e))
                return
            self._next()

        self._nav(on_next=on_next)

    def _browse(self, var: tk.StringVar):
        cur = var.get() or os.path.expanduser("~")
        d = filedialog.askdirectory(initialdir=cur, title="选择文件夹")
        if d:
            var.set(d)

    def _step_api(self):
        self._heading("Anthropic API Key")
        self._para(
            "用于 Claude Haiku 生成 eBay 标题。可以留空（不启用 AI 标题）。"
        )
        existing = self.values.get("api_key", "")
        if existing:
            self._para(f"现有 key：{mask_api_key(existing)}\n"
                       f"输入 “keep” 或留空 = 保留现值。")
        var = tk.StringVar(value="")
        ttk.Entry(self.frame, textvariable=var, font=("Consolas", 9),
                  width=70).pack(anchor="w", pady=4)

        def on_next():
            v = var.get().strip()
            if v and v.lower() != "keep":
                self.values["api_key"] = v
            # else keep existing
            self._next()

        self._nav(on_next=on_next)

    def _step_done(self):
        self._heading("完成")
        env_values = {
            "SHEIN_SUBMITTED_DIR": self.values["submitted_dir"],
            "SHEIN_OUTPUT_DIR": self.values["output_dir"],
        }
        if self.values["input_filename"]:
            env_values["SHEIN_INPUT_FILENAME"] = self.values["input_filename"]
        if self.values["api_key"]:
            env_values["ANTHROPIC_API_KEY"] = self.values["api_key"]
        try:
            written = write_config_env(env_values)
            self._para(f"配置已保存到：\n  {written}")
        except OSError as e:
            messagebox.showerror("保存失败", str(e))
            return

        bar = ttk.Frame(self.frame)
        bar.pack(side="bottom", fill="x", pady=(20, 0))
        ttk.Button(bar, text="关闭", command=self._finish).pack(side="right", padx=4)

    def _finish(self):
        self.completed = True
        self.root.destroy()

    def run(self) -> bool:
        self.root.mainloop()
        return self.completed and not self.cancelled


def run_wizard() -> bool:
    return Wizard().run()


if __name__ == "__main__":
    sys.exit(0 if run_wizard() else 1)
```

- [ ] **Step 2: Run tests — verify they pass**

```
C:\Users\ak\anaconda3\python.exe test_setup_wizard.py
```
Expected: `ALL PASS`

- [ ] **Step 3: Compile-check**

```
C:\Users\ak\anaconda3\python.exe -m py_compile setup_wizard.py
```
Expected: silent.

- [ ] **Step 4: Commit**

```bash
git add setup_wizard.py test_setup_wizard.py
git commit -m "feat(installer): setup_wizard.py with tested pure helpers + Tkinter UI"
```

---

## Task 8: Write failing tests for `app_main` menu logic

**Files:**
- Create: `test_app_main_menu.py`

Testable pure helpers in `app_main.py`:
- `read_menu_choice(input_fn=input) -> str` — prompts, returns `"1"`, `"2"`, or `"Q"` (uppercase for q/Q/empty); rejects others by re-prompting
- `has_valid_config(config_file: Path | None = None) -> bool` — wraps `setup_wizard.is_first_run_complete`

- [ ] **Step 1: Write the failing test file**

```python
"""Tests for app_main's menu logic and config check."""
import tempfile
from pathlib import Path

import app_main


def _input_seq(*responses):
    """Returns a callable that yields successive responses."""
    it = iter(responses)
    return lambda prompt="": next(it)


def test_menu_returns_1():
    assert app_main.read_menu_choice(_input_seq("1")) == "1"


def test_menu_returns_2():
    assert app_main.read_menu_choice(_input_seq("2")) == "2"


def test_menu_q_lower_returns_Q():
    assert app_main.read_menu_choice(_input_seq("q")) == "Q"


def test_menu_q_upper_returns_Q():
    assert app_main.read_menu_choice(_input_seq("Q")) == "Q"


def test_menu_empty_returns_Q():
    assert app_main.read_menu_choice(_input_seq("")) == "Q"


def test_menu_invalid_then_valid():
    # First "x" rejected, then "2" accepted
    assert app_main.read_menu_choice(_input_seq("x", "2")) == "2"


def test_has_valid_config_false_when_missing():
    assert app_main.has_valid_config(Path("/__nonexistent__")) is False


def test_has_valid_config_true_when_submitted_dir_set():
    with tempfile.TemporaryDirectory() as td:
        p = Path(td) / "config.env"
        p.write_text("SHEIN_SUBMITTED_DIR=D:\\x\n", encoding="utf-8")
        assert app_main.has_valid_config(p) is True


if __name__ == "__main__":
    test_menu_returns_1()
    test_menu_returns_2()
    test_menu_q_lower_returns_Q()
    test_menu_q_upper_returns_Q()
    test_menu_empty_returns_Q()
    test_menu_invalid_then_valid()
    test_has_valid_config_false_when_missing()
    test_has_valid_config_true_when_submitted_dir_set()
    print("ALL PASS")
```

- [ ] **Step 2: Run — verify it fails**

```
C:\Users\ak\anaconda3\python.exe test_app_main_menu.py
```
Expected: `ModuleNotFoundError: No module named 'app_main'`

---

## Task 9: Write `app_main.py` — menu-driven entry

**Files:**
- Create: `app_main.py`

- [ ] **Step 1: Write the file**

```python
"""
PyInstaller entry point for SheinExtractAU.

Flow on every launch:
  1. UTF-8 console init + version banner
  2. Password gate (auth.gate). Fail → exit 2
  3. Update check (no-op when running from source)
  4. CLI arg dispatch:
       --config         → run wizard, exit
       --run [args...]  → run pipeline (args passthrough to run_excel.main), exit
       (no args)        → console menu loop
  5. Pause before close so the user can read the console
"""
import os
import sys
import traceback
from pathlib import Path

# Force UTF-8 console (Anaconda CPython on Chinese Windows sometimes defaults
# to cp936; the .cmd shim and Inno-Setup-spawned shell both should already
# be UTF-8 but be defensive).
try:
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    sys.stderr.reconfigure(encoding="utf-8", errors="replace")
    sys.stdin.reconfigure(encoding="utf-8", errors="replace")
except (AttributeError, ValueError):
    pass


def _pause_before_exit():
    try:
        print("\n" + "=" * 60)
        input("按 Enter 关闭窗口...")
    except (EOFError, KeyboardInterrupt):
        pass


# ── Pure helpers ─────────────────────────────────────────────────────────────


def has_valid_config(config_file: Path | None = None) -> bool:
    """True if config.env exists and contains a non-empty SHEIN_SUBMITTED_DIR."""
    from setup_wizard import is_first_run_complete
    return is_first_run_complete(config_file)


def read_menu_choice(input_fn=input) -> str:
    """Prompt for menu choice, return "1" / "2" / "Q". Re-prompts on invalid input.

    Empty input is treated as "Q" (user pressed Enter on the menu).
    """
    prompt = (
        "\n请选择:\n"
        "  [1] 配置目标路径\n"
        "  [2] 直接跑\n"
        "  [Q] 退出\n"
        "> "
    )
    while True:
        raw = (input_fn(prompt) or "").strip()
        if raw == "":
            return "Q"
        c = raw.upper()
        if c in ("1", "2", "Q"):
            return c
        print(f"  无效输入 “{raw}”;请输入 1 / 2 / Q。")


# ── Action handlers ──────────────────────────────────────────────────────────


def _action_config() -> int:
    from setup_wizard import run_wizard
    print("[配置] 打开设置向导...")
    ok = run_wizard()
    if ok:
        print("[配置] 已保存。")
        return 0
    print("[配置] 已取消(未保存)。")
    return 1


def _action_run() -> int:
    if not has_valid_config():
        print("[运行] 尚未配置;请先选 [1] 配置目标路径,或用 --config 跑配置向导。")
        return 1

    # Reload config now that wizard may have written config.env
    for mod_name in list(sys.modules.keys()):
        if mod_name == "config" or mod_name.startswith("config."):
            del sys.modules[mod_name]

    print()
    print("=" * 60)
    print("开始抓取...")
    print("=" * 60)
    print()
    from run_excel import main as run_excel_main
    run_excel_main()
    return 0


# ── Entry ────────────────────────────────────────────────────────────────────


def main():
    try:
        from version import VERSION
        print(f"SHEIN 上架工具 AU  v{VERSION}")
        print("=" * 60)

        # 0. Password gate
        from auth import gate
        if not gate():
            return 2

        # 1. Update check (no-op from source)
        try:
            from update_check import check_for_update
            check_for_update()
        except Exception as e:
            print(f"[更新检查] 跳过({e.__class__.__name__})")

        # 2. CLI dispatch — strip known flags, pass the rest to run_excel
        argv = sys.argv[1:]
        if "--config" in argv:
            return _action_config()

        if "--run" in argv:
            # Strip --run; rebuild sys.argv so run_excel.main()'s argparse sees the rest
            argv = [a for a in argv if a != "--run"]
            sys.argv = [sys.argv[0]] + argv
            return _action_run()

        # 3. Menu loop — show menu once. After action, exit (pause shows result).
        choice = read_menu_choice()
        if choice == "Q":
            return 0
        if choice == "1":
            return _action_config()
        if choice == "2":
            # In the menu path we don't have extra CLI args for run_excel —
            # just call its main() with the bare argv (file from INPUT_FILENAME).
            sys.argv = [sys.argv[0]]
            return _action_run()

        return 0  # unreachable

    except KeyboardInterrupt:
        print("\n[中断] 用户按 Ctrl+C")
        return 130
    except SystemExit as e:
        return int(e.code) if isinstance(e.code, int) else 1
    except Exception:
        print("\n[严重错误] 发生未处理的异常:")
        traceback.print_exc()
        return 1


if __name__ == "__main__":
    rc = main()
    _pause_before_exit()
    sys.exit(rc)
```

- [ ] **Step 2: Run tests — verify they pass**

```
C:\Users\ak\anaconda3\python.exe test_app_main_menu.py
```
Expected: `ALL PASS`

- [ ] **Step 3: Compile-check**

```
C:\Users\ak\anaconda3\python.exe -m py_compile app_main.py
```
Expected: silent.

- [ ] **Step 4: Smoke test — run from source, pick Q immediately**

```
echo Q | C:\Users\ak\anaconda3\python.exe app_main.py
```
Expected: banner prints; password prompt may appear (auth.py needs network to fetch `auth.json` from GitHub — see note below); after Q, exits. If `auth.json` isn't yet pushed, gate will fail offline — that's expected at this stage.

**Note:** Full app_main smoke test requires `auth.json` to be reachable on GitHub raw. This becomes possible only after `git push origin main`. For now, accept that smoke testing from source is limited to compile + unit tests until release time.

- [ ] **Step 5: Commit**

```bash
git add app_main.py test_app_main_menu.py
git commit -m "feat(installer): app_main.py — menu-driven entry with --config/--run bypass"
```

---

## Task 10: All-suite test run + push so `auth.json` is reachable

**Files:** (no new files)

- [ ] **Step 1: Run all tests in repo**

```
C:\Users\ak\anaconda3\python.exe test_variant_merge.py
C:\Users\ak\anaconda3\python.exe test_setup_wizard.py
C:\Users\ak\anaconda3\python.exe test_app_main_menu.py
```
Expected: each prints `ALL PASS`.

- [ ] **Step 2: Push commits so `auth.json` is reachable from `raw.githubusercontent.com`**

```bash
git push origin main
```

- [ ] **Step 3: Verify auth.json is fetchable**

```
C:\Users\ak\anaconda3\python.exe -c "
import requests
url = 'https://raw.githubusercontent.com/MikeLiu93/shein-extract-au/main/auth.json'
r = requests.get(url, timeout=10)
print('status:', r.status_code)
print('body:', r.text[:200])
"
```
Expected: `status: 200`, body contains `active_passwords`.

- [ ] **Step 4: Full from-source smoke test of `app_main`**

```
C:\Users\ak\anaconda3\python.exe app_main.py
```
Expected:
- Banner: `SHEIN 上架工具 AU  v0.1.0`
- Password prompt; type `Ace2025`, press Enter
- `[验证] 通过 ✓`
- Menu prompt with [1] [2] [Q]; type `Q` → exits with `按 Enter 关闭窗口...`

If anything misbehaves, debug before proceeding.

---

## Task 11: Create `pyinstaller.spec`

**Files:**
- Create: `pyinstaller.spec`

- [ ] **Step 1: Write the file**

Based on 母项目's spec, with these differences:
- Entry: `app_main.py` (same)
- Name: `SheinExtractAU` (not `SheinExtract`)
- Hidden imports drop `key_store`, `check_stock`, `merge_master` (AU doesn't have them)
- Add `auth.json` as a bundled data file (so the binary can read its own copy if needed; actually auth.py fetches live URL so optional — but include `auth.json` next to exe via Inno Setup if you ever want to ship a default snapshot. **Skip from datas for now;** rely entirely on the live URL fetch.)

```python
# -*- mode: python ; coding: utf-8 -*-
"""
PyInstaller spec for SheinExtractAU — produces dist/SheinExtractAU.exe.

Build with:
    pyinstaller pyinstaller.spec --clean --noconfirm

Outputs a single console-mode .exe (employees see a console with progress).
"""

block_cipher = None

a = Analysis(
    ['app_main.py'],
    pathex=['.'],
    binaries=[],
    datas=[],
    hiddenimports=[
        # openpyxl pulls these dynamically
        'openpyxl.styles.alignment',
        'openpyxl.styles.borders',
        'openpyxl.styles.fills',
        'openpyxl.styles.fonts',
        'openpyxl.drawing.image',
        # PIL via openpyxl image support
        'PIL',
        'PIL.Image',
        # Tkinter (wizard + update dialog)
        'tkinter',
        'tkinter.ttk',
        'tkinter.messagebox',
        'tkinter.filedialog',
        # websocket-client (used by Chrome CDP code in shein_scraper)
        'websocket._abnf',
        'websocket._app',
        'websocket._core',
        'websocket._exceptions',
        'websocket._handshake',
        'websocket._http',
        'websocket._logging',
        'websocket._socket',
        'websocket._ssl_compat',
        'websocket._url',
        'websocket._utils',
        # our own modules
        'config',
        'shein_scraper',
        'run_excel',
        'notify',
        'setup_wizard',
        'update_check',
        'auth',
        'version',
    ],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[
        'matplotlib',
        'numpy',
        'pandas',
        'scipy',
        'IPython',
        'jupyter',
        'notebook',
        'sphinx',
        'pytest',
    ],
    win_no_prefer_redirects=False,
    win_private_assemblies=False,
    cipher=block_cipher,
    noarchive=False,
)

pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.zipfiles,
    a.datas,
    [],
    name='SheinExtractAU',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=False,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=True,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    # icon='resources\\app.ico',  # add later if a .ico is provided
)
```

- [ ] **Step 2: Commit**

```bash
git add pyinstaller.spec
git commit -m "build(installer): pyinstaller.spec for SheinExtractAU"
```

---

## Task 12: Create `build.bat`

**Files:**
- Create: `build.bat`

Differences from 母项目:
- No `.build_key.txt` / `make_key_store.py` step (AU has no obfuscated key)
- Output exe name `SheinExtractAU.exe`
- Output installer pattern `SheinExtractAU-Setup-{version}.exe`

- [ ] **Step 1: Write the file**

```bat
@echo off
REM ============================================================
REM  Owner-only build script. NOT for employees.
REM  Produces:
REM    dist\SheinExtractAU.exe              (PyInstaller output)
REM    dist\SheinExtractAU-Setup-X.Y.Z.exe  (Inno Setup output)
REM
REM Prerequisites (one-time):
REM   1. Python 3.10+ on PATH (Anaconda fine)
REM   2. pip install pyinstaller
REM   3. pip install -r requirements.txt
REM   4. Inno Setup 6.x installed (so iscc.exe is on PATH or at
REM      C:\Program Files (x86)\Inno Setup 6\)
REM
REM Usage:
REM   build.bat        full build + installer
REM   build.bat exe    only PyInstaller, skip Inno Setup
REM ============================================================

setlocal enabledelayedexpansion
chcp 65001 >nul
cd /d "%~dp0"

REM Anaconda ships tk/tcl/openssl DLLs in Library\bin. PyInstaller walks
REM dependencies via PATH — without this, DLLs aren't found and the EXE
REM crashes at runtime (e.g. ImportError: _tkinter).
for /f "delims=" %%P in ('python -c "import sys,os; print(os.path.join(sys.prefix,'Library','bin'))" 2^>nul') do set "CONDA_LIBBIN=%%P"
if defined CONDA_LIBBIN if exist "%CONDA_LIBBIN%" (
    echo [build] Prepending conda Library\bin to PATH: %CONDA_LIBBIN%
    set "PATH=%CONDA_LIBBIN%;%PATH%"
)

echo.
echo ============================================================
echo  Step 1/2: PyInstaller - build SheinExtractAU.exe
echo ============================================================
if exist build rmdir /s /q build
if exist dist\SheinExtractAU.exe del /q dist\SheinExtractAU.exe
pyinstaller pyinstaller.spec --clean --noconfirm
if errorlevel 1 (
    echo [ERROR] PyInstaller failed.
    exit /b 1
)

if /i "%1"=="exe" (
    echo.
    echo Skipping Inno Setup (--exe-only). dist\SheinExtractAU.exe is ready.
    exit /b 0
)

echo.
echo ============================================================
echo  Step 2/2: Inno Setup - wrap into installer
echo ============================================================
where iscc >nul 2>nul
if errorlevel 1 (
    if exist "C:\Program Files (x86)\Inno Setup 6\iscc.exe" (
        set "ISCC=C:\Program Files (x86)\Inno Setup 6\iscc.exe"
    ) else (
        echo [ERROR] iscc.exe not found. Install Inno Setup 6 from
        echo https://jrsoftware.org/isdl.php
        exit /b 1
    )
) else (
    set "ISCC=iscc"
)
"%ISCC%" installer.iss
if errorlevel 1 (
    echo [ERROR] Inno Setup compilation failed.
    exit /b 1
)

echo.
echo ============================================================
echo  Build complete!
echo ============================================================
echo Installer: dist\SheinExtractAU-Setup-*.exe
echo Next: test locally, then `release.bat X.Y.Z` to publish.
echo ============================================================
```

- [ ] **Step 2: First build attempt — PyInstaller only**

```
build.bat exe
```
Expected: prints progress; ends with "Skipping Inno Setup..."; produces `dist\SheinExtractAU.exe` (~50-80 MB).

If it fails: read the PyInstaller log for missing modules and add to `pyinstaller.spec`'s `hiddenimports`. Re-run.

- [ ] **Step 3: Run the built exe directly**

```
dist\SheinExtractAU.exe
```
Expected: banner, password prompt, accepts `Ace2025`, shows menu. Type `Q` → exits.

- [ ] **Step 4: Commit**

```bash
git add build.bat
git commit -m "build(installer): build.bat - PyInstaller + Inno Setup chain"
```

---

## Task 13: Create `installer.iss` and build the installer

**Files:**
- Create: `installer.iss`

- [ ] **Step 1: Generate a fresh AppId GUID**

```
C:\Users\ak\anaconda3\python.exe -c "import uuid; print(str(uuid.uuid4()).upper())"
```
Record the GUID (format `XXXXXXXX-XXXX-XXXX-XXXX-XXXXXXXXXXXX`).

- [ ] **Step 2: Write `installer.iss`**

Replace `<NEW-GUID-HERE>` with the GUID from Step 1.

```ini
; Inno Setup script for SheinExtractAU.
; Compile with:  iscc installer.iss
; Produces: dist\SheinExtractAU-Setup-{version}.exe

#define MyAppName "SHEIN 上架工具 AU"
#define MyAppNameAscii "SheinExtractAU"
#define MyAppVersion "0.1.0"          ; Keep in sync with version.py
#define MyAppPublisher "MikeLiu93"
#define MyAppURL "https://github.com/MikeLiu93/shein-extract-au"
#define MyAppExeName "SheinExtractAU.exe"

[Setup]
; AppId is a unique GUID identifying this app — do NOT change between versions.
; Distinct from the 母项目 SheinExtract AppId so they coexist.
AppId={{<NEW-GUID-HERE>}
AppName={#MyAppName}
AppVersion={#MyAppVersion}
AppVerName={#MyAppName} {#MyAppVersion}
AppPublisher={#MyAppPublisher}
AppPublisherURL={#MyAppURL}
AppSupportURL={#MyAppURL}
AppUpdatesURL={#MyAppURL}/releases

DefaultDirName={localappdata}\{#MyAppNameAscii}
DefaultGroupName={#MyAppName}
DisableProgramGroupPage=yes
PrivilegesRequired=lowest

OutputDir=dist
OutputBaseFilename={#MyAppNameAscii}-Setup-{#MyAppVersion}
Compression=lzma2/max
SolidCompression=yes
WizardStyle=modern

UsePreviousAppDir=yes
UsePreviousGroup=yes

ShowLanguageDialog=no

[Languages]
Name: "chinesesimplified"; MessagesFile: "compiler:Default.isl"

[Tasks]
Name: "desktopicon"; Description: "在桌面创建快捷方式"; GroupDescription: "附加任务:"; Flags: unchecked
Name: "startmenuicon"; Description: "在开始菜单创建快捷方式"; GroupDescription: "附加任务:"

[Files]
Source: "dist\{#MyAppExeName}"; DestDir: "{app}"; Flags: ignoreversion

[Icons]
; SINGLE program icon — no separate "配置" icon. Menu handles that in-app.
Name: "{group}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"; Tasks: startmenuicon
Name: "{group}\卸载 {#MyAppName}"; Filename: "{uninstallexe}"; Tasks: startmenuicon
Name: "{userdesktop}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"; Tasks: desktopicon

[Run]
Filename: "{app}\{#MyAppExeName}"; Description: "立即启动 {#MyAppName}"; Flags: nowait postinstall skipifsilent

[UninstallDelete]
; Preserve user data: %APPDATA%\shein-extract-au\, %USERPROFILE%\shein-cdp-profile-au\.
; Only files we drop in {app} get cleaned automatically.
```

- [ ] **Step 3: Full build (PyInstaller + Inno Setup)**

```
build.bat
```
Expected: produces `dist\SheinExtractAU-Setup-0.1.0.exe` (~20-30 MB compressed).

- [ ] **Step 4: Commit**

```bash
git add installer.iss
git commit -m "build(installer): installer.iss - 1-icon Inno Setup script"
```

---

## Task 14: Manual smoke test of installer

**Files:** (no new files)

- [ ] **Step 1: Install via the setup**

Double-click `dist\SheinExtractAU-Setup-0.1.0.exe`. Click through with defaults. At the end check "立即启动".

- [ ] **Step 2: Verify menu launch**

The installed app should:
- Show password prompt → accept `Ace2025`
- Show menu [1] [2] [Q]
- [1] → wizard opens (Tkinter), fill paths/key, save
- Re-launch from Start Menu / desktop icon → menu shows → [2] → pipeline runs (with the LU file currently set by .env? No — the installed app uses `config.env` only, not the dev `.env`. The wizard must have set `SHEIN_INPUT_FILENAME=希音链接 - LU.xlsx` or left it blank.)

If any step misbehaves, debug and rebuild.

- [ ] **Step 3: Verify coexistence with 母项目 (if 母项目 is installed)**

- Both shortcuts in Start Menu / desktop work independently
- Both pass their respective password gates
- Output directories don't collide

- [ ] **Step 4: Verify uninstall preserves user data**

- Uninstall AU from Settings → Apps
- `%APPDATA%\shein-extract-au\config.env` still exists ✓
- `%USERPROFILE%\shein-cdp-profile-au\` still exists ✓

- [ ] **Step 5: No commit — manual test only**

If you found and fixed bugs, those fixes are separate commits in their respective tasks.

---

## Task 15: Write `INSTALL_GUIDE_CN.md`

**Files:**
- Create: `INSTALL_GUIDE_CN.md`

- [ ] **Step 1: Write the file**

```markdown
# SHEIN 上架工具 AU — 安装指南

## 安装

1. 下载最新版安装包：https://github.com/MikeLiu93/shein-extract-au/releases/latest
2. 双击 `SheinExtractAU-Setup-X.Y.Z.exe`
3. 按提示安装(默认安装到 `%LOCALAPPDATA%\SheinExtractAU\`,**不需要管理员权限**)
4. 装好后开始菜单 / 桌面会出现 **SHEIN 上架工具 AU** 一个图标

## 第一次启动

1. 双击图标
2. 输入访问密码(找管理员要)
3. 看到菜单:
   ```
   请选择:
     [1] 配置目标路径
     [2] 直接跑
     [Q] 退出
   ```
4. **首次必须先选 [1]** 把目录配好,会弹出配置向导(Tkinter 窗口)
5. 向导里填:
   - 输入表所在目录(共享盘里那个 `澳洲站` 文件夹)
   - 输出根目录(留空 = 自动用 输入目录\上架资料-已完成)
   - 指定输入文件名(可选;留空 = 处理目录下所有 .xlsx)
   - Anthropic API key(找管理员要;留空 = 不启用 AI 标题)
6. 保存 → 关闭向导 → 程序自动退出

## 之后每次启动

1. 双击图标 → 输密码 → 看到菜单
2. 选 [2] 直接跑
3. 弹出 Chrome 窗口(独立的 profile,**第一次需要登录 SHEIN 卖家账号**——以后自动)
4. 工具开始抓取,console 实时显示进度
5. 跑完会停在 "按 Enter 关闭窗口...",看一眼结果再关

## 修改配置

任何时候双击图标 → 选 [1] 可以重新打开配置向导改路径/换文件/换 API key。

## 卸载

- 控制面板 → 程序和功能 → 找 "SHEIN 上架工具 AU" → 卸载
- **不会删除**:配置文件(`%APPDATA%\shein-extract-au\`)、SHEIN 登录信息(`%USERPROFILE%\shein-cdp-profile-au\`)。完全干净卸载需要你手动删这两个文件夹

## 常见问题

- **未找到 Chrome**:装 https://www.google.com/chrome/ 再开向导
- **密码错**:每个版本最多 3 次,超了直接退;再打开重试
- **Chrome 卡在登录页**:你手动登一次,工具会等;之后不用再登
- **运行报错"输入目录不存在"**:检查 Google Drive 是否同步完成,或重新选 [1] 改路径

## 与"SHEIN 上架工具"(美国站)共存

两个工具可以同时装、同时用,互不影响:
- 不同安装目录、不同密码、不同 Chrome profile、不同 Chrome 端口
- 共享盘里的输入/输出文件夹是分开配的
```

- [ ] **Step 2: Commit**

```bash
git add INSTALL_GUIDE_CN.md
git commit -m "docs: AU install guide"
```

---

## Task 16: Write `release.bat`

**Files:**
- Create: `release.bat`

- [ ] **Step 1: Write the file**

```bat
@echo off
REM ============================================================
REM  Semi-automated release.
REM
REM Usage:
REM   release.bat 0.1.1
REM
REM Does:
REM   1. Verify clean working tree + on main + in sync with origin
REM   2. Update version.py to the given version
REM   3. Commit + push the version bump
REM   4. Run build.bat (PyInstaller + Inno Setup)
REM   5. Verify dist\SheinExtractAU-Setup-X.Y.Z.exe exists
REM   6. git tag v{version} + git push origin v{version}
REM   7. Open the GitHub "Draft a new release" page in your browser
REM      with the tag pre-selected. You manually drag the setup
REM      .exe into Assets, write release notes, click Publish.
REM ============================================================

setlocal enabledelayedexpansion
chcp 65001 >nul
cd /d "%~dp0"

if "%1"=="" (
    echo Usage: release.bat X.Y.Z
    exit /b 1
)
set VERSION=%1

echo.
echo ============================================================
echo  Step 1/7: Pre-flight checks
echo ============================================================

git rev-parse --abbrev-ref HEAD > %TEMP%\branch.txt
set /p BRANCH=<%TEMP%\branch.txt
del %TEMP%\branch.txt
if not "%BRANCH%"=="main" (
    echo [ERROR] Not on main branch ^(on %BRANCH%^). Aborting.
    exit /b 1
)

git status --porcelain > %TEMP%\status.txt
for %%A in (%TEMP%\status.txt) do if not %%~zA==0 (
    echo [ERROR] Working tree not clean. Commit or stash first.
    del %TEMP%\status.txt
    exit /b 1
)
del %TEMP%\status.txt

git fetch origin main >nul 2>&1
git rev-list HEAD..origin/main --count > %TEMP%\behind.txt
set /p BEHIND=<%TEMP%\behind.txt
del %TEMP%\behind.txt
if not "%BEHIND%"=="0" (
    echo [ERROR] Local main is %BEHIND% commit^(s^) behind origin/main. Pull first.
    exit /b 1
)

echo OK: clean main, in sync.

echo.
echo ============================================================
echo  Step 2/7: Bump version.py to %VERSION%
echo ============================================================

python -c "import pathlib,re; p=pathlib.Path('version.py'); s=p.read_text(encoding='utf-8'); s2=re.sub(r'VERSION\s*=\s*\".*?\"', 'VERSION = \"%VERSION%\"', s); p.write_text(s2, encoding='utf-8'); print(s2.strip())"
if errorlevel 1 (
    echo [ERROR] Failed to bump version.py
    exit /b 1
)

echo.
echo ============================================================
echo  Step 3/7: Bump installer.iss MyAppVersion to %VERSION%
echo ============================================================

python -c "import pathlib,re; p=pathlib.Path('installer.iss'); s=p.read_text(encoding='utf-8'); s2=re.sub(r'#define MyAppVersion \"[^\"]*\"', '#define MyAppVersion \"%VERSION%\"', s); p.write_text(s2, encoding='utf-8'); print('installer.iss bumped')"
if errorlevel 1 (
    echo [ERROR] Failed to bump installer.iss
    exit /b 1
)

echo.
echo ============================================================
echo  Step 4/7: Commit + push version bump
echo ============================================================

git add version.py installer.iss
git commit -m "release: v%VERSION%"
git push origin main
if errorlevel 1 (
    echo [ERROR] git push failed.
    exit /b 1
)

echo.
echo ============================================================
echo  Step 5/7: Build (PyInstaller + Inno Setup)
echo ============================================================

call build.bat
if errorlevel 1 (
    echo [ERROR] build.bat failed.
    exit /b 1
)

if not exist "dist\SheinExtractAU-Setup-%VERSION%.exe" (
    echo [ERROR] Expected dist\SheinExtractAU-Setup-%VERSION%.exe but it's missing.
    exit /b 1
)

echo.
echo ============================================================
echo  Step 6/7: Tag + push tag
echo ============================================================

git tag v%VERSION%
git push origin v%VERSION%
if errorlevel 1 (
    echo [ERROR] tag push failed.
    exit /b 1
)

echo.
echo ============================================================
echo  Step 7/7: Open GitHub "Draft a new release" page
echo ============================================================
echo.
echo Drag this file into the Assets section, write release notes,
echo then click "Publish release":
echo.
echo   dist\SheinExtractAU-Setup-%VERSION%.exe
echo.

start "" "https://github.com/MikeLiu93/shein-extract-au/releases/new?tag=v%VERSION%"

echo Done. Browser opened. Finish the release in the browser.
```

- [ ] **Step 2: Verify the script doesn't have syntax errors (dry interactive read)**

```
type release.bat
```
Just visually skim — look for unbalanced quotes, missing `endlocal`, etc. Don't run it now (it would actually try to release).

- [ ] **Step 3: Commit**

```bash
git add release.bat
git commit -m "build(installer): release.bat - semi-automated release flow"
```

---

## Task 17: First release — v0.1.0 via `release.bat`

**Files:** (no new files; this task uses `release.bat`)

- [ ] **Step 1: Verify working tree clean + on main + in sync**

```bash
git status -sb
```
Expected: `## main...origin/main` and no other lines.

- [ ] **Step 2: Run release.bat at the same version it already is (0.1.0)**

Since `version.py` already says `0.1.0`, running `release.bat 0.1.0` will:
- Try to bump version.py to `0.1.0` (no-op — same value)
- Try to commit (no changes → `git commit` will fail because nothing to commit)

So for the FIRST release, skip release.bat and do steps manually:

```bash
# Build
build.bat

# Tag and push
git tag v0.1.0
git push origin v0.1.0
```

- [ ] **Step 3: Open the GitHub release draft page**

Visit: `https://github.com/MikeLiu93/shein-extract-au/releases/new?tag=v0.1.0`

Fill in:
- Title: `v0.1.0 — first installable release`
- Notes: brief summary (initial AU installer, single-icon menu UX, mirrors 母项目 feature set)
- Drag `dist\SheinExtractAU-Setup-0.1.0.exe` into the Assets area
- Click **Publish release**

- [ ] **Step 4: Verify the release is live**

```
C:\Users\ak\anaconda3\python.exe -c "
import requests
r = requests.get('https://api.github.com/repos/MikeLiu93/shein-extract-au/releases/latest', timeout=10)
print('status:', r.status_code)
j = r.json()
print('tag:', j.get('tag_name'))
print('assets:', [a.get('name') for a in j.get('assets', [])])
"
```
Expected: `tag: v0.1.0`, assets list includes `SheinExtractAU-Setup-0.1.0.exe`.

- [ ] **Step 5: No commit — release is done; subsequent bumps use `release.bat`**

---

## Self-Review Checklist (done by me before handoff)

**Spec coverage** (each Spec section → task):
- §3 (naming/identifiers) → Tasks 1, 3, 4, 11, 13
- §4 (startup flow / CLI bypass) → Task 9
- §5 (component list) → Tasks 1-13
- §6 (wizard fields) → Task 7
- §7 (installer.iss skeleton) → Task 13
- §8 (coexistence) → Tasks 3, 4, 5, 13 (separate AppId/profile/port/data dirs all set)
- §9 (error handling) → Task 9 (app_main wraps everything in try/except)
- §10 (testing strategy) → Tasks 6, 8 (unit tests) + Task 14 (manual smoke)
- §11 (decisions) → encoded in Task 2 (`Ace2025`), Task 9 (passthrough), Task 13 (no .ico)
- §14 (release flow) → Tasks 16, 17

**Placeholder scan**: No "TBD"/"TODO"/"implement later". `<NEW-GUID-HERE>` in Task 13 has explicit generation step right above it.

**Type/name consistency**:
- `read_menu_choice`, `has_valid_config`, `write_config_env`, `load_existing_config`, `is_first_run_complete`, `mask_api_key`, `find_chrome` — used identically across tasks ✓
- `VERSION` value `0.1.0` consistent across version.py, installer.iss, release commit
- Chrome profile dir `shein-cdp-profile-au` consistent in Task 5 + Spec §8
- CDP port `9223` consistent ✓
