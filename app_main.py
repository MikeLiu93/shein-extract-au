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
