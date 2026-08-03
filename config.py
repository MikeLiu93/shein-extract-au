"""
Path + env config for the Shein extract pipeline (澳洲站分支). Reads .env on import.

设计原则
========
所有"用户决定"的路径都通过 .env 配置：

    SHEIN_SUBMITTED_DIR  — 输入 Excel 所在目录（共享盘上）
    SHEIN_INPUT_FILENAME — 输入 Excel 文件名（可选；不设则处理目录下所有 .xlsx）
    SHEIN_OUTPUT_DIR     — 输出根目录（共享盘上）
    ANTHROPIC_API_KEY    — Claude Haiku key（AI 标题）

进程环境优先级 > .env 文件 > 默认值。
"""
import os
from pathlib import Path


def _load_env_file() -> None:
    """Load .env into os.environ. Supports project-local .env and
    %APPDATA%\\shein-extract-au\\config.env."""
    candidates = []
    appdata = os.environ.get("APPDATA")
    if appdata:
        candidates.append(Path(appdata) / "shein-extract-au" / "config.env")
    candidates.append(Path(__file__).resolve().parent / ".env")

    for env_path in candidates:
        if not env_path.exists():
            continue
        for raw in env_path.read_text(encoding="utf-8", errors="ignore").splitlines():
            line = raw.strip()
            if not line or line.startswith("#") or "=" not in line:
                continue
            key, val = line.split("=", 1)
            key = key.strip()
            val = val.strip()
            if len(val) >= 2 and val[0] == val[-1] and val[0] in ('"', "'"):
                val = val[1:-1]
            if key and key not in os.environ:
                os.environ[key] = val


_load_env_file()

# ── Australian Shein 默认路径 ─────────────────────────────────────────────────
_AU_BASE = r"D:\共享云端硬盘\02 希音\澳洲站"

SUBMITTED_DIR = Path(os.environ.get("SHEIN_SUBMITTED_DIR", _AU_BASE))

# 留空 / 不设：run_excel.py 会自动处理 SUBMITTED_DIR 顶层下所有 .xlsx 文件。
INPUT_FILENAME = os.environ.get("SHEIN_INPUT_FILENAME", "").strip()

OUTPUT_ROOT_2ND = Path(os.environ.get(
    "SHEIN_OUTPUT_DIR",
    _AU_BASE,  # 默认与 SUBMITTED_DIR 同级；店铺名自动作为下一级子文件夹
))
