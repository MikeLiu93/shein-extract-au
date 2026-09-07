"""
Path + env config for the Shein extract pipeline (澳洲站分支). Reads .env on import.

设计原则
========
所有"用户决定"的路径和数值都通过 .env / config.env 配置：

    SHEIN_SUBMITTED_DIR  — 输入 Excel 所在目录（共享盘上）
    SHEIN_INPUT_FILENAME — 输入 Excel 文件名（可选；不设则处理目录下所有 .xlsx）
    SHEIN_OUTPUT_DIR     — 输出根目录（per-seq 媒体落这里，共享盘上）
    SHEIN_BACKUP_DIR     — 备份根目录（每次跑前把输入表复制一份到这里）
    SHEIN_EBAY_MARKUP    — eBay 定价系数（默认 1.2；员工在向导里改）
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

# 备份根目录：每次跑前把输入表复制一份到这里；空字符串 = 用 <SUBMITTED_DIR>/_backups
_backup_env = os.environ.get("SHEIN_BACKUP_DIR", "").strip()
BACKUP_DIR = Path(_backup_env) if _backup_env else SUBMITTED_DIR / "_backups"

# ── eBay 定价 & 运费 ─────────────────────────────────────────────────────────
# eBay 价格分档公式：
#   price <  LOW_PRICE_THRESHOLD  →  price + LOW_PRICE_FLAT_MARKUP + shipping
#   price >= LOW_PRICE_THRESHOLD  →  price × EBAY_MARKUP           + shipping
# 低价档用固定加价（$10）而不是乘系数——$5 商品乘 1.2 只赚 $1，覆盖不了 eBay
# 上架费。高价档用系数（员工在向导里选 1.2/1.5/2.0）。$20 是阈值，员工不改。
EBAY_MARKUP            = float(os.environ.get("SHEIN_EBAY_MARKUP", "1.2"))
LOW_PRICE_THRESHOLD    = float(os.environ.get("SHEIN_AU_LOW_PRICE_THRESHOLD", "20.0"))
LOW_PRICE_FLAT_MARKUP  = float(os.environ.get("SHEIN_AU_LOW_PRICE_MARKUP", "10.0"))

# 澳洲站 Standard shipping：价格 ≥ FREE_SHIPPING_THRESHOLD 免运，否则收
# DEFAULT_SHIPPING_FEE。表格 D 列的手动运费在 run_excel.py 一层做覆盖。
DEFAULT_SHIPPING_FEE    = float(os.environ.get("SHEIN_AU_SHIPPING_FEE", "7.95"))
FREE_SHIPPING_THRESHOLD = float(os.environ.get("SHEIN_AU_FREE_SHIP_MIN", "9.0"))
