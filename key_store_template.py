"""
TEMPLATE — do not edit directly. The build script `make_key_store.py` reads
.build_key.txt (owner-only, gitignored) and writes the real
`key_store.py` (also gitignored) with the encoded key.

key_store.py is bundled into the .exe by PyInstaller so employees never see
the plaintext key.

Source-run developers can rely on the `.env` fallback in `shein_scraper.py`
instead — set ANTHROPIC_API_KEY there.
"""

ENCODED_KEY = ""
PASSPHRASE = "shein-extract-2026"


# Will be replaced by make_key_store.py with the actual encoded value.
def get_obfuscated_key() -> str:
    return ""
