"""Central configuration. Secrets come from environment variables only."""
import os
from pathlib import Path

BASE_DIR = Path(__file__).resolve().parent

# Minimal .env loader (no extra dependency). Real env vars take priority.
_env = BASE_DIR / ".env"
if _env.exists():
    for line in _env.read_text(encoding="utf-8").splitlines():
        if "=" in line and not line.lstrip().startswith("#"):
            k, v = line.split("=", 1)
            os.environ.setdefault(k.strip(), v.strip())
DATA_DIR = BASE_DIR / "sample_data"
OUTPUT_DIR = BASE_DIR / "output"

# Sample workbook produced by generate_sample_data.py
WORKBOOK = DATA_DIR / "Long_Outage_Summary_SAMPLE.xlsx"

# Tables to snapshot: operator -> (sheet name, title shown on the image)
TABLES = {
    "OperatorA": ("OperatorA RCA", "OperatorA Long Outage RCA With Details"),
    "OperatorB": ("OperatorB RCA", "OperatorB Long Outage RCA With Details"),
}

# Delivery. DRY_RUN=1 (default) only logs what would be sent.
DRY_RUN = os.getenv("DRY_RUN", "1") == "1"
TELEGRAM_BOT_TOKEN = os.getenv("TELEGRAM_BOT_TOKEN", "")
TELEGRAM_CHAT_ID = os.getenv("TELEGRAM_CHAT_ID", "")
WHATSAPP_GROUP = os.getenv("WHATSAPP_GROUP", "Sample Operations Group")
WHATSAPP_PROFILE_DIR = BASE_DIR / "wa_profile"

# Render settings
DPI = 200
USE_EXCEL = os.getenv("USE_EXCEL", "0") == "1"   # Windows + Excel backend
