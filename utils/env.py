# utils/env.py
from dotenv import load_dotenv, find_dotenv
from pathlib import Path
import os

ENV_PATH = Path(__file__).resolve().parents[1] / ".env"
load_dotenv(dotenv_path=ENV_PATH, override=True)  # override=True biztosítja a felülírást

def _clean(s): return (s or "").strip()

LOGIN_URL = _clean(os.getenv("LOGIN_URL"))
USERNAME  = _clean(os.getenv("APP_USERNAME"))
PASSWORD  = _clean(os.getenv("APP_PASSWORD"))
HEADLESS  = _clean(os.getenv("HEADLESS") or "true").lower() == "true"

print("DEBUG ENV FILE =", ENV_PATH, "exists:", ENV_PATH.is_file())
print("DEBUG env USERNAME repr =", repr(USERNAME))
