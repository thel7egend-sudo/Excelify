import json
from pathlib import Path

DATA_DIR = Path(__file__).parent / "data"
STATE_FILE = DATA_DIR / "app_state.json"

def ensure_storage():
    DATA_DIR.mkdir(exist_ok=True)

def save_state(state: dict):
    ensure_storage()
    with open(STATE_FILE, "w", encoding="utf-8") as f:
        json.dump(state, f, indent=2)

def load_state():
    if not STATE_FILE.exists():
        return None
    with open(STATE_FILE, "r", encoding="utf-8") as f:
        return json.load(f)
