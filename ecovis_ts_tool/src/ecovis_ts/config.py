import json
from pathlib import Path
from typing import Dict, Any

CONFIG_PATH = Path("settings.json")

DEFAULT_SETTINGS = {
    "theme": "minty",
    "ui_scale_pct": 100,
    "language": "hu",
    "ts_folder": str(Path.cwd()),
    "output_folder": "",
    "backup_enabled": False,
    "default_client_codes": "AIF,BRD,ECL,GRE",  # Based on your current config
    "auto_open_output_on_success": True,
    "auto_open_details_on_error": True,
    "sound_enabled": True,
}


def load_settings() -> Dict[str, Any]:
    if CONFIG_PATH.exists():
        try:
            with open(CONFIG_PATH, "r", encoding="utf-8") as f:
                data = json.load(f)
                return {**DEFAULT_SETTINGS, **data}
        except Exception:
            pass
    return DEFAULT_SETTINGS.copy()


def save_settings(data: Dict[str, Any]):
    with open(CONFIG_PATH, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)


SETTINGS = load_settings()
MONTHS = [
    "Teljes év",
    "januar",
    "februar",
    "marcius",
    "aprilis",
    "majus",
    "junius",
    "julius",
    "augusztus",
    "szeptember",
    "oktober",
    "november",
    "december",
]
