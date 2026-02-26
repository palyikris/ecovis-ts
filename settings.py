# settings.py
# -*- coding: utf-8 -*-
from __future__ import annotations
from pathlib import Path
import json

try:
    # Csak a hibaüzenetekhez – ha nincs GUI, simán elnyeljük.
    from tkinter import messagebox  # type: ignore
except Exception:  # pragma: no cover
    messagebox = None  # type: ignore

# Publikus konstansok
CONFIG_PATH = Path("settings.json")

# Alapértelmezett beállítások (változatlanul átemelve)
DEFAULT_SETTINGS = {
    "theme": "minty",
    "ui_scale_pct": 100,
    "language": "hu",
    "ts_folder": str(Path.cwd()),
    "output_folder": "",  # üres => TS mappa
    "backup_enabled": False,
    "backup_folder": "",
    "default_client_codes": "",  # "ABC123,XYZ987"
    "remember_last_selection": True,
    "auto_open_output_on_success": False,
    "auto_open_details_on_error": True,
    "popup_autoclose_sec": 0,  # 0 => nem zárja automatikusan
    "sound_enabled": True,
    # Napi emlékeztető
    "daily_reminder_enabled": False,
    "daily_reminder_time": "18:00",  # HH:MM
    # Heti riport
    "weekly_report_enabled": False,
    "weekly_report_weekday": 0,  # 0=Hétfő ... 6=Vasárnap
    "weekly_report_time": "08:30",  # HH:MM
    "weekly_report_recipients": "palyizoltan@t-online.hu",
    "email_method": "outlook",  # outlook | smtp
    "smtp_host": "",
    "smtp_port": 587,
    "smtp_tls": True,
    "smtp_user": "",
    "smtp_password": "",
    "last_weekly_report_key": "",  # YYYY-WW hogy ne duplázzon
}


def _merge_defaults(data: dict) -> dict:
    """Hiányzó kulcsok pótlása új verziók esetén."""
    out = dict(DEFAULT_SETTINGS)
    out.update(data or {})
    return out


def load_settings() -> dict:
    """Beállítások betöltése, sérült fájl esetén safe fallback."""
    if CONFIG_PATH.exists():
        try:
            raw = json.loads(CONFIG_PATH.read_text(encoding="utf-8"))
            return _merge_defaults(raw)
        except Exception:
            if messagebox:
                try:
                    messagebox.showwarning(
                        "Beállítások", "A settings.json sérült, visszaállítom alapra."
                    )
                except Exception:
                    pass
    return dict(DEFAULT_SETTINGS)


def save_settings(data: dict) -> None:
    """Beállítások mentése (UTF-8, pretty)."""
    try:
        CONFIG_PATH.write_text(
            json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8"
        )
    except Exception as e:
        if messagebox:
            try:
                messagebox.showerror(
                    "Beállítások mentése", f"Nem sikerült menteni: {e}"
                )
                return
            except Exception:
                pass
        # GUI nélkül: csendes fallback
        print(f"[settings] Mentési hiba: {e!r}")


SETTINGS = load_settings()

__all__ = [
    "CONFIG_PATH",
    "DEFAULT_SETTINGS",
    "SETTINGS",
    "load_settings",
    "save_settings",
]
