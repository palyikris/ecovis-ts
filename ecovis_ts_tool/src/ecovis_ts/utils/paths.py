import os
import sys
import subprocess
from pathlib import Path
from typing import Optional

# Note: SETTINGS will eventually be moved to src/ecovis_ts/config.py
from ecovis_ts.config import SETTINGS


def ts_root() -> Path:
    """Returns the root folder containing the timesheet Excel files."""
    p = Path(SETTINGS.get("ts_folder") or Path.cwd())
    return p if p.exists() else Path.cwd()


def output_root() -> Path:
    """Returns the destination folder for generated reports."""
    p = Path(SETTINGS.get("output_folder") or "")
    if not p:
        return ts_root()
    return p if p.exists() else ts_root()


def reports_root() -> Path:
    """Ensures and returns the 'reports' subdirectory within the output folder."""
    r = output_root() / "reports"
    r.mkdir(parents=True, exist_ok=True)
    return r


def backup_root() -> Optional[Path]:
    """Returns the backup folder path if backups are enabled."""
    if not SETTINGS.get("backup_enabled"):
        return None
    p = Path(SETTINGS.get("backup_folder") or "")
    return p if p.exists() else None


def open_path(path: Path):
    """Opens a directory or file using the system's default application."""
    try:
        if sys.platform.startswith("win"):
            os.startfile(path)
        elif sys.platform == "darwin":
            subprocess.Popen(["open", str(path)])
        else:
            subprocess.Popen(["xdg-open", str(path)])
    except Exception as e:
        # This will be replaced by a proper logging call later
        print(f"Error opening path {path}: {e}")


def open_file(path: Path):
    """Safely opens a file if it exists."""
    if not path:
        return
    p = path if path.is_absolute() else (Path.cwd() / path)
    if not p.exists():
        print(f"File not found: {p}")
        return
    open_path(p)
