import logging
import queue
from pathlib import Path
import datetime
from typing import Optional



class UIHandler(logging.Handler):
    """A custom logging handler that pushes records to a queue for the UI thread."""

    def __init__(self, log_queue: queue.Queue):
        super().__init__()
        self.log_queue = log_queue

    def emit(self, record):
        # We format the level for the UI pump used in the original main.py
        level_map = {
            logging.INFO: "info",
            logging.WARNING: "warn",
            logging.ERROR: "err",
            logging.CRITICAL: "err",
        }
        level = level_map.get(record.levelno, "info")
        self.log_queue.put((level, self.format(record)))


def setup_logging(log_queue: Optional[queue.Queue] = None):
    """Configures the root logger for file and optional UI output."""
    logger = logging.getLogger()
    logger.setLevel(logging.INFO)

    formatter = logging.Formatter("%(asctime)s - %(levelname)s - %(message)s")

    # File logging
    log_dir = Path("logs")
    log_dir.mkdir(exist_ok=True)
    file_handler = logging.FileHandler(
        log_dir / f"app_{datetime.date.today()}.log", encoding="utf-8"
    )
    file_handler.setFormatter(formatter)
    logger.addHandler(file_handler)

    # UI logging via queue
    if log_queue:
        ui_handler = UIHandler(log_queue)
        ui_handler.setFormatter(logging.Formatter("%(message)s"))
        logger.addHandler(ui_handler)
