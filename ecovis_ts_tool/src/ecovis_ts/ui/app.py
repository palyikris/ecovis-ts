import threading
import logging
import queue
import ttkbootstrap as tb
from ttkbootstrap.constants import *

from ..config import SETTINGS
from ..utils.logging import UIHandler
from ..core import (
    aggregate_timesheets,
    sync_dropdown_lists,
    validate_client_project_pairs,
)
from .dashboard import DashboardFrame
from .settings import SettingsFrame


class EcovisApp(tb.Window):
    def __init__(self):
        theme = SETTINGS.get("theme", "minty")
        super().__init__(title="Ecovis Timesheet Tool", themename=theme)
        self.state("zoomed")

        # Logging Setup
        self.log_queue = queue.Queue()
        self._init_logging()

        # UI Components
        self.notebook = tb.Notebook(self, bootstyle=PRIMARY)
        self.notebook.pack(fill=BOTH, expand=True)

        self.dashboard = DashboardFrame(self.notebook, self.execute_task)
        self.settings = SettingsFrame(self.notebook)

        self.notebook.add(self.dashboard, text="📊 Vezérlőpult")
        self.notebook.add(self.settings, text="⚙️ Beállítások")

        # Start the log pump
        self.after(100, self._pump_logs)

    def _init_logging(self):
        """Redirects root logging to our UI handler."""
        handler = UIHandler(self.log_queue)
        logging.getLogger().addHandler(handler)
        logging.getLogger().setLevel(logging.INFO)

    def execute_task(self, task_type: str, *args):
        """Standard runner that executes core logic in a background thread."""

        def worker():
            try:
                if task_type == "aggregate":
                    aggregate_timesheets(*args)
                elif task_type == "sync":
                    sync_dropdown_lists()
                elif task_type == "validate":
                    validate_client_project_pairs(*args)

                # Update UI stats on completion
                self.after(0, self.dashboard.update)
            except Exception as e:
                logging.error(f"Váratlan hiba: {str(e)}")

        threading.Thread(target=worker, daemon=True).start()

    def _pump_logs(self):
        """Checks the queue for new logs to display in the Dashboard Treeview."""
        try:
            while True:
                level, msg = self.log_queue.get_nowait()
                self.dashboard.add_log(level, msg)
        except queue.Empty:
            pass
        self.after(100, self._pump_logs)


def main():
    app = EcovisApp()
    app.mainloop()
