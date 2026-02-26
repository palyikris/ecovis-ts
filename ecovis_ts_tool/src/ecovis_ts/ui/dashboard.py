import ttkbootstrap as tb
from ttkbootstrap.constants import *
from datetime import datetime


class StatCard(tb.Labelframe):
    """A small reusable card for displaying status values."""

    def __init__(self, parent, title, **kwargs):
        super().__init__(parent, text=title, padding=12, bootstyle=SECONDARY, **kwargs)
        self.value_label = tb.Label(self, text="—", font=("Segoe UI", 18, "bold"))
        self.value_label.pack(fill=X)
        self.sub_label = tb.Label(
            self, text="", font=("Segoe UI", 9), bootstyle=SECONDARY
        )
        self.sub_label.pack(fill=X, pady=(6, 0))

    def update_stats(self, value, subtitle=""):
        self.value_label.config(text=value)
        self.sub_label.config(text=subtitle)


class DashboardFrame(tb.Frame):
    """The main dashboard container used by app.py."""

    def __init__(self, parent, callback):
        super().__init__(parent, padding=20)
        self.callback = callback  # This is the execute_task function from app.py

        # --- 1. Top Control Bar (Month & Buttons) ---
        ctrl_frame = tb.LabelFrame(self, text="Műveletek", padding=15)
        ctrl_frame.pack(fill=X, pady=(0, 20))

        # Month Selection
        tb.Label(ctrl_frame, text="Hónap:").pack(side=LEFT, padx=(0, 5))
        months = [
            "januar",
            "februar",
            "marcius",
            "aprilis",
            "majus",
            "junius",
            "július",
            "augusztus",
            "szeptember",
            "oktober",
            "november",
            "december",
        ]
        self.month_var = tb.StringVar(value=months[datetime.now().month - 1])
        self.month_combo = tb.Combobox(
            ctrl_frame, values=months, textvariable=self.month_var, width=12
        )
        self.month_combo.pack(side=LEFT, padx=(0, 20))

        # Action Buttons (Calling the callback passed from app.py)
        tb.Button(
            ctrl_frame,
            text="Legördülők Frissítése",
            bootstyle=INFO,
            command=lambda: self.callback("sync"),
        ).pack(side=LEFT, padx=5)

        tb.Button(
            ctrl_frame,
            text="Párellenőrzés",
            bootstyle=WARNING,
            command=lambda: self.callback("validate", self.month_var.get()),
        ).pack(side=LEFT, padx=5)

        tb.Button(
            ctrl_frame,
            text="Összesítés",
            bootstyle=PRIMARY,
            command=lambda: self.callback("aggregate", self.month_var.get()),
        ).pack(side=LEFT, padx=5)

        tb.Button(
            ctrl_frame,
            text="Számlamelléklet",
            bootstyle=SUCCESS,
            command=lambda: self.callback("invoice", self.month_var.get()),
        ).pack(side=LEFT, padx=5)

        # --- 2. Status Cards Area ---
        self.status_card = StatCard(self, title="Rendszer Állapot")
        self.status_card.pack(fill=X, pady=(0, 20))

        # --- 3. Log Output Area ---
        log_frame = tb.LabelFrame(self, text="Eseménynapló", padding=10)
        log_frame.pack(fill=BOTH, expand=YES)

        self.log_text = tb.ScrolledText(
            log_frame, height=10, font=("Consolas", 9), state=DISABLED
        )
        self.log_text.pack(fill=BOTH, expand=YES)

    def add_log(self, level, message):
        """This matches the level/msg format sent by app.py _pump_logs."""
        timestamp = datetime.now().strftime("%H:%M:%S")
        full_msg = f"[{timestamp}] {level.upper()}: {message}\n"

        self.log_text.config(state=NORMAL)
        self.log_text.insert(END, full_msg)
        self.log_text.see(END)
        self.log_text.config(state=DISABLED)

    def update(self):
        """Called by app.py's execute_task after a thread finishes."""
        self.status_card.update_stats(
            "Kész", f"Utolsó futtatás: {datetime.now().strftime('%H:%M:%S')}"
        )
