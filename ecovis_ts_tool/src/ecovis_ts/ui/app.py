import tkinter as tk
import ttkbootstrap as tb
from ttkbootstrap.constants import *
from .widgets import DashboardCard
from ecovis_ts.config import MONTHS

class DashboardFrame(tb.Frame):
    def __init__(self, parent, run_command_callback, **kwargs):
        super().__init__(parent, padding=20, **kwargs)
        self.run_command = run_command_callback
        self._setup_ui()

    def _setup_ui(self):
        # Header and Cards (as defined previously)
        self.cards_frame = tb.Frame(self)
        self.cards_frame.pack(fill=X, pady=10)
        # ... (Card setup from previous turn) ...

        # Action Buttons Row
        btns_frame = tb.Frame(self)
        btns_frame.pack(fill=X, pady=20)

        self.month_var = tk.StringVar(value="januar")
        month_menu = tb.Combobox(
            btns_frame, textvariable=self.month_var, values=MONTHS, width=15
        )
        month_menu.pack(side=LEFT, padx=5)

        tb.Button(
            btns_frame,
            text="🔄 Frissítés",
            bootstyle=INFO,
            command=lambda: self.run_command("sync"),
        ).pack(side=LEFT, padx=5)

        tb.Button(
            btns_frame,
            text="📊 Összesítés",
            bootstyle=PRIMARY,
            command=lambda: self.run_command("aggregate", self.month_var.get()),
        ).pack(side=LEFT, padx=5)

        # Log Panel
        log_group = tb.Labelframe(self, text="Eseménynapló", padding=10)
        log_group.pack(fill=BOTH, expand=True)

        self.log_tree = tb.Treeview(
            log_group, columns=("msg"), show="headings", height=10
        )
        self.log_tree.heading("msg", text="Üzenet")
        self.log_tree.column("msg", width=800)
        self.log_tree.pack(fill=BOTH, expand=True)

        # Tags for coloring logs
        self.log_tree.tag_configure("info", foreground="#084298")
        self.log_tree.tag_configure("warn", foreground="#664d03")
        self.log_tree.tag_configure("err", foreground="#842029")

    def add_log(self, level: str, message: str):
        self.log_tree.insert("", "end", values=(message,), tags=(level,))
        self.log_tree.see(self.log_tree.get_children()[-1])
