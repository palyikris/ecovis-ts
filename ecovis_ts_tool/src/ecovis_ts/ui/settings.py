import tkinter as tk
import ttkbootstrap as tb
from ttkbootstrap.constants import *
from tkinter import filedialog, messagebox
from ..config import SETTINGS, save_settings


class SettingsFrame(tb.Frame):
    def __init__(self, parent, **kwargs):
        super().__init__(parent, padding=30, **kwargs)
        self._setup_ui()

    def _setup_ui(self):
        tb.Label(self, text="⚙️ Beállítások", font=("Segoe UI", 18, "bold")).pack(
            anchor=W, pady=(0, 20)
        )

        # Path Settings
        path_group = tb.Labelframe(self, text="Elérési utak", padding=15)
        path_group.pack(fill=X, pady=10)

        # TS Folder
        row1 = tb.Frame(path_group)
        row1.pack(fill=X, pady=5)
        tb.Label(row1, text="Timesheet mappa:", width=20).pack(side=LEFT)
        self.ts_folder_var = tk.StringVar(value=SETTINGS.get("ts_folder", ""))
        tb.Entry(row1, textvariable=self.ts_folder_var).pack(
            side=LEFT, fill=X, expand=True, padx=5
        )
        tb.Button(
            row1, text="Tallózás", command=self._browse_ts, bootstyle=SECONDARY
        ).pack(side=LEFT)

        # Output Folder
        row2 = tb.Frame(path_group)
        row2.pack(fill=X, pady=5)
        tb.Label(row2, text="Kimeneti mappa:", width=20).pack(side=LEFT)
        self.out_folder_var = tk.StringVar(value=SETTINGS.get("output_folder", ""))
        tb.Entry(row2, textvariable=self.out_folder_var).pack(
            side=LEFT, fill=X, expand=True, padx=5
        )
        tb.Button(
            row2, text="Tallózás", command=self._browse_out, bootstyle=SECONDARY
        ).pack(side=LEFT)

        # Preferences
        pref_group = tb.Labelframe(self, text="Preferenciák", padding=15)
        pref_group.pack(fill=X, pady=10)

        self.auto_open_var = tk.BooleanVar(
            value=SETTINGS.get("auto_open_output_on_success", True)
        )
        tb.Checkbutton(
            pref_group,
            text="Fájlok automatikus megnyitása futás után",
            variable=self.auto_open_var,
        ).pack(anchor=W)

        # Save Button
        tb.Button(
            self, text="Beállítások mentése", bootstyle=SUCCESS, command=self._save
        ).pack(pady=20, side=RIGHT)

    def _browse_ts(self):
        path = filedialog.askdirectory()
        if path:
            self.ts_folder_var.set(path)

    def _browse_out(self):
        path = filedialog.askdirectory()
        if path:
            self.out_folder_var.set(path)

    def _save(self):
        SETTINGS["ts_folder"] = self.ts_folder_var.get()
        SETTINGS["output_folder"] = self.out_folder_var.get()
        SETTINGS["auto_open_output_on_success"] = self.auto_open_var.get()
        save_settings(SETTINGS)
        messagebox.showinfo("Siker", "A beállítások mentése megtörtént.")
