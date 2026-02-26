import tkinter as tk
import ttkbootstrap as tb
from ttkbootstrap.constants import *
from ecovis_ts.config import SETTINGS, save_settings


class SettingsFrame(tb.Frame):
    def __init__(self, parent, **kwargs):
        super().__init__(parent, **kwargs)
        self._setup_scrollable_area()
        self._setup_sections()

    def _setup_scrollable_area(self):
        self.canvas = tb.Canvas(self)
        self.canvas.pack(side=LEFT, fill=BOTH, expand=True)

        self.scrollbar = tb.Scrollbar(self, orient=VERTICAL, command=self.canvas.yview)
        self.scrollbar.pack(side=RIGHT, fill=Y)

        self.canvas.configure(yscrollcommand=self.scrollbar.set)
        self.scroll_content = tb.Frame(self.canvas, padding=30)

        self.canvas.create_window((0, 0), window=self.scroll_content, anchor="nw")
        self.scroll_content.bind(
            "<Configure>",
            lambda e: self.canvas.configure(scrollregion=self.canvas.bbox("all")),
        )

    def _setup_sections(self):
        # Paths Group
        path_group = tb.Labelframe(self.scroll_content, text="📂 Mappák", padding=12)
        path_group.pack(fill=X, pady=8)

        # Example Entry
        tb.Label(path_group, text="TS mappa:").grid(
            row=0, column=0, sticky=W, padx=4, pady=4
        )
        self.ts_folder_var = tk.StringVar(value=SETTINGS.get("ts_folder", ""))
        tb.Entry(path_group, textvariable=self.ts_folder_var, width=50).grid(
            row=0, column=1, padx=4
        )

        # Save Button at bottom
        tb.Button(
            self.scroll_content,
            text="Beállítások mentése",
            bootstyle=SUCCESS,
            command=self.save,
        ).pack(side=RIGHT, pady=20)

    def save(self):
        SETTINGS["ts_folder"] = self.ts_folder_var.get()
        save_settings(SETTINGS)
        tk.messagebox.showinfo("Siker", "Beállítások elmentve!")
