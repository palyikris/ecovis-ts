import ttkbootstrap as tb
from ttkbootstrap.constants import *


class DashboardFrame(tb.Labelframe):
    """A reusable card for the dashboard displaying a title, value, and subtitle."""

    def __init__(self, parent, title, **kwargs):
        super().__init__(parent, text=title, padding=12, bootstyle=SECONDARY, **kwargs)

        self.value_label = tb.Label(
            self, text="—", font=("Segoe UI", 18, "bold"), anchor="w"
        )
        self.value_label.pack(fill=X)

        self.sub_label = tb.Label(
            self,
            text="",
            font=("Segoe UI", 9),
            bootstyle=SECONDARY,
            anchor="w",
            justify="left",
        )
        self.sub_label.pack(fill=X, pady=(6, 0))

    def update_stats(self, value: str, subtitle: str):
        """Updates the labels on the card."""
        self.value_label.config(text=value)
        self.sub_label.config(text=subtitle)

    def set_wrap(self, width: int):
        """Adjusts text wrapping based on container size."""
        self.sub_label.configure(wraplength=width)

    def add_log(self, message: str):
        """Appends a message to the subtitle log."""
        current = self.sub_label.cget("text")
        new_text = f"{current}\n{message}" if current else message
        self.sub_label.config(text=new_text)
