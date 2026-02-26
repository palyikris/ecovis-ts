import ttkbootstrap as tb
from ttkbootstrap.constants import *


class DashboardCard(tb.Labelframe):
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

    def update_card(self, value, subtitle):
        self.value_label.config(text=value)
        self.sub_label.config(text=subtitle)

    def set_wrap(self, width):
        self.sub_label.configure(wraplength=width)
