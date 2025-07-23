"""
project_context.py
------------------
Hält Anlage-Nr. und Excel-Pfad zentral.
"""
from __future__ import annotations
import os
import tkinter as tk
from tkinter import filedialog, messagebox, simpledialog


class ProjectContext:
    def __init__(self):
        self.nr: int | None = None
        self.xlsx: str | None = None

    # ------------------------------------------------------------ #
    def ensure_loaded(self, parent) -> bool:
        """Fragt interaktiv nach Nr./Datei, falls noch nicht gesetzt."""
        if self.nr is None:
            self.nr = simpledialog.askinteger(
                "Anlage-Nr.", "Anlage Nummer:", minvalue=1, parent=parent)
            if not self.nr:
                return False

        if self.xlsx is None:
            self.xlsx = _find_excel_path(self.nr)
            if self.xlsx is None:
                messagebox.showinfo(
                    "Excel auswählen",
                    "Liste nicht gefunden – bitte manuell wählen.",
                    parent=parent,
                )
                f = filedialog.askopenfilename(
                    title="Excel auswählen",
                    filetypes=[("Excel", "*.xlsm *.xlsx *.xls")],
                    parent=parent,
                )
                if not f:
                    return False
                self.xlsx = f
        return True


# ------------------------------------------------------------ #
def _find_excel_path(anlage_nr: int) -> str | None:
    first_two = str(anlage_nr)[:2]
    base = os.path.join(
        r"T:\INOSENT_Projekte",
        "20" + first_two,
        str(anlage_nr),
        f"{anlage_nr}_Anlageinfos",
        f"DS_{anlage_nr}",
    )
    for ext in (".xlsm", ".xlsx", ".xls"):
        p = os.path.join(base, f"Liste_{anlage_nr}{ext}")
        if os.path.exists(p):
            return p
    return None
