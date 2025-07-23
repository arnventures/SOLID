"""
label_printer.py
----------------
Druckt LED- & Sensor-Etiketten über GoLabel II.
"""
from __future__ import annotations
import subprocess
import os
from tkinter import messagebox
from project_context import ProjectContext
from excel_service import copy_led_data, copy_sensor_data

DEST_FOLDER = r"C:\PowerAutomateArnDoNotDelete"


def _run_godex(template: str, printer_ip: str) -> bool:
    try:
        subprocess.run(
            [r"C:\Program Files (x86)\GoDEX\GoLabel II\GoLabel.exe",
             "-f", template, "-c", "2", "-i", printer_ip],
            check=True)
        return True
    except subprocess.CalledProcessError:
        return False


def print_led_labels(ctx: ProjectContext, parent, log):
    if not ctx.ensure_loaded(parent):
        log("Abgebrochen: keine Anlage gewählt."); return
    dst = os.path.join(DEST_FOLDER, "PowerAutomateGodexLightDe.xlsx")
    if copy_led_data(ctx.xlsx, dst):
        if messagebox.askyesno("Drucken", "LED-Etiketten drucken?", parent=parent):
            ok = _run_godex(
                template=os.path.join(DEST_FOLDER, "AutomaticLightDE.ezpx"),
                printer_ip="10.1.40.88:9100")
            log("LED-Etiketten gedruckt." if ok else "Druckfehler (LED).")


def print_sensor_labels(ctx: ProjectContext, parent, log):
    if not ctx.ensure_loaded(parent):
        log("Abgebrochen: keine Anlage gewählt."); return
    dst = os.path.join(DEST_FOLDER, "PowerAutomateGodexSensorDe.xlsx")
    if copy_sensor_data(ctx.xlsx, dst):
        if messagebox.askyesno("Drucken", "Sensor-Etiketten drucken?", parent=parent):
            ok = _run_godex(
                template=os.path.join(DEST_FOLDER, "AutomaticSensorDE.ezpx"),
                printer_ip="10.1.40.87:9100")
            log("Sensor-Etiketten gedruckt." if ok else "Druckfehler (Sensor).")
