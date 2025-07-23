"""
gui.py – Inosent Anlage Aufbau Tool
Frontend mit Checkboxen, Skip-Button und endlosem Warten auf Gerät @1
--------------------------------------------------------------------
Abhängigkeiten:
    serial_manager.py
    project_context.py
    excel_service.py
    label_printer.py
"""

from __future__ import annotations
import tkinter as tk, sys
from tkinter import ttk, scrolledtext, messagebox
import tkinter.font as tkFont
import threading, time, serial.tools.list_ports, pathlib

from serial_manager import SerialManager
from project_context import ProjectContext
from excel_service import load_sensor_data
from label_printer import print_led_labels, print_sensor_labels

# ------------------------------------------------------------------ #
CTX  = ProjectContext()
SER  = SerialManager()              # ggf. log_cb=self._log später setzen
STOP = threading.Event()
SKIP = threading.Event()

# ------------------------------------------------------------------ #
class SensorGUI(tk.Tk):
    
    COLS = ("Sel", "Index", "Model", "Location", "Address", "Buzzer", "Serial", "Status")

    # ------------------------ Init ------------------------------------ #
    def __init__(self):
        super().__init__()

        menufont = tkFont.Font(family="Segoe UI", size=12, weight="bold")
        self.option_add("*Menu*Font", menufont)
        
        self.title("Inosent Anlage Aufbau Tool")
        self.geometry("1020x740")

        base = pathlib.Path(getattr(sys, "_MEIPASS", pathlib.Path(__file__).parent))
        ico  = base / "inosent.ico"
        if ico.is_file():
            try: self.iconbitmap(ico)
            except Exception: pass

        self._selected: dict[str, bool] = {}
        self._build_ui()

    # ------------------------ GUI-Aufbau ------------------------------ #
    def _build_ui(self):
        
        # Menü
        menu = tk.Menu(self); self.config(menu=menu)
        cfg_m = tk.Menu(menu, tearoff=0)
        cfg_m.add_command(label="Stop", command=self._stop_cfg)
        menu.add_cascade(label="Configuration", menu=cfg_m)

        prn_m = tk.Menu(menu, tearoff=0)
        prn_m.add_command(label="LED-Labels",
                          command=lambda: print_led_labels(CTX, self, self._log))
        prn_m.add_command(label="Sensor-Labels",
                          command=lambda: print_sensor_labels(CTX, self, self._log))
        menu.add_cascade(label="Print", menu=prn_m)

        help_m = tk.Menu(menu, tearoff=0)
        help_m.add_command(label="Hilfe / Über …", command=self._show_help)
        menu.add_cascade(label="Hilfe", menu=help_m)

        # Top-Bar
        top = ttk.Frame(self); top.pack(pady=10, padx=6, anchor="w")
        self.btn_start = ttk.Button(top, text="▶  Start", width=11,
                                    command=self._start_cfg)
        self.btn_start.grid(row=0, column=0, rowspan=2, sticky="ns", padx=(0,12))

        self.btn_skip = ttk.Button(top, text="⤼  Skip", width=7,
                                   command=lambda: SKIP.set(), state="disabled")
        self.btn_skip.grid(row=0, column=1, rowspan=2, sticky="ns", padx=(0,18))

        ttk.Label(top, text="COM-Port:").grid(row=0, column=2, sticky="e")
        self.cb_port = ttk.Combobox(top, width=12, state="readonly")
        self.cb_port.grid(row=0, column=3, padx=5)
        ttk.Button(top, text="Refresh", command=self._refresh_ports)\
            .grid(row=0, column=4, padx=5)

        ttk.Label(top, text="Anlage-Nr.:").grid(row=1, column=2, sticky="e")
        self.ent_nr = ttk.Entry(top, width=18)
        self.ent_nr.grid(row=1, column=3, padx=5)
        ttk.Button(top, text="OK", command=self._load_excel)\
            .grid(row=1, column=4, padx=5)

        self.lab_current = ttk.Label(top, text="Current Anlage: –")
        self.lab_current.grid(row=1, column=5, padx=15, sticky="w")

        # Status / Tree / Log
        self.lab_status = ttk.Label(self, text="Bereit", background="yellow")
        self.lab_status.pack(fill="x", pady=5)

        self.tree = ttk.Treeview(self, columns=self.COLS,
                                 show="headings", height=18)

        for col in self.COLS:
            if   col == "Sel":      w = 60       # Basiseinheit
            elif col == "Index":    w = 60
            elif col == "Model":    w = 120      # 3× Basis
            elif col == "Location": w = 300      # 10× Basis
            elif col == "Address":  w = 60
            else:                   w = 120      # Buzzer / Serial / Status
            self.tree.heading(col, text=col)
            self.tree.column(col, width=w, anchor="center")

        self.tree.pack(fill="x", padx=6, pady=10)
        self.tree.bind("<Button-1>", self._on_tree_click)

        # Zeilenfarben
        self.tree.tag_configure("active", background="#6089FA")   # gelb
        self.tree.tag_configure("ok",     background="#16D304")   # grün
        self.tree.tag_configure("fail",   background="#D45353")   # rot

        self.log = scrolledtext.ScrolledText(self, height=6, width=108)
        self.log.pack(padx=6, pady=6)

        # Init
        self._refresh_ports()
        self.cb_port.bind("<<ComboboxSelected>>", self._on_port_change)
        self.protocol("WM_DELETE_WINDOW", self._on_close)

    # ------------------------ Helper ---------------------------------- #
    def _log(self, msg: str):
        self.log.insert(tk.END, f"{time.strftime('%H:%M:%S')}: {msg}\n")
        self.log.see(tk.END)

    # ---------- COM-Port --------------------------------------------- #
    def _refresh_ports(self):
        ports = [p.device for p in serial.tools.list_ports.comports()]
        self.cb_port["values"] = ports
        if ports:
            self.cb_port.set(ports[0]); self._on_port_change()

    def _on_port_change(self, *_):
        p = self.cb_port.get()
        ok = p and SER.connect(p)
        self.lab_status.config(text=f"{p}: connected" if ok else "No Port",
                               background="lightgreen" if ok else "red")


    # ---------- Excel laden ------------------------------------------ #
    def _load_excel(self):
        nr = self.ent_nr.get()
        if not nr.isdigit():
            messagebox.showerror("Fehler",
                                 "Bitte gültige Anlage-Nr. eingeben.")
            return

        CTX.nr, CTX.xlsx = int(nr), None
        if not CTX.ensure_loaded(self):
            return

        # Excel öffnen – Sperre abfangen
        try:
            wb, ws, sensors = load_sensor_data(CTX.xlsx)
        except PermissionError:
            messagebox.showwarning(
                "Excel-Datei gesperrt",
                "Die Excel-Liste ist noch geöffnet oder wird von einem "
                "anderen Programm (z. B. OneDrive / Excel) benutzt.\n\n"
                "Bitte schließen Sie die Datei und versuchen Sie es erneut.",
                parent=self
            )
            self._log("Excel gesperrt – Vorgang abgebrochen.")
            return
        except Exception as e:
            messagebox.showerror("Fehler",
                                 f"Excel konnte nicht geladen werden:\n{e}",
                                 parent=self)
            self._log(f"[ERROR] Excel-Load: {e}")
            return

        # Tabelle füllen
        self.tree.delete(*self.tree.get_children())
        self._selected.clear()

        for i, (_, model, loc, addr, buzzer, _sn) in enumerate(sensors, 1):
            buz = "Disable" if buzzer == "Buzzer Disable" else "Enable"
            iid = self.tree.insert(
                "", "end",
                values=("☑", i, model, loc, addr, buz, "", "Pending"),
                tags=()
            )
            self._selected[iid] = True

        self.lab_current.config(text=f"Current Anlage: {CTX.nr}")
        self._log(f"Excel geladen – {len(sensors)} Sensoren.")

    # ---------- Checkbox-Toggle -------------------------------------- #
    def _on_tree_click(self, event):
        if self.tree.identify_column(event.x) != "#1": return
        iid = self.tree.identify_row(event.y)
        if iid:
            sel = not self._selected[iid]
            self._selected[iid] = sel
            self.tree.set(iid, "Sel", "☑" if sel else "☐")

    # ---------- Start / Stop / Skip ---------------------------------- #
    def _start_cfg(self):
        if CTX.xlsx is None:
            messagebox.showinfo("Info", "Bitte zuerst Anlage laden (OK)."); return
        if not SER.port:
            messagebox.showerror("Fehler", "Kein COM-Port verbunden."); return

        wb, ws, all_sensors = load_sensor_data(CTX.xlsx)
        ids = [iid for iid in self.tree.get_children() if self._selected[iid]]
        sensors = [all_sensors[int(self.tree.set(iid, "Index")) - 1] for iid in ids]

        if not sensors:
            messagebox.showinfo("Info", "Keine Sensoren ausgewählt."); return

        STOP.clear(); SKIP.clear()
        self.btn_start.config(state="disabled"); self.btn_skip.config(state="normal")
        threading.Thread(target=self._worker,
                         args=(wb, ws, sensors, ids), daemon=True).start()

    def _stop_cfg(self):
        STOP.set(); SKIP.set()
        self.btn_skip.config(state="disabled"); self.btn_start.config(state="normal")
        self._log("Konfiguration gestoppt.")

    # ---------- Worker-Thread ---------------------------------------- #
    def _worker(self, wb, ws, sensors, ids):
        prev_iid = None
        try:
            self.lab_status.config(text="Läuft …", background="yellow")

            for (row, model, loc, addr, buzzer, _), iid in zip(sensors, ids):
                if STOP.is_set():
                    break

                # Zeile als aktiv markieren
                if prev_iid:
                    self.tree.item(prev_iid, tags=())
                self.tree.item(iid, tags=("active",))
                prev_iid = iid

                # Skip?
                if SKIP.is_set():
                    SKIP.clear()
                    self.tree.set(iid, "Status", "Skipped")
                    self.tree.item(iid, tags=())
                    self._log(f"Sensor {row-1} übersprungen.")
                    continue

                # Auf Sensor @1 warten
                if not self._wait_for_device_one():
                    self.tree.set(iid, "Status", "Skipped")
                    self.tree.item(iid, tags=())
                    continue

                # Konfigurieren (ohne Boot-Wait / Poll)
                sn = self._configure_single(ws, row, addr,
                                            buzzer == "Buzzer Disable")

                # Ergebnis sofort anzeigen
                if sn is not None:
                    self.tree.set(iid, "Serial", sn)
                    self.tree.set(iid, "Status", "OK")
                    self.tree.item(iid, tags=("ok",))
                    wb.save(CTX.xlsx)
                else:
                    self.tree.set(iid, "Status", "Fail")
                    self.tree.item(iid, tags=("fail",))

        except Exception as e:
            self._log(f"[ERROR] {e}")

        finally:
            if prev_iid:
                self.tree.item(prev_iid, tags=())
            done = not STOP.is_set()
            self.lab_status.config(text="Fertig" if done else "Abgebrochen",
                                   background="lightgreen")
            self.btn_skip.config(state="disabled")
            self.btn_start.config(state="normal")

    # ---------- Gerät @1 suchen -------------------------------------- #
    def _wait_for_device_one(self) -> bool:
        self._log("Suche Gerät @1 … (Skip überspringt; Stop beendet)")
        while not STOP.is_set():
            if SKIP.is_set():
                SKIP.clear()
                return False
            try:
                r = SER.read_holding(2, unit=1)
                if not r.isError():
                    return True
            except Exception as e:
                self._log(f"Modbus-Fehler (wait): {e}")
            time.sleep(1)
        return False

POLL_MAX_S = 2.0     # max. 2 s Verifikation
POLL_STEP  = 0.2     # 5 Polls pro Sekunde

# ---------- Einzel-Konfiguration ---------------------------------- #
def _configure_single(self, ws_gas, row, new_addr, disable_bz) -> int | None:
    """
    Schnelle Konfig mit Kurz-Verifikation (max. 2 s).
    Schreibt Serien-Nr. in Import!E.
    """
    if SKIP.is_set(): SKIP.clear(); return None
    try:
        # 1) Serien-Nr. lesen
        res = SER.read_holding(3, unit=1)
        if res.isError(): return None
        serial = res.registers[0]

        # 2) Adresse schreiben
        if SER.write_single(4, new_addr, unit=1).isError(): return None

        # 3) Buzzer optional
        if disable_bz:
            rr = SER.read_holding(255, unit=1)
            if rr.isError(): return None
            SER.write_single(255, rr.registers[0] & ~(1 << 9), unit=1)

        # 4) Reboot
        if SER.write_single(17, 42330, unit=1).isError(): return None

        # 5) Kurze Verifikation ≤ 2 s
        deadline = time.time() + POLL_MAX_S
        while time.time() < deadline:
            if SKIP.is_set(): SKIP.clear(); return None
            if not SER.read_holding(2, unit=new_addr).isError():
                break                          # Sensor antwortet ⇒ OK
            time.sleep(POLL_STEP)
        else:
            self._log(f"Sensor @{new_addr} keine Antwort – Fail")
            return None                        # Timeout ⇒ Fail

        # 6) Serien-Nr. in Blatt „Import“ schreiben
        wb        = ws_gas.parent
        ws_imp    = wb["Import"]
        if row > ws_imp.max_row:
            while ws_imp.max_row < row - 1:
                ws_imp.append([None])
            ws_imp.append([None, None, None, None, serial])
        else:
            ws_imp.cell(row=row, column=5, value=serial)

        return serial

    except Exception as e:
        self._log(f"Config-Fehler: {e}")
        return None




    # ---------- Help-Dialog ------------------------------------------ #
    def _show_help(self):
        msg = (
            "Inosent Anlage Aufbau Tool  v1.0\n\n"
            "Ablauf:\n"
            "1. COM-Port verbinden (grüne Statusleiste).\n"
            "2. Anlage-Nr. eingeben → OK (Excel-Liste lädt).\n"
            "3. Checkbox abwählen = Sensor überspringen.\n"
            "4. ▶ Start – Tool wartet auf Gerät @1, konfiguriert, "
            "schreibt Serien-Nr., fährt fort …\n"
            "5. ⤼ Skip überspringt aktuellen Sensor, Stop bricht ab.\n\n"
            "Menü Print druckt LED- bzw. Sensor-Etiketten."
        )
        messagebox.showinfo("Hilfe / Über", msg, parent=self)

    # ---------- Cleanup ---------------------------------------------- #
    def _on_close(self):
        STOP.set(); SKIP.set(); SER.close(); self.destroy()


# ------------------------------------------------------------------ #
if __name__ == "__main__":
    SensorGUI().mainloop()
