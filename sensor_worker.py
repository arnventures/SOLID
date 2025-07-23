"""
sensor_worker.py – robust & schneller
-------------------------------------
1. Warten auf *neuen* Sensor @1            (Serial noch nicht gesehen)
2. Adresse / Buzzer setzen, Soft-Restart
3. Warten bis Sensor unter neuer Adresse antwortet (≤ 10 s)
4. Serien-Nr. in Excel & TreeView schreiben
"""
from __future__ import annotations
import time
import openpyxl
from typing import List, Dict, Callable
from serial_manager import SerialManager

# Tuning-Konstanten (ggf. anpassen)
BOOT_WAIT  = 0.5      # feste Pause nach Soft-Restart (alt 5 s)
POLL_STEP  = 0.25     # Poll-Intervall Adresse-Check   (alt 0.5 s)
POLL_MAX   = 3.0     # max. Wartezeit Adresse-Check   (alt 15 s)
POLL_ROUNDS = int(POLL_MAX / POLL_STEP)


class SensorWorker:
    def __init__(self,
                 rows: List[Dict],
                 workbook: openpyxl.Workbook,
                 worksheet: openpyxl.worksheet.worksheet.Worksheet,
                 ser: SerialManager,
                 log: Callable[[str], None],
                 stop_event,
                 skip_event):
        self.rows, self.wb, self.ws = rows, workbook, worksheet
        self.ser, self.log = ser, log
        self.stop_event, self.skip_event = stop_event, skip_event
        self.prev_serials: set[int] = set()        # bereits konfigurierte SN

    # ------------------------------------------------------------------ #
    def run(self):
        self.log("Worker gestartet.")
        for item in self.rows:
            if not item["enabled"]:
                self.log(f"Sensor {item['row']-1} übersprungen (deselektiert).")
                continue
            if self.stop_event.is_set():
                break

            # 1) auf neues Gerät @1 warten
            if not self._wait_for_addr1():
                break

            # 2) konfigurieren
            ok, serial = self._configure(item)
            item["status_cb"](ok)                 # GUI-Callback

            # 3) Serien-Nr. merken + Excel speichern
            if ok and serial:
                self.prev_serials.add(serial)
                self.wb.save(self.wb.filename)

        self.log("Worker beendet.")

    # ------------------------------------------------------------------ #
    def _wait_for_addr1(self) -> bool:
        """Blockiert bis *neuer* Sensor @1 antwortet oder Skip/Stop."""
        self.log("Warte auf *neues* Gerät @1 …")
        while not self.stop_event.is_set():
            if self.skip_event.is_set():
                self.skip_event.clear()
                self.log("Warten abgebrochen (Skip).")
                return False
            try:
                res = self.ser.read_holding(3, unit=1)  # Serial lesen
                if not res.isError():
                    serial = res.registers[0]
                    if serial not in self.prev_serials:
                        self.log(f"Gerät @1 gefunden (SN={serial}).")
                        return True
            except Exception as e:
                self.log(f"Modbus-Fehler (wait): {e}")
            time.sleep(1)
        return False

    # ------------------------------------------------------------------ #
    def _configure(self, item: Dict) -> tuple[bool, int | None]:
        row, new_addr, disable_bz = item["row"], item["new_addr"], item["buzzer"]

        if self.skip_event.is_set():
            self.skip_event.clear()
            return False, None

        self.log(f"Config {row-1} → Adresse {new_addr}")
        try:
            # Serial lesen
            res = self.ser.read_holding(3, unit=1)
            if res.isError():
                raise RuntimeError(res)
            serial = res.registers[0]

            # Adresse schreiben
            if self.ser.write_single(4, new_addr, unit=1).isError():
                raise RuntimeError("addr write")

            # Buzzer optional deaktivieren
            if disable_bz:
                r = self.ser.read_holding(255, unit=1)
                if r.isError():
                    raise RuntimeError(r)
                self.ser.write_single(255, r.registers[0] & ~(1 << 9), unit=1)

            # Soft-Restart
            if self.ser.write_single(17, 42330, unit=1).isError():
                raise RuntimeError("restart")

            # feste Boot-Pause
            time.sleep(BOOT_WAIT)

            # Poll-Schleife ≤ POLL_MAX
            deadline = time.time() + POLL_MAX
            while time.time() < deadline:
                if self.skip_event.is_set():
                    self.skip_event.clear()
                    return False, None
                if not self.ser.read_holding(2, unit=new_addr).isError():
                    break
                time.sleep(POLL_STEP)
            else:
                raise RuntimeError("no reply @ new addr")

            # Excel & Dict aktualisieren
            self.ws[f"E{row}"] = serial
            item["serial"] = serial
            self.log(f"Sensor {row-1} OK (SN={serial})")
            return True, serial

        except Exception as e:
            self.log(f"Config-Fehler: {e}")
            return False, None
