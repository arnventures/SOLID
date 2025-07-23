"""
serial_manager.py – Thread-sicherer Modbus-Wrapper
mit automatischem Reconnect (Watch-Dog) und retries=0.
"""

from __future__ import annotations
import time
import threading
from typing import Callable
from pymodbus.client import ModbusSerialClient
from pymodbus.exceptions import ModbusIOException
import serial  # Direktes Importieren von pyserial

class SerialManager:
    def __init__(self, baudrate: int = 9600, log_cb: Callable[[str], None] | None = None, timeout: float = 1.0):
        self._baud = baudrate
        self._timeout = timeout
        self._log = log_cb or (lambda msg: print(msg))  # Fallback auf print
        self._lock = threading.Lock()
        self._client: ModbusSerialClient | None = None
        self._port: str | None = None
        self._wd_stop = threading.Event()
        self._wd_thread = threading.Thread(
            target=self._watchdog, name="SerialWatchdog", daemon=True
        )
        # Debugging: pyserial-Version protokollieren, falls verfügbar
        try:
            serial_version = getattr(serial, '__version__', 'unbekannt')
            self._log(f"pyserial version: {serial_version}")
        except Exception as e:
            self._log(f"Fehler beim Abrufen der pyserial-Version: {e}")
        self._wd_thread.start()

    def connect(self, port: str) -> bool:
        """Verbindet mit dem angegebenen COM-Port."""
        with self._lock:
            if self._client and self._client.is_socket_open():
                self._log(f"Bisherige Verbindung ({self._port}) wird geschlossen.")
                self._client.close()

            self._port = port
            self._log(f"Initialisiere ModbusSerialClient für Port {port} (Baudrate: {self._baud}, Timeout: {self._timeout})")
            try:
                self._client = ModbusSerialClient(
                    framer="rtu",
                    port=port,
                    baudrate=self._baud,
                    bytesize=8,
                    parity="N",
                    stopbits=1,
                    timeout=self._timeout,
                    retries=0  # Keine Auto-Retries
                )
                ok = self._client.connect()
                self._log(f"Verbindung zu {port}: {'OK' if ok else 'FEHLGESCHLAGEN'}")
                if not ok:
                    self._log(f"Fehler: Verbindung zu {port} konnte nicht hergestellt werden.")
                    self._client = None
                    self._port = None
                return ok
            except Exception as e:
                self._log(f"Fehler beim Verbindungsaufbau zu {port}: {e}")
                self._client = None
                self._port = None
                return False

    def close(self):
        """Schließt die Verbindung und stoppt den Watchdog."""
        self._wd_stop.set()
        with self._lock:
            if self._client:
                self._client.close()
                self._log(f"Verbindung zu {self._port} geschlossen.")
                self._client = None
                self._port = None

    @property
    def port(self):
        return self._port

    @property
    def is_open(self):
        return self._client is not None and self._client.is_socket_open()

    def read_holding(self, addr: int, count: int = 1, unit: int = 1):
        return self._call(self._client.read_holding_registers, addr, count=count, slave=unit)

    def write_single(self, addr: int, value: int, unit: int = 1):
        return self._call(self._client.write_register, addr, value=value, slave=unit)

    def _call(self, fn, *args, **kw):
        """Führt eine Modbus-Operation mit Retry aus."""
        for attempt in (1, 2):
            self._ensure_open()
            try:
                if not self.is_open:
                    raise ModbusIOException(f"Port {self._port} nicht verbunden")
                res = fn(*args, **kw)
                return res
            except ModbusIOException as e:
                self._log(f"Serial I/O-Fehler bei {self._port} (Versuch {attempt}): {e}")
                with self._lock:
                    if self._client:
                        self._client.close()
                time.sleep(0.2)
            except Exception as e:
                self._log(f"Unerwarteter Fehler bei Modbus-Operation (Versuch {attempt}): {e}")
                raise
        raise ModbusIOException(f"Serial-Port {self._port} nach Retry nicht verfügbar")

    def _ensure_open(self):
        """Stellt sicher, dass der Port geöffnet ist."""
        with self._lock:
            if self._client and not self._client.is_socket_open() and self._port:
                self._log(f"Port {self._port} geschlossen – reconnect …")
                try:
                    ok = self._client.connect()
                    self._log(f"Reconnect zu {self._port}: {'OK' if ok else 'FEHLGESCHLAGEN'}")
                except Exception as e:
                    self._log(f"Reconnect-Fehler bei {self._port}: {e}")

    def _watchdog(self):
        """Überwacht die Verbindung und reconnectet bei Bedarf."""
        while not self._wd_stop.is_set():
            time.sleep(5)
            with self._lock:
                if self._client and not self._client.is_socket_open() and self._port:
                    self._log(f"Watch-Dog: Port {self._port} geschlossen – reconnect …")
                    try:
                        ok = self._client.connect()
                        self._log(f"Watch-Dog: Reconnect zu {self._port} {'OK' if ok else 'FEHLGESCHLAGEN'}")
                    except Exception as e:
                        self._log(f"Watch-Dog: Reconnect-Fehler bei {self._port}: {e}")