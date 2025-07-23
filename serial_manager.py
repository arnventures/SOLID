"""
serial_manager.py – Thread-sicherer Modbus-Wrapper
mit automatischem Reconnect (Watch-Dog) und retries=0.
"""

from __future__ import annotations
import time, threading
from typing import Callable
from pymodbus.client import ModbusSerialClient
from pymodbus.exceptions import ModbusIOException


class SerialManager:
    def __init__(self,
                 baudrate: int = 9600,
                 log_cb: Callable[[str], None] | None = None,
                 timeout: float = 1.0):
        self._baud = baudrate
        self._timeout = timeout
        self._log = log_cb or (lambda *_: None)
        self._lock = threading.Lock()

        self._client: ModbusSerialClient | None = None
        self._port: str | None = None

        # Watch-Dog
        self._wd_stop = threading.Event()
        self._wd_thread = threading.Thread(
            target=self._watchdog, name="SerialWatchdog", daemon=True
        )
        self._wd_thread.start()

    # ------------------------------------------------------------------ #
    # Public API
    # ------------------------------------------------------------------ #
    def connect(self, port: str) -> bool:
        with self._lock:
            self._port = port
            self._client = ModbusSerialClient(
                framer="rtu",
                port=port,
                baudrate=self._baud,
                bytesize=8,
                parity="N",
                stopbits=1,
                timeout=self._timeout,
                retries=0                 # << keine Auto-Retries, kein Auto-Close
            )
            ok = self._client.connect()
            self._log(f"Serial connect {port}: {'OK' if ok else 'FAIL'}")
            return ok

    def close(self):
        self._wd_stop.set()
        with self._lock:
            if self._client:
                self._client.close()

    @property
    def port(self):      return self._port
    @property
    def is_open(self):   return self._client and self._client.is_socket_open()

    # ------------------------------------------------------------------ #
    # Modbus Convenience
    # ------------------------------------------------------------------ #
    def read_holding(self, addr: int, count: int = 1, unit: int = 1):
        return self._call(self._client.read_holding_registers,
                          addr, count=count, slave=unit)

    def write_single(self, addr: int, value: int, unit: int = 1):
        return self._call(self._client.write_register,
                          addr, value=value, slave=unit)

    # ------------------------------------------------------------------ #
    # Intern – Aufruf mit einfachem Retry falls Client geschlossen
    # ------------------------------------------------------------------ #
    def _call(self, fn, *args, **kw):
        for attempt in (1, 2):          # maximal 1 Reconnect-Versuch
            self._ensure_open()
            try:
                res = fn(*args, **kw)
                return res
            except ModbusIOException as e:
                self._log(f"Serial I/O-Error: {e} (Attempt {attempt})")
                with self._lock:
                    if self._client:
                        self._client.close()
                time.sleep(0.2)
        raise ModbusIOException("Serial port unavailable after retry")

    def _ensure_open(self):
        with self._lock:
            if self._client and not self._client.is_socket_open():
                self._log("Port geschlossen – reconnect …")
                self._client.connect()

    # ------------------------------------------------------------------ #
    # Background Watch-Dog – check every 5 s
    # ------------------------------------------------------------------ #
    def _watchdog(self):
        while not self._wd_stop.is_set():
            time.sleep(5)
            with self._lock:
                if self._client and not self._client.is_socket_open():
                    self._log("Watch-Dog: Port zu – reconnect …")
                    ok = self._client.connect()
                    self._log("Watch-Dog: reconnect OK" if ok else
                              "Watch-Dog: reconnect FAILED")
