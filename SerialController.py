import serial
import time
import threading
# ----------------------------
# Serial controller
# ----------------------------

BUTTON_MAP = {
    "L": ("high", 0x01),
    "R": ("high", 0x02),
    "X": ("high", 0x04),
    "Y": ("high", 0x08),
    "A": ("low", 0x01),
    "B": ("low", 0x02),
    "Right": ("low", 0x04),
    "Left": ("low", 0x08),
    "Up": ("low", 0x10),
    "Down": ("low", 0x20),
    "Select": ("low", 0x40),
    "Start": ("low", 0x80),
}

ALL_BUTTONS = ["A", "B", "X", "Y", "Up", "Down", "Left", "Right", "Start", "Select", "L", "R"]


def buttons_to_bytes(buttons):
    high, low = 0, 0
    for b in buttons:
        if b not in BUTTON_MAP:
            raise ValueError(f"Unknown button: {b}")
        which, mask = BUTTON_MAP[b]
        if which == "high":
            high |= mask
        else:
            low |= mask
    return high & 0xFF, low & 0xFF


class SerialController:
    def __init__(self, status_cb=None ,app=None):
        self.status_cb = status_cb or (lambda s: None)
        self.app = app or (lambda s: None)
        self.ser = None
        self.interval_s = 0.05

        self._lock = threading.Lock()
        self._running = False
        self._thread = None

        self._high = 0
        self._low = 0

    @property
    def connected(self):
        return self.ser is not None and self.ser.is_open

    def _send_packet(self, high, low):
        """Immediately send a packet to the serial device."""
        if not self.connected:
            return
        try:
            self.ser.write(bytearray([0x54, high & 0xFF, low & 0xFF]))
            self.ser.flush()  # Ensure immediate transmission
        except Exception as e:
            self.status_cb(f"Serial write error: {e}")

    def set_state(self, high, low):
        """Set button state and immediately send to device."""
        with self._lock:
            self._high = high & 0xFF
            self._low = low & 0xFF
            # Immediately send the new state
            self._send_packet(self._high, self._low)

    def set_buttons(self, buttons):
        """Set buttons and immediately send to device."""
        high, low = buttons_to_bytes(buttons)
        self.set_state(high, low)

    def connect(self, port, baud=1_000_000):
        if self.connected:
            self.disconnect()

        self.ser = serial.Serial(port, baud, timeout=1)
        self.status_cb(f"Serial connected: {port} @ {baud}")

        self._running = True
        self._thread = threading.Thread(target=self._keepalive_loop, daemon=True)
        self._thread.start()

        # Pairing warm-up (neutral for ~3s)
        self.status_cb("Pairing warm-up: neutral for ~3 seconds...")
        self.set_state(0, 0)
        time.sleep(3.0)
        self.status_cb("Pairing warm-up done.")

    def disconnect(self):
        self._running = False
        if self._thread and self._thread.is_alive():
            self._thread.join(timeout=1.0)

        if self.ser:
            try:
                self.ser.close()
            except Exception:
                pass
        self.ser = None
        self.status_cb("Serial disconnected.")

    def send_channel_set(self, channel_byte):
        if not self.connected:
            raise RuntimeError("Not connected.")
        ch = int(channel_byte) & 0xFF
        pkt = bytearray([0x43, ch, 0x00])
        self.ser.write(pkt)
        self.ser.flush()
        self.status_cb(f"Sent channel set: 0x{ch:02X} (power cycle receiver required)")

    def _keepalive_loop(self):

        while self._running:
            if not self.connected or (self.app.backend_var.get() != "USB Serial"):
                break
            with self._lock:
                high = self._high
                low = self._low
            try:
                self.ser.write(bytearray([0x54, high, low]))
            except Exception as e:
                self.status_cb(f"Serial write error: {e}")
                break
            time.sleep(self.interval_s)

