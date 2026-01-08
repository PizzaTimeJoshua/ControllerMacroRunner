from typing import Tuple
from dataclasses import dataclass, field
import socket
import struct
from enum import IntEnum
import time
# -----------------------------
# 3DS Input Redirection backend
# -----------------------------

TOUCH_W_PX = 320
TOUCH_H_PX = 240
HID_AXIS_MAX = 0xFFF

def precise_sleep(duration_sec):
    """High-precision sleep for 3DS timing"""
    if duration_sec <= 0:
        return
    if duration_sec < 0.002:
        end = time.perf_counter() + duration_sec
        while time.perf_counter() < end:
            pass
        return
    sleep_until = time.perf_counter() + duration_sec - 0.002
    while time.perf_counter() < sleep_until:
        remaining = sleep_until - time.perf_counter()
        if remaining > 0.005:
            time.sleep(min(remaining * 0.5, 0.001))
        else:
            break
    end = time.perf_counter() + duration_sec - (time.perf_counter() - (sleep_until - duration_sec + 0.002))
    while time.perf_counter() < end:
        pass

def clamp(v: int, lo: int, hi: int) -> int:
    return max(lo, min(hi, v))

def px_to_hid(x_px: int, y_px: int) -> Tuple[int, int]:
    # Your priority snippet used integer multipliers for precision.
    x = x_px * 12
    y = y_px * 17
    return clamp(x, 0, HID_AXIS_MAX), clamp(y, 0, HID_AXIS_MAX)

class ThreeDSButton(IntEnum):
    A = 0
    B = 1
    SELECT = 2
    START = 3
    RIGHT = 4
    LEFT = 5
    UP = 6
    DOWN = 7
    R = 8
    L = 9
    X = 10
    Y = 11
    ZL = 14
    ZR = 15

class ThreeDSInterfaceButton(IntEnum):
    HOME = 0
    POWER = 1

@dataclass
class ThreeDSClient:
    ip: str
    port: int = 4950

    # Default values (from your priority snippet)
    hid_pad: bytearray = field(default_factory=lambda: bytearray.fromhex("ff 0f 00 00"))
    touch_state: bytearray = field(default_factory=lambda: bytearray.fromhex("00 00 00 02"))
    circle_pad: bytearray = field(default_factory=lambda: bytearray.fromhex("ff f7 7f 00"))
    cpp_state: bytearray = field(default_factory=lambda: bytearray.fromhex("81 00 80 80"))
    interface_buttons: bytearray = field(default_factory=lambda: bytearray.fromhex("00 00 00 00"))

    _sock: socket.socket = field(init=False, repr=False)
    _addr: tuple = field(init=False, repr=False)

    def __post_init__(self) -> None:
        self._addr = (self.ip, self.port)
        self._sock = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
        self._sock.setblocking(False)

    @staticmethod
    def _set_bit(arr: bytearray, bit_index: int) -> None:
        byte_i, bit = divmod(bit_index, 8)
        arr[byte_i] ^= (1 << bit)

    # Sends 20-byte packet
    def send_update(self) -> None:
        packet = self.hid_pad + self.touch_state + self.circle_pad + self.cpp_state + self.interface_buttons
        self._sock.sendto(packet, self._addr)

    # Clears Touch Screen Inputs
    def reset_touch(self) -> None:
        self.touch_state[:] = bytearray.fromhex("00 00 00 02")

    # Touch Screen (press/hold)
    def press_touch(self, x_px: int, y_px: int) -> None:
        x12, y12 = px_to_hid(x_px, y_px)
        xy = (y12 << 12) | x12
        data = bytearray(struct.pack("<I", xy))
        data[3] = 1
        self.touch_state[:] = data

    # Held buttons API (clean for our app)
    def set_buttons_held(
        self,
        pressed: set[ThreeDSButton],
        pressed_iface: set[ThreeDSInterfaceButton] = set()
    ) -> None:
        # Reset to default idle then toggle bits for pressed
        self.hid_pad[:] = bytearray.fromhex("ff 0f 00 00")
        for b in pressed:
            self._set_bit(self.hid_pad, int(b))

        self.interface_buttons[:] = bytearray.fromhex("00 00 00 00")
        for b in pressed_iface:
            self._set_bit(self.interface_buttons, int(b))

    def reset_neutral(self) -> None:
        # Neutral = no buttons, touch cleared, neutral circle/cpp/interface
        self.hid_pad[:] = bytearray.fromhex("ff 0f 00 00")
        self.reset_touch()
        self.circle_pad[:] = bytearray.fromhex("ff f7 7f 00")
        self.cpp_state[:] = bytearray.fromhex("81 00 80 80")
        self.interface_buttons[:] = bytearray.fromhex("00 00 00 00")


class ThreeDSBackend:
    """
    Backend used by the app when Output Backend = 3DS Input Redirection.
    """
    def __init__(self, ip: str, port: int = 4950):
        self.ip = ip
        self.port = port
        self.client = ThreeDSClient(ip=ip, port=port)
        self._connected = True  # UDP is stateless; treat as enabled

        self._pressed: set[ThreeDSButton] = set()
        self._pressed_iface: set[ThreeDSInterfaceButton] = set()

    @property
    def connected(self) -> bool:
        return self._connected

    def connect(self):
        self._connected = True

    def disconnect(self):
        # Best effort neutral on disconnect
        try:
            self._select_active_backend()
            if self.active_backend and getattr(self.active_backend, "connected", False):
                # prefer backend reset if available; else neutral buttons
                if hasattr(self.active_backend, "reset_neutral"):
                    self.active_backend.reset_neutral()
                else:
                    self.active_backend.set_buttons([])

        except Exception:
            pass
        self._connected = False

    def set_buttons(self, buttons: list[str]):
        """
        buttons are app-level names, e.g. ["A","Up","L"] etc.
        """
        mapping = {
            "A": ThreeDSButton.A, "B": ThreeDSButton.B, "X": ThreeDSButton.X, "Y": ThreeDSButton.Y,
            "Up": ThreeDSButton.UP, "Down": ThreeDSButton.DOWN, "Left": ThreeDSButton.LEFT, "Right": ThreeDSButton.RIGHT,
            "L": ThreeDSButton.L, "R": ThreeDSButton.R,
            "Start": ThreeDSButton.START, "Select": ThreeDSButton.SELECT,
            "ZL": ThreeDSButton.ZL, "ZR": ThreeDSButton.ZR,
        }
        pressed = set()
        for b in buttons:
            if b in mapping:
                pressed.add(mapping[b])

        self._pressed = pressed
        self.client.set_buttons_held(self._pressed, self._pressed_iface)
        self.client.send_update()

    def tap_touch(self, x_px: int, y_px: int, down_time: float = 0.1, settle: float = 0.1):
        # Down
        self.client.press_touch(int(x_px), int(y_px))
        self.client.send_update()
        if down_time > 0:
            precise_sleep(float(down_time))
        # Up
        self.client.reset_touch()
        self.client.send_update()
        if settle > 0:
            precise_sleep(float(settle))

    def reset_neutral(self):
        self.client.reset_neutral()
        self.client.send_update()

