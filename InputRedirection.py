from typing import Optional
from dataclasses import dataclass, field
import socket
import struct
from enum import IntEnum
import time
import math
# -----------------------------
# 3DS Input Redirection backend
# -----------------------------

TOUCH_SCREEN_WIDTH = 320
TOUCH_SCREEN_HEIGHT = 240
HID_AXIS_MAX = 0xFFF
CPAD_BOUND = 0x5D0
CPP_BOUND = 0x7F
SQRT_1_2 = math.sqrt(0.5)

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

def clamp(v: float, lo: float, hi: float) -> float:
    return max(lo, min(hi, v))

class InputRedirectionButton(IntEnum):
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

class InputRedirectionIrButton(IntEnum):
    ZR = 1
    ZL = 2

class InputRedirectionInterfaceButton(IntEnum):
    HOME = 0
    POWER = 1
    POWER_LONG = 2

@dataclass
class TouchState:
    pressed: bool = False
    x: int = 0
    y: int = 0

@dataclass
class StickState:
    x: float = 0.0
    y: float = 0.0

@dataclass
class InputRedirectionClient:
    ip: str
    port: int = 4950

    _sock: socket.socket = field(init=False, repr=False)
    _addr: tuple = field(init=False, repr=False)

    _buttons: set[InputRedirectionButton] = field(default_factory=set, repr=False)
    _ir_buttons: set[InputRedirectionIrButton] = field(default_factory=set, repr=False)
    _interface_buttons: set[InputRedirectionInterfaceButton] = field(default_factory=set, repr=False)
    _circle_pad: StickState = field(default_factory=StickState, repr=False)
    _c_stick: StickState = field(default_factory=StickState, repr=False)
    _touch: TouchState = field(default_factory=TouchState, repr=False)

    def __post_init__(self) -> None:
        self._addr = (self.ip, self.port)
        self._sock = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
        self._sock.setblocking(False)

    def set_buttons_held(
        self,
        pressed: set[InputRedirectionButton],
        pressed_iface: Optional[set[InputRedirectionInterfaceButton]] = None,
        pressed_ir: Optional[set[InputRedirectionIrButton]] = None,
    ) -> None:
        self._buttons = set(pressed or [])
        self._interface_buttons = set(pressed_iface or [])
        self._ir_buttons = set(pressed_ir or [])

    def set_ir_buttons_held(self, pressed: set[InputRedirectionIrButton]) -> None:
        self._ir_buttons = set(pressed or [])

    def set_interface_buttons_held(self, pressed: set[InputRedirectionInterfaceButton]) -> None:
        self._interface_buttons = set(pressed or [])

    def set_circle_pad(self, x: float, y: float) -> None:
        self._circle_pad.x = clamp(float(x), -1.0, 1.0)
        self._circle_pad.y = clamp(float(y), -1.0, 1.0)

    def reset_circle_pad(self) -> None:
        self._circle_pad.x = 0.0
        self._circle_pad.y = 0.0

    def set_c_stick(self, x: float, y: float) -> None:
        self._c_stick.x = clamp(float(x), -1.0, 1.0)
        self._c_stick.y = clamp(float(y), -1.0, 1.0)

    def reset_c_stick(self) -> None:
        self._c_stick.x = 0.0
        self._c_stick.y = 0.0

    def press_touch(self, x_px: int, y_px: int) -> None:
        self._touch.pressed = True
        self._touch.x = max(0, min(TOUCH_SCREEN_WIDTH - 1, int(x_px)))
        self._touch.y = max(0, min(TOUCH_SCREEN_HEIGHT - 1, int(y_px)))

    def reset_touch(self) -> None:
        self._touch.pressed = False
        self._touch.x = 0
        self._touch.y = 0

    def reset_neutral(self) -> None:
        self._buttons.clear()
        self._ir_buttons.clear()
        self._interface_buttons.clear()
        self.reset_circle_pad()
        self.reset_c_stick()
        self.reset_touch()

    def _encode_hid_pad(self) -> int:
        hid_pad = 0xFFF
        mask = 0
        for b in self._buttons:
            mask |= (1 << int(b))
        hid_pad &= ~mask
        return hid_pad

    def _encode_touch_screen(self) -> int:
        if not self._touch.pressed:
            return 0x2000000

        x = (HID_AXIS_MAX * self._touch.x) // TOUCH_SCREEN_WIDTH
        y = (HID_AXIS_MAX * self._touch.y) // TOUCH_SCREEN_HEIGHT
        return (1 << 24) | (y << 12) | x

    def _encode_circle_pad(self) -> int:
        if self._circle_pad.x == 0.0 and self._circle_pad.y == 0.0:
            return 0x7FF7FF

        x = int(self._circle_pad.x * CPAD_BOUND + 0x800)
        y = int(self._circle_pad.y * CPAD_BOUND + 0x800)

        if x >= 0xFFF:
            x = 0x000 if self._circle_pad.x < 0 else 0xFFF
        if y >= 0xFFF:
            y = 0x000 if self._circle_pad.y < 0 else 0xFFF

        x = max(0, x)
        y = max(0, y)
        return (y << 12) | x

    def _encode_cpp_state(self) -> int:
        ir_mask = 0
        for b in self._ir_buttons:
            ir_mask |= (1 << int(b))

        if self._c_stick.x == 0.0 and self._c_stick.y == 0.0 and ir_mask == 0:
            return 0x80800081

        rx = self._c_stick.x
        ry = self._c_stick.y
        rotated_x = SQRT_1_2 * (rx + ry)
        rotated_y = SQRT_1_2 * (ry - rx)

        x = int(rotated_x * CPP_BOUND + 0x80)
        y = int(rotated_y * CPP_BOUND + 0x80)

        if x >= 0xFF:
            x = 0x00 if rotated_x < 0 else 0xFF
        if y >= 0xFF:
            y = 0x00 if rotated_y < 0 else 0xFF

        x = max(0, min(0xFF, x))
        y = max(0, min(0xFF, y))

        return (y << 24) | (x << 16) | (ir_mask << 8) | 0x81

    def _encode_interface_buttons(self) -> int:
        mask = 0
        for b in self._interface_buttons:
            mask |= (1 << int(b))
        return mask

    def _build_packet(self) -> bytes:
        hid_pad = self._encode_hid_pad()
        touch_screen = self._encode_touch_screen()
        circle_pad = self._encode_circle_pad()
        cpp_state = self._encode_cpp_state()
        interface_buttons = self._encode_interface_buttons()

        return struct.pack("<IIIII", hid_pad, touch_screen, circle_pad, cpp_state, interface_buttons)

    def send_update(self) -> None:
        packet = self._build_packet()
        self._sock.sendto(packet, self._addr)


class InputRedirectionBackend:
    """
    Backend used by the app when Output Backend = 3DS Input Redirection.
    """
    def __init__(self, ip: str, port: int = 4950):
        self.ip = ip
        self.port = port
        self.client = InputRedirectionClient(ip=ip, port=port)
        self._connected = True  # UDP is stateless; treat as enabled

        self._pressed: set[InputRedirectionButton] = set()
        self._pressed_iface: set[InputRedirectionInterfaceButton] = set()
        self._pressed_ir: set[InputRedirectionIrButton] = set()

    @property
    def connected(self) -> bool:
        return self._connected

    def connect(self):
        self._connected = True

    def disconnect(self):
        # Best effort neutral on disconnect
        try:
            self.reset_neutral()
        except Exception:
            pass
        self._connected = False

    def set_buttons(self, buttons: list[str]):
        """
        buttons are app-level names, e.g. ["A","Up","L"] etc.
        """
        mapping = {
            "A": InputRedirectionButton.A, "B": InputRedirectionButton.B,
            "X": InputRedirectionButton.X, "Y": InputRedirectionButton.Y,
            "Up": InputRedirectionButton.UP, "Down": InputRedirectionButton.DOWN,
            "Left": InputRedirectionButton.LEFT, "Right": InputRedirectionButton.RIGHT,
            "L": InputRedirectionButton.L, "R": InputRedirectionButton.R,
            "Start": InputRedirectionButton.START, "Select": InputRedirectionButton.SELECT,
        }
        pressed = set()
        for b in buttons:
            if b in mapping:
                pressed.add(mapping[b])

        self._pressed = pressed
        self.client.set_buttons_held(self._pressed, self._pressed_iface, self._pressed_ir)
        self.client.send_update()

    def set_ir_buttons(self, buttons: list[str]):
        mapping = {
            "ZL": InputRedirectionIrButton.ZL,
            "ZR": InputRedirectionIrButton.ZR,
        }
        pressed_ir = set()
        for b in buttons:
            if b in mapping:
                pressed_ir.add(mapping[b])

        self._pressed_ir = pressed_ir
        self.client.set_buttons_held(self._pressed, self._pressed_iface, self._pressed_ir)
        self.client.send_update()

    def set_interface_buttons(self, buttons: list[str]):
        mapping = {
            "Home": InputRedirectionInterfaceButton.HOME,
            "Power": InputRedirectionInterfaceButton.POWER,
            "PowerLong": InputRedirectionInterfaceButton.POWER_LONG,
            "POWER_LONG": InputRedirectionInterfaceButton.POWER_LONG,
            "Power_Long": InputRedirectionInterfaceButton.POWER_LONG,
            "Power Long": InputRedirectionInterfaceButton.POWER_LONG,
        }
        pressed_iface = set()
        for b in buttons:
            if b in mapping:
                pressed_iface.add(mapping[b])

        self._pressed_iface = pressed_iface
        self.client.set_buttons_held(self._pressed, self._pressed_iface, self._pressed_ir)
        self.client.send_update()

    def set_circle_pad(self, x: float, y: float):
        self.client.set_circle_pad(x, y)
        self.client.send_update()

    def reset_circle_pad(self):
        self.client.reset_circle_pad()
        self.client.send_update()

    def set_c_stick(self, x: float, y: float):
        self.client.set_c_stick(x, y)
        self.client.send_update()

    def reset_c_stick(self):
        self.client.reset_c_stick()
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
        self._pressed.clear()
        self._pressed_iface.clear()
        self._pressed_ir.clear()
        self.client.reset_neutral()
        self.client.send_update()
