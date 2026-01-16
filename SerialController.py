"""
Serial controller backends for USB-connected controllers.

Supports two protocols:
- USB TX: Simple 3-byte packets (header + button bytes). Used by basic Arduino
  controllers that relay button state over RF to the console.
- PABotBase: Bidirectional protocol with acknowledgments, used by Pokémon
  automation hardware (Teensy-based). Supports timed button presses natively.

The SerialController wrapper auto-detects which protocol a device speaks.
"""
import os
import serial
import time
import threading
import struct

import pabotbase_controller as pabotbase
from serial.tools import list_ports

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
LEFT_STICK_BINDINGS = ["Left Stick Up", "Left Stick Down", "Left Stick Left", "Left Stick Right"]
RIGHT_STICK_BINDINGS = ["Right Stick Up", "Right Stick Down", "Right Stick Left", "Right Stick Right"]
KEYBIND_TARGETS = ALL_BUTTONS + LEFT_STICK_BINDINGS + RIGHT_STICK_BINDINGS

USB_TX_HEADER_BYTES = {0x43, 0x54}
PABOTBASE_HINTS = ("pabot", "pokemon", "teensy", "pjrc")

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


def bytes_to_buttons(high, low):
    buttons = []
    for name in ALL_BUTTONS:
        which, mask = BUTTON_MAP[name]
        if which == "high" and (high & mask):
            buttons.append(name)
        if which == "low" and (low & mask):
            buttons.append(name)
    return buttons


def _contains_usbtx_header(data):
    return any(b in USB_TX_HEADER_BYTES for b in data)


def _build_safe_seqnum_reset():
    for seqnum in range(1, 10000):
        seq_bytes = struct.pack("<I", seqnum)
        if _contains_usbtx_header(seq_bytes):
            continue
        message = pabotbase.PABotBaseMessage(pabotbase.MessageType.SEQNUM_RESET, seq_bytes)
        encoded = message.encode()
        if not _contains_usbtx_header(encoded):
            return seqnum, encoded
    return None, None


class UsbTxSerialBackend:
    def __init__(self, status_cb=None, app=None):
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

        ser = serial.Serial()
        ser.port = port
        ser.baudrate = baud
        ser.timeout = 1
        try:
            ser.dtr = False
            ser.rts = False
        except Exception:
            pass
        ser.open()
        try:
            ser.dtr = False
            ser.rts = False
        except Exception:
            pass
        self.ser = ser
        self.status_cb(f"USB TX serial connected: {port} @ {baud}")

        self._running = True
        self._thread = threading.Thread(target=self._keepalive_loop, daemon=True)
        self._thread.start()

        # Pairing warm-up (neutral for ~1s)
        self.status_cb("Pairing warm-up: neutral for ~1 second...")
        self.set_state(0, 0)
        time.sleep(1.0)
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
        self.status_cb("USB TX serial disconnected.")

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
            if not self.connected:
                break
            if self.app and getattr(self.app, "backend_var", None):
                if self.app.backend_var.get() != "USB Serial":
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


class PABotBaseSerialBackend:
    backend_name = "PABotBase"
    supports_timed_press = True

    def __init__(self, status_cb=None, app=None):
        self.status_cb = status_cb or (lambda s: None)
        self.app = app or (lambda s: None)
        self.ser = None
        self.controller = None

        self.interval_s = 0.01
        self._hold_duration_ms = 1000
        self._refresh_margin_s = 0.05

        self._lock = threading.Lock()
        self._io_lock = threading.Lock()
        self._send_event = threading.Event()
        self._running = False
        self._thread = None

        self._buttons = pabotbase.Button.NONE
        self._dpad = pabotbase.DPad.NONE
        self._left_x = pabotbase.STICK_CENTER
        self._left_y = pabotbase.STICK_CENTER
        self._right_x = pabotbase.STICK_CENTER
        self._right_y = pabotbase.STICK_CENTER
        self._last_sent_buttons = pabotbase.Button.NONE
        self._last_sent_dpad = pabotbase.DPad.NONE
        self._last_sent_left_x = pabotbase.STICK_CENTER
        self._last_sent_left_y = pabotbase.STICK_CENTER
        self._last_sent_right_x = pabotbase.STICK_CENTER
        self._last_sent_right_y = pabotbase.STICK_CENTER
        self._last_sent_at = 0.0
        self._last_sent_duration_ms = 0
        self._command_timeout_s = 0.2
        self._use_interrupt_on_change = True
        self._wait_for_ack = False

    @property
    def connected(self):
        return self.ser is not None and self.ser.is_open and self.controller is not None

    def _buttons_to_state(self, buttons):
        btns = pabotbase.Button.NONE
        dpad = pabotbase.DPad.NONE

        mapping = {
            "A": pabotbase.Button.A,
            "B": pabotbase.Button.B,
            "X": pabotbase.Button.X,
            "Y": pabotbase.Button.Y,
            "L": pabotbase.Button.L,
            "R": pabotbase.Button.R,
            "Start": pabotbase.Button.PLUS,
            "Select": pabotbase.Button.MINUS,
            "ZL": pabotbase.Button.ZL,
            "ZR": pabotbase.Button.ZR,
            "Home": pabotbase.Button.HOME,
            "Capture": pabotbase.Button.CAPTURE,
        }

        up = "Up" in buttons
        down = "Down" in buttons
        left = "Left" in buttons
        right = "Right" in buttons

        vertical = None
        if up and not down:
            vertical = "up"
        elif down and not up:
            vertical = "down"

        horizontal = None
        if right and not left:
            horizontal = "right"
        elif left and not right:
            horizontal = "left"

        if vertical == "up" and horizontal == "right":
            dpad = pabotbase.DPad.UP_RIGHT
        elif vertical == "up" and horizontal == "left":
            dpad = pabotbase.DPad.UP_LEFT
        elif vertical == "down" and horizontal == "right":
            dpad = pabotbase.DPad.DOWN_RIGHT
        elif vertical == "down" and horizontal == "left":
            dpad = pabotbase.DPad.DOWN_LEFT
        elif vertical == "up":
            dpad = pabotbase.DPad.UP
        elif vertical == "down":
            dpad = pabotbase.DPad.DOWN
        elif horizontal == "right":
            dpad = pabotbase.DPad.RIGHT
        elif horizontal == "left":
            dpad = pabotbase.DPad.LEFT

        for b in buttons:
            if b in mapping:
                btns |= mapping[b]

        return btns, dpad

    def _stick_axis_to_byte(self, v):
        try:
            value = float(v)
        except (TypeError, ValueError):
            value = 0.0
        value = max(-1.0, min(1.0, value))
        return int(round((value + 1.0) * 127.5))

    def _drain_serial_input_locked(self):
        if not self.ser:
            return
        try:
            waiting = self.ser.in_waiting
            if waiting:
                self.ser.read(waiting)
        except Exception:
            pass

    def _send_state(self, buttons, dpad, left_x, left_y, right_x, right_y, duration_ms, interrupt_next=False):
        if not self.connected:
            return False
        duration_ms = int(round(duration_ms))
        duration_ms = max(1, min(duration_ms, 65535))
        try:
            with self._io_lock:
                if self._wait_for_ack:
                    responses = []
                    ok = self.controller.send_controller_state(
                        buttons=buttons,
                        dpad=dpad,
                        left_x=left_x,
                        left_y=left_y,
                        right_x=right_x,
                        right_y=right_y,
                        duration_ms=duration_ms,
                        timeout_s=self._command_timeout_s,
                        debug_cb=lambda resp: responses.append(resp),
                        wait_for_ack=True,
                    )
                    if not ok:
                        if responses:
                            last = responses[-1]
                            payload_len = len(last.payload) if last.payload is not None else 0
                            self.status_cb(
                                f"PABotBase command ack not received "
                                f"(last type=0x{last.msg_type:02X} payload_len={payload_len})."
                            )
                        else:
                            self.status_cb("PABotBase command ack not received (no responses).")
                else:
                    ok = self.controller.send_controller_state(
                        buttons=buttons,
                        dpad=dpad,
                        left_x=left_x,
                        left_y=left_y,
                        right_x=right_x,
                        right_y=right_y,
                        duration_ms=duration_ms,
                        wait_for_ack=False,
                    )
                if ok and interrupt_next:
                    self.controller.interrupt_next_command(wait_for_ack=self._wait_for_ack)
                    ok = True
                if not self._wait_for_ack:
                    self._drain_serial_input_locked()
            return ok
        except Exception as e:
            self.status_cb(f"PABotBase send error: {e}")
            return False

    def connect(self, port, baud=pabotbase.BAUD_RATE):
        if self.connected:
            self.disconnect()

        self.ser = serial.Serial(port, baud, timeout=0)
        self.controller = pabotbase.PABotBaseController(self.ser)
        if not self.controller.connect():
            try:
                self.ser.close()
            except Exception:
                pass
            self.ser = None
            self.controller = None
            raise RuntimeError("PABotBase device not detected on this port.")

        self.status_cb(f"PABotBase serial connected: {port} @ {baud}")

        self._running = True
        self._thread = threading.Thread(target=self._keepalive_loop, daemon=True)
        self._thread.start()

        # Start from neutral
        self.set_buttons([])

    def disconnect(self):
        self._running = False
        self._send_event.set()
        if self._thread and self._thread.is_alive():
            self._thread.join(timeout=1.0)

        if self.controller:
            try:
                self.controller.stop_all_commands()
            except Exception:
                pass
            try:
                self.controller.reset_to_neutral()
            except Exception:
                pass

        if self.ser:
            try:
                self.ser.close()
            except Exception:
                pass
        self.ser = None
        self.controller = None
        self.status_cb("PABotBase serial disconnected.")

    def set_state(self, high, low):
        buttons = bytes_to_buttons(high, low)
        self.set_buttons(buttons)

    def set_buttons(self, buttons):
        if not self.connected:
            return
        btns, dpad = self._buttons_to_state(buttons)
        with self._lock:
            state_changed = (btns != self._buttons) or (dpad != self._dpad)
            self._buttons = btns
            self._dpad = dpad
        if state_changed:
            self._send_event.set()

    def set_left_stick(self, x, y):
        if not self.connected:
            return
        left_x = self._stick_axis_to_byte(x)
        left_y = self._stick_axis_to_byte(-y)
        with self._lock:
            state_changed = (left_x != self._left_x) or (left_y != self._left_y)
            self._left_x = left_x
            self._left_y = left_y
        if state_changed:
            self._send_event.set()

    def reset_left_stick(self):
        self.set_left_stick(0.0, 0.0)

    def set_right_stick(self, x, y):
        if not self.connected:
            return
        right_x = self._stick_axis_to_byte(x)
        right_y = self._stick_axis_to_byte(-y)
        with self._lock:
            state_changed = (right_x != self._right_x) or (right_y != self._right_y)
            self._right_x = right_x
            self._right_y = right_y
        if state_changed:
            self._send_event.set()

    def reset_right_stick(self):
        self.set_right_stick(0.0, 0.0)

    def press_buttons(self, buttons, duration_ms):
        if not self.connected:
            return
        btns, dpad = self._buttons_to_state(buttons)
        with self._lock:
            left_x = self._left_x
            left_y = self._left_y
            right_x = self._right_x
            right_y = self._right_y
        self._send_state(btns, dpad, left_x, left_y, right_x, right_y, duration_ms)

    def send_channel_set(self, channel_byte):
        raise RuntimeError("Channel set is not supported on PABotBase devices.")

    def reset_neutral(self):
        if not self.connected:
            return
        with self._lock:
            state_changed = (
                self._buttons != pabotbase.Button.NONE
                or self._dpad != pabotbase.DPad.NONE
                or self._left_x != pabotbase.STICK_CENTER
                or self._left_y != pabotbase.STICK_CENTER
                or self._right_x != pabotbase.STICK_CENTER
                or self._right_y != pabotbase.STICK_CENTER
            )
            self._buttons = pabotbase.Button.NONE
            self._dpad = pabotbase.DPad.NONE
            self._left_x = pabotbase.STICK_CENTER
            self._left_y = pabotbase.STICK_CENTER
            self._right_x = pabotbase.STICK_CENTER
            self._right_y = pabotbase.STICK_CENTER
        if state_changed:
            self._send_event.set()

    def _keepalive_loop(self):
        while self._running:
            if not self.connected:
                break
            if self.app and getattr(self.app, "backend_var", None):
                if self.app.backend_var.get() != "USB Serial":
                    break

            self._send_event.wait(self.interval_s)
            self._send_event.clear()

            if not self.connected:
                break

            now = time.monotonic()
            with self._lock:
                btns = self._buttons
                dpad = self._dpad
                left_x = self._left_x
                left_y = self._left_y
                right_x = self._right_x
                right_y = self._right_y
                last_btns = self._last_sent_buttons
                last_dpad = self._last_sent_dpad
                last_left_x = self._last_sent_left_x
                last_left_y = self._last_sent_left_y
                last_right_x = self._last_sent_right_x
                last_right_y = self._last_sent_right_y
                last_sent_at = self._last_sent_at
                last_duration_ms = self._last_sent_duration_ms

            state_changed = (
                btns != last_btns
                or dpad != last_dpad
                or left_x != last_left_x
                or left_y != last_left_y
                or right_x != last_right_x
                or right_y != last_right_y
            )
            send_needed = False
            send_interrupt = False

            elapsed = 0.0
            duration_s = 0.0
            if last_sent_at > 0 and last_duration_ms > 0:
                elapsed = now - last_sent_at
                duration_s = last_duration_ms / 1000.0
            active = duration_s > 0 and elapsed < duration_s

            refresh_due = False
            has_input = (
                btns != pabotbase.Button.NONE
                or dpad != pabotbase.DPad.NONE
                or left_x != pabotbase.STICK_CENTER
                or left_y != pabotbase.STICK_CENTER
                or right_x != pabotbase.STICK_CENTER
                or right_y != pabotbase.STICK_CENTER
            )
            if has_input:
                if last_sent_at == 0.0:
                    refresh_due = True
                else:
                    refresh_threshold = max(0.0, duration_s - self._refresh_margin_s)
                    if elapsed >= refresh_threshold:
                        refresh_due = True

            if state_changed:
                send_needed = True
                send_interrupt = self._use_interrupt_on_change and active
            elif refresh_due:
                send_needed = True

            if send_needed:
                ok = self._send_state(
                    btns,
                    dpad,
                    left_x,
                    left_y,
                    right_x,
                    right_y,
                    self._hold_duration_ms,
                    interrupt_next=send_interrupt,
                )
                if ok:
                    with self._lock:
                        self._last_sent_buttons = btns
                        self._last_sent_dpad = dpad
                        self._last_sent_left_x = left_x
                        self._last_sent_left_y = left_y
                        self._last_sent_right_x = right_x
                        self._last_sent_right_y = right_y
                        self._last_sent_at = time.monotonic()
                        self._last_sent_duration_ms = self._hold_duration_ms


class SerialController:
    def __init__(self, status_cb=None, app=None):
        self.status_cb = status_cb or (lambda s: None)
        self.app = app or (lambda s: None)
        self.backend = None
        self.usbtx_serial = None
        self.pabotbase_serial = None
        self._debug_enabled = True

    def _open_serial_no_reset(self, port, baud, timeout=0):
        ser = serial.Serial()
        ser.port = port
        ser.baudrate = baud
        ser.timeout = timeout
        try:
            ser.dtr = False
            ser.rts = False
        except Exception:
            pass
        ser.open()
        try:
            ser.dtr = False
            ser.rts = False
        except Exception:
            pass
        return ser

    def _debug(self, msg):
        if not self._debug_enabled:
            return
        try:
            self.status_cb(msg)
        except Exception:
            pass

    def _get_port_info(self, port):
        try:
            for info in list_ports.comports():
                if info.device == port:
                    return info
        except Exception:
            return None
        return None

    def _format_port_info(self, info):
        if info is None:
            return "unknown device"
        parts = []
        if getattr(info, "description", None):
            parts.append(info.description)
        if getattr(info, "manufacturer", None):
            parts.append(info.manufacturer)
        if getattr(info, "product", None):
            parts.append(info.product)
        if getattr(info, "hwid", None):
            parts.append(info.hwid)
        return " | ".join(parts) if parts else "unknown device"

    def _should_probe_pabotbase(self, port):
        override = os.environ.get("CMR_PABOTBASE_PROBE", "").strip().lower()
        if override in ("0", "false", "no", "off", "skip", "passive"):
            return False, "env=skip"
        if override in ("1", "true", "yes", "on", "force", "active"):
            return True, "env=force"

        info = self._get_port_info(port)
        if info is None:
            return True, "safe probe (no port info)"

        haystack = " ".join(
            str(x).lower()
            for x in (
                getattr(info, "description", ""),
                getattr(info, "manufacturer", ""),
                getattr(info, "product", ""),
                getattr(info, "hwid", ""),
            )
        )
        if any(hint in haystack for hint in PABOTBASE_HINTS):
            return True, "hint match"

        return True, "safe probe (default)"

    def _read_pabotbase_response(self, ser, timeout_s, max_log_bytes=64, accept_fn=None):
        start = time.time()
        buffer = bytearray()
        raw_log = bytearray()

        while (time.time() - start) < timeout_s:
            waiting = ser.in_waiting
            if waiting:
                data = ser.read(waiting)
                buffer.extend(data)
                if len(raw_log) < max_log_bytes:
                    raw_log.extend(data[: max_log_bytes - len(raw_log)])

            for i in range(len(buffer)):
                length_inverted = buffer[i]
                expected_length = (~length_inverted) & 0xFF

                if expected_length == 0 or expected_length > pabotbase.MAX_PACKET_SIZE:
                    continue
                if i + expected_length <= len(buffer):
                    msg_data = bytes(buffer[i:i + expected_length])
                    message = pabotbase.PABotBaseMessage.decode(msg_data)
                    if message is not None:
                        del buffer[:i + expected_length]
                        if accept_fn is None or accept_fn(message):
                            return message, bytes(raw_log)
                        break

            time.sleep(0.001)

        return None, bytes(raw_log)

    @property
    def connected(self):
        return self.backend is not None and self.backend.connected

    @property
    def supports_timed_press(self):
        return bool(getattr(self.backend, "supports_timed_press", False))

    def _probe_pabotbase(self, port, timeout_s=0.25, allow_write=False):
        seqnum, message = _build_safe_seqnum_reset()
        if message is None:
            self._debug("PABotBase probe: failed to build safe SEQNUM_RESET.")
            return False

        ser = None
        try:
            self._debug(f"PABotBase probe: opening {port} @ {pabotbase.BAUD_RATE}.")
            ser = self._open_serial_no_reset(port, pabotbase.BAUD_RATE, timeout=0)
            try:
                ser.reset_input_buffer()
                ser.reset_output_buffer()
            except Exception:
                pass

            self._debug(f"PABotBase probe: using seqnum={seqnum}, packet_len={len(message)}.")
            ack_types = {
                pabotbase.MessageType.ACK_REQUEST,
                pabotbase.MessageType.ACK_REQUEST_I8,
                pabotbase.MessageType.ACK_REQUEST_I16,
                pabotbase.MessageType.ACK_REQUEST_I32,
                pabotbase.MessageType.ACK_REQUEST_DATA,
            }

            def accept_passive(resp):
                if resp.msg_type == pabotbase.MessageType.ERROR_READY:
                    return True
                if resp.msg_type in ack_types:
                    return len(resp.payload) >= 4
                return False

            def accept_active(resp):
                if resp.msg_type not in ack_types:
                    return False
                if len(resp.payload) < 4:
                    return False
                ack_seqnum = struct.unpack("<I", resp.payload[:4])[0]
                if ack_seqnum != seqnum:
                    self._debug(f"PABotBase probe: ack seqnum mismatch ({ack_seqnum}).")
                    return False
                return True

            if not allow_write:
                self._debug("PABotBase probe: passive listen only (no writes).")
                response, raw_log = self._read_pabotbase_response(
                    ser,
                    timeout_s,
                    accept_fn=accept_passive,
                )
                if response is None:
                    if raw_log:
                        hex_dump = " ".join(f"{b:02X}" for b in raw_log)
                        self._debug(f"PABotBase probe: response, raw={hex_dump}")
                        return True
                    else:
                        self._debug("PABotBase probe: no response bytes.")
                    return False

                self._debug(
                    f"PABotBase probe: response type=0x{response.msg_type:02X} "
                    f"payload_len={len(response.payload)}."
                )
                return True
            
            for attempt in range(2):
                if attempt:
                    time.sleep(0.1)
                self._debug(f"PABotBase probe: attempt {attempt + 1}/2 send.")
                ser.write(message)
                ser.flush()

                response, raw_log = self._read_pabotbase_response(
                    ser,
                    timeout_s,
                    accept_fn=accept_active,
                )
                if response is None:
                    if raw_log:
                        hex_dump = " ".join(f"{b:02X}" for b in raw_log)
                        self._debug(f"PABotBase probe: no valid response, raw={hex_dump}")
                    else:
                        self._debug("PABotBase probe: no response bytes.")
                    continue

                self._debug(
                    f"PABotBase probe: response type=0x{response.msg_type:02X} "
                    f"payload_len={len(response.payload)}."
                )
                ack_seqnum = struct.unpack("<I", response.payload[:4])[0]
                self._debug(f"PABotBase probe: ack seqnum={ack_seqnum}.")
                return True

            self._debug("PABotBase probe: no valid response.")
            return False
        except Exception as e:
            self._debug(f"PABotBase probe: error {e}.")
            return False
        finally:
            if ser:
                try:
                    ser.close()
                except Exception:
                    pass

    def connect(self, port, baud=1_000_000):
        if self.connected:
            self.disconnect()

        port_info = self._get_port_info(port)
        self._debug(f"Serial connect: probing {port}.")
        if port_info:
            self._debug(f"Serial connect: port info: {self._format_port_info(port_info)}")
        allow_write, reason = False, "default to no-write probe"
        if not allow_write:
            self._debug(f"Serial connect: PABotBase probe write disabled ({reason}).")
        if self._probe_pabotbase(port, allow_write=allow_write):
            self._debug("Serial connect: PABotBase detected.")
            self.pabotbase_serial = PABotBaseSerialBackend(status_cb=self.status_cb, app=self.app)
            self.pabotbase_serial.connect(port)
            self.backend = self.pabotbase_serial
            return

        self._debug("Serial connect: falling back to USB TX.")
        self.usbtx_serial = UsbTxSerialBackend(status_cb=self.status_cb, app=self.app)
        try:
            self.usbtx_serial.connect(port, baud=baud)
            self.backend = self.usbtx_serial
        except Exception:
            self.usbtx_serial = None
            raise

    def disconnect(self):
        if self.backend:
            self.backend.disconnect()
        self.backend = None
        self.usbtx_serial = None
        self.pabotbase_serial = None

    def set_state(self, high, low):
        if not self.backend:
            return
        if hasattr(self.backend, "set_state"):
            self.backend.set_state(high, low)
        else:
            self.set_buttons(bytes_to_buttons(high, low))

    def set_buttons(self, buttons):
        if not self.backend:
            return
        self.backend.set_buttons(buttons)

    def set_left_stick(self, x, y):
        if not self.backend:
            raise RuntimeError("Not connected.")
        if not hasattr(self.backend, "set_left_stick"):
            raise RuntimeError("Left stick is not supported by this backend.")
        self.backend.set_left_stick(x, y)

    def reset_left_stick(self):
        if not self.backend:
            raise RuntimeError("Not connected.")
        if hasattr(self.backend, "reset_left_stick"):
            self.backend.reset_left_stick()
            return
        if hasattr(self.backend, "set_left_stick"):
            self.backend.set_left_stick(0.0, 0.0)
            return
        raise RuntimeError("Left stick is not supported by this backend.")

    def set_right_stick(self, x, y):
        if not self.backend:
            raise RuntimeError("Not connected.")
        if not hasattr(self.backend, "set_right_stick"):
            raise RuntimeError("Right stick is not supported by this backend.")
        self.backend.set_right_stick(x, y)

    def reset_right_stick(self):
        if not self.backend:
            raise RuntimeError("Not connected.")
        if hasattr(self.backend, "reset_right_stick"):
            self.backend.reset_right_stick()
            return
        if hasattr(self.backend, "set_right_stick"):
            self.backend.set_right_stick(0.0, 0.0)
            return
        raise RuntimeError("Right stick is not supported by this backend.")

    def press_buttons(self, buttons, duration_ms):
        if not self.backend or not self.supports_timed_press:
            return
        self.backend.press_buttons(buttons, duration_ms)

    def send_channel_set(self, channel_byte):
        if not self.backend:
            raise RuntimeError("Not connected.")
        self.backend.send_channel_set(channel_byte)

    def reset_neutral(self):
        if not self.backend:
            return
        if hasattr(self.backend, "reset_neutral"):
            self.backend.reset_neutral()
        else:
            self.set_buttons([])
