#!/usr/bin/env python3
"""
PABotBase Controller - Python Implementation
From: https://github.com/PokemonAutomation/

This module provides a Python interface to control a Nintendo Switch via an
Arduino/Teensy running the PABotBase firmware. It implements the complete
PABotBase serial communication protocol including:

- CRC32C message checksumming for reliable transmission
- Button press and joystick control
- Connection management and device verification
- Command queueing and acknowledgment handling

The PABotBase protocol uses binary messages with the following format:
    byte 0: Message length (inverted bits)
    byte 1: Message type
    bytes 2-N: Message payload (varies by type)
    last 4 bytes: CRC32C checksum

Usage Example:
    import serial
    from pabotbase_controller import PABotBaseController, Button

    # Connect to Arduino on COM3 (Windows) or /dev/ttyACM0 (Linux)
    ser = serial.Serial('COM3', 115200, timeout=1)
    controller = PABotBaseController(ser)

    # Verify connection and firmware
    controller.connect()
    print(f"Protocol Version: {controller.get_protocol_version()}")
    print(f"Program Name: {controller.get_program_name()}")

    # Press the A button for 100ms
    controller.press_button(Button.A, duration_ms=100)

    # Move left joystick fully right for 500ms
    controller.move_joystick(left_x=0xFF, left_y=0x80, duration_ms=500)

    # Press A+B while moving joystick
    controller.send_controller_state(
        buttons=Button.A | Button.B,
        left_x=0xFF, left_y=0x80,
        duration_ms=200
    )
"""

import struct
import serial
import time
from enum import IntEnum, IntFlag
from typing import Optional, List, Tuple


# ============================================================================
# Protocol Constants
# ============================================================================

# Serial communication settings
BAUD_RATE = 115200  # PABotBase uses 115200 baud
PROTOCOL_OVERHEAD = 6  # 2 bytes header + 4 bytes CRC32C
MAX_PACKET_SIZE = 64  # Maximum packet size in bytes
DEVICE_MINIMUM_QUEUE_SIZE = 4  # Minimum command queue size

# Message timeout and retry settings
DEFAULT_TIMEOUT = 2.0  # seconds
RETRANSMIT_DELAY = 0.1  # seconds between retransmissions


# ============================================================================
# Button Constants
# ============================================================================
# Each button is represented as a bit flag in a 32-bit integer.
# Multiple buttons can be combined using bitwise OR (|)

class Button(IntFlag):
    """
    Nintendo Switch button flags.

    These represent the physical buttons on a Nintendo Switch controller.
    Multiple buttons can be pressed simultaneously by combining flags:
        Button.A | Button.B  # Press A and B together
    """
    NONE = 0
    Y = 1 << 0  # Face button Y
    B = 1 << 1  # Face button B
    A = 1 << 2  # Face button A
    X = 1 << 3  # Face button X
    L = 1 << 4  # Left shoulder button
    R = 1 << 5  # Right shoulder button
    ZL = 1 << 6  # Left trigger
    ZR = 1 << 7  # Right trigger
    MINUS = 1 << 8  # Minus button (-)
    PLUS = 1 << 9  # Plus button (+)
    LCLICK = 1 << 10  # Left stick click
    RCLICK = 1 << 11  # Right stick click
    HOME = 1 << 12  # Home button
    CAPTURE = 1 << 13  # Capture button
    GR = 1 << 14  # Additional button GR
    GL = 1 << 15  # Additional button GL

    # D-pad buttons (also available as directional pad byte)
    UP = 1 << 16  # D-pad up
    RIGHT = 1 << 17  # D-pad right
    DOWN = 1 << 18  # D-pad down
    LEFT = 1 << 19  # D-pad left

    # Additional Joy-Con specific buttons
    LEFT_SL = 1 << 20  # Left Joy-Con SL button
    LEFT_SR = 1 << 21  # Left Joy-Con SR button
    RIGHT_SL = 1 << 22  # Right Joy-Con SL button
    RIGHT_SR = 1 << 23  # Right Joy-Con SR button
    C = 1 << 24  # C button


# ============================================================================
# Joystick Constants
# ============================================================================

# Joystick analog values range from 0x00 to 0xFF
STICK_MIN = 0x00  # Fully left or down
STICK_CENTER = 0x80  # Neutral/center position
STICK_MAX = 0xFF  # Fully right or up


# ============================================================================
# D-Pad Constants
# ============================================================================

class DPad(IntEnum):
    """D-pad directional positions (used in some controller modes)."""
    UP = 0
    UP_RIGHT = 1
    RIGHT = 2
    DOWN_RIGHT = 3
    DOWN = 4
    DOWN_LEFT = 5
    LEFT = 6
    UP_LEFT = 7
    NONE = 8  # D-pad not pressed


# ============================================================================
# Protocol Message Types
# ============================================================================

class MessageType:
    """
    PABotBase protocol message type identifiers.

    Message types are categorized as:
    - Errors (0x00-0x0F): Framework error messages
    - Acks (0x10-0x1F): Acknowledgment responses
    - Info (0x20-0x3F): One-way informational messages
    - Requests (0x40-0x7F): Short requests requiring acknowledgment
    - Commands (0x80-0xFF): Long-running commands
    """

    # Framework Errors (0x00-0x0F)
    ERROR_READY = 0x00  # Device ready signal
    ERROR_INVALID_MESSAGE = 0x01  # Malformed message received
    ERROR_CHECKSUM_MISMATCH = 0x02  # CRC check failed
    ERROR_INVALID_TYPE = 0x03  # Unknown message type
    ERROR_INVALID_REQUEST = 0x04  # Invalid request
    ERROR_MISSED_REQUEST = 0x05  # Sequence number gap detected
    ERROR_COMMAND_DROPPED = 0x06  # Command queue full
    ERROR_WARNING = 0x07  # Warning condition
    ERROR_DISCONNECTED = 0x08  # Connection lost

    # Acknowledgments (0x10-0x1F)
    ACK_COMMAND = 0x10  # Acknowledge command received
    ACK_REQUEST = 0x11  # Acknowledge request (no data)
    ACK_REQUEST_I8 = 0x12  # Ack with 8-bit response
    ACK_REQUEST_I16 = 0x13  # Ack with 16-bit response
    ACK_REQUEST_I32 = 0x14  # Ack with 32-bit response
    ACK_REQUEST_DATA = 0x1F  # Ack with variable data

    # Info Messages (0x20-0x3F)
    INFO_I32 = 0x20  # Info message with 32-bit value
    INFO_DATA = 0x21  # Info message with data payload
    INFO_STRING = 0x23  # Info message with string
    INFO_I32_LABEL = 0x24  # Labeled 32-bit value
    INFO_H32_LABEL = 0x25  # Labeled 32-bit hex value
    INFO_DEVICE_RESET_REASON = 0x26  # Device reset notification

    # Requests (0x40-0x7F)
    SEQNUM_RESET = 0x40  # Reset sequence numbers
    REQUEST_PROTOCOL_VERSION = 0x41  # Query protocol version
    REQUEST_PROGRAM_VERSION = 0x42  # Query firmware version
    REQUEST_PROGRAM_ID = 0x43  # Query program ID
    REQUEST_PROGRAM_NAME = 0x44  # Query program name string
    REQUEST_CONTROLLER_LIST = 0x45  # Query supported controllers
    REQUEST_QUEUE_SIZE = 0x46  # Query command queue capacity
    REQUEST_READ_CONTROLLER_MODE = 0x47  # Get current controller type
    REQUEST_CHANGE_CONTROLLER_MODE = 0x48  # Switch controller type
    REQUEST_RESET_TO_CONTROLLER = 0x49  # Reset and switch controller
    REQUEST_COMMAND_FINISHED = 0x4A  # Command execution complete
    REQUEST_STOP = 0x4B  # Stop all commands
    REQUEST_NEXT_CMD_INTERRUPT = 0x4C  # Interrupt next command
    REQUEST_STATUS = 0x50  # Query device status

    # Commands (0x80-0xFF)
    # Command 0x90: Send Nintendo Switch controller state
    # This is the main command for controlling the Switch
    COMMAND_NS_WIRED_CONTROLLER_STATE = 0x90


# ============================================================================
# Controller IDs
# ============================================================================

class ControllerID(IntEnum):
    """Controller type identifiers for different emulation modes."""
    NONE = 0

    # Standard HID
    KEYBOARD = 0x0100

    # Nintendo Switch (Original)
    NS_WIRED_CONTROLLER = 0x1000
    NS_WIRED_PRO_CONTROLLER = 0x1100
    NS_WIRED_LEFT_JOYCON = 0x1101
    NS_WIRED_RIGHT_JOYCON = 0x1102
    NS_WIRELESS_PRO_CONTROLLER = 0x1180
    NS_WIRELESS_LEFT_JOYCON = 0x1181
    NS_WIRELESS_RIGHT_JOYCON = 0x1182

    # Nintendo Switch 2
    NS2_WIRED_CONTROLLER = 0x1010
    NS2_WIRED_PRO_CONTROLLER = 0x1200
    NS2_WIRED_LEFT_JOYCON = 0x1201
    NS2_WIRED_RIGHT_JOYCON = 0x1202
    NS2_WIRELESS_PRO_CONTROLLER = 0x1280
    NS2_WIRELESS_LEFT_JOYCON = 0x1281
    NS2_WIRELESS_RIGHT_JOYCON = 0x1282


# ============================================================================
# CRC32C Implementation
# ============================================================================

_CRC32C_TABLE = None


def _get_crc32c_table():
    global _CRC32C_TABLE
    if _CRC32C_TABLE is not None:
        return _CRC32C_TABLE

    table = []
    poly = 0x82F63B78  # reflected 0x1EDC6F41
    for i in range(256):
        crc = i
        for _ in range(8):
            if crc & 1:
                crc = (crc >> 1) ^ poly
            else:
                crc >>= 1
        table.append(crc & 0xFFFFFFFF)

    _CRC32C_TABLE = table
    return table


def calculate_crc32c(data: bytes) -> int:
    """
    Calculate CRC32C (Castagnoli) checksum for message verification.

    The PABotBase firmware uses CRC32C with an initial value of 0xFFFFFFFF and
    no final XOR.
    """
    table = _get_crc32c_table()
    crc = 0xFFFFFFFF
    for b in data:
        crc = table[(crc ^ b) & 0xFF] ^ (crc >> 8)
    return crc & 0xFFFFFFFF


# ============================================================================
# Message Encoding/Decoding
# ============================================================================

class PABotBaseMessage:
    """
    Represents a PABotBase protocol message.

    Message structure:
        [0]: Length byte (inverted)
        [1]: Message type
        [2:-4]: Payload data
        [-4:]: CRC32C checksum (little-endian)
    """

    def __init__(self, msg_type: int, payload: bytes = b''):
        """
        Create a new message.

        Args:
            msg_type: Message type identifier (0x00-0xFF)
            payload: Message payload bytes (optional)
        """
        self.msg_type = msg_type
        self.payload = payload

    def encode(self) -> bytes:
        """
        Encode message into bytes for transmission.

        Returns:
            Complete encoded message with length, type, payload, and CRC
        """
        # Calculate total message length (length byte + type + payload + CRC)
        total_length = 1 + 1 + len(self.payload) + 4

        # Build message without CRC
        msg_without_crc = bytes([
            (~total_length) & 0xFF,  # Inverted length byte
            self.msg_type  # Message type
        ]) + self.payload

        # Calculate and append CRC32C checksum
        crc = calculate_crc32c(msg_without_crc)
        crc_bytes = struct.pack('<I', crc)  # Little-endian 32-bit

        return msg_without_crc + crc_bytes

    @staticmethod
    def decode(data: bytes) -> Optional['PABotBaseMessage']:
        """
        Decode a received message.

        Args:
            data: Raw bytes received from serial port

        Returns:
            Decoded PABotBaseMessage or None if invalid
        """
        if len(data) < PROTOCOL_OVERHEAD:
            return None

        # Extract components
        length_inverted = data[0]
        expected_length = (~length_inverted) & 0xFF

        if len(data) != expected_length:
            return None

        msg_type = data[1]
        payload = data[2:-4]
        received_crc = struct.unpack('<I', data[-4:])[0]

        # Verify CRC
        calculated_crc = calculate_crc32c(data[:-4])
        if calculated_crc != received_crc:
            return None

        return PABotBaseMessage(msg_type, payload)


# ============================================================================
# Controller State
# ============================================================================

class ControllerState:
    """
    Represents the complete state of a Nintendo Switch controller.

    This includes all buttons, both joysticks, and the D-pad.
    The state is packed into a binary format matching the PABotBase protocol.
    """

    def __init__(self,
                 buttons: Button = Button.NONE,
                 dpad: DPad = DPad.NONE,
                 left_x: int = STICK_CENTER,
                 left_y: int = STICK_CENTER,
                 right_x: int = STICK_CENTER,
                 right_y: int = STICK_CENTER):
        """
        Initialize controller state.

        Args:
            buttons: Button flags (can be combined with |)
            dpad: D-pad position
            left_x: Left joystick X (0x00=left, 0x80=center, 0xFF=right)
            left_y: Left joystick Y (0x00=down, 0x80=center, 0xFF=up)
            right_x: Right joystick X
            right_y: Right joystick Y
        """
        self.buttons = buttons
        self.dpad = dpad
        self.left_x = max(0, min(255, left_x))
        self.left_y = max(0, min(255, left_y))
        self.right_x = max(0, min(255, right_x))
        self.right_y = max(0, min(255, right_y))

    def encode(self) -> bytes:
        """
        Encode controller state into binary format for transmission.

        Returns:
            7 bytes representing controller state:
                [0]: buttons0 (lower 8 bits of button flags)
                [1]: buttons1 (middle 8 bits of button flags)
                [2]: dpad_byte (D-pad position)
                [3]: left_joystick_x
                [4]: left_joystick_y
                [5]: right_joystick_x
                [6]: right_joystick_y
        """
        button_value = int(self.buttons)
        buttons0 = button_value & 0xFF  # Lower 8 bits
        buttons1 = (button_value >> 8) & 0xFF  # Middle 8 bits

        return struct.pack('BBBBBBB',
                          buttons0,
                          buttons1,
                          self.dpad,
                          self.left_x,
                          self.left_y,
                          self.right_x,
                          self.right_y)

    @staticmethod
    def neutral() -> 'ControllerState':
        """
        Create a neutral controller state (no buttons, centered joysticks).

        Returns:
            Neutral ControllerState
        """
        return ControllerState()


# ============================================================================
# Main Controller Class
# ============================================================================

class PABotBaseController:
    """
    Main interface for controlling a Nintendo Switch via PABotBase.

    This class manages the serial connection, handles the communication
    protocol, and provides high-level methods for button presses and
    joystick movements.
    """

    def __init__(self, serial_port: serial.Serial):
        """
        Initialize the PABotBase controller.

        Args:
            serial_port: Opened pyserial Serial object configured for 115200 baud
        """
        self.serial = serial_port
        self.seqnum = 1  # Message sequence number (increments for each request/command)
        self.connected = False

    def _send_message(self, message: PABotBaseMessage) -> None:
        """
        Send a message over serial.

        Args:
            message: PABotBaseMessage to send
        """
        data = message.encode()
        self.serial.write(data)
        self.serial.flush()

    def _receive_message(self, timeout: float = DEFAULT_TIMEOUT) -> Optional[PABotBaseMessage]:
        """
        Receive and decode a message from serial.

        This method attempts to find valid message boundaries by trying to
        parse each byte as a potential message start. This provides
        synchronization recovery if bytes are dropped.

        Args:
            timeout: Maximum time to wait for a message (seconds)

        Returns:
            Decoded PABotBaseMessage or None if timeout/invalid
        """
        start_time = time.time()
        buffer = bytearray()

        while (time.time() - start_time) < timeout:
            # Read available bytes
            if self.serial.in_waiting > 0:
                buffer.extend(self.serial.read(self.serial.in_waiting))

            # Try to decode from each position in buffer
            for i in range(len(buffer)):
                # Get expected length from inverted length byte
                if i < len(buffer):
                    length_inverted = buffer[i]
                    expected_length = (~length_inverted) & 0xFF

                    # Skip zero bytes (used for synchronization)
                    if expected_length == 0 or expected_length > MAX_PACKET_SIZE:
                        continue

                    # Check if we have enough bytes for a complete message
                    if i + expected_length <= len(buffer):
                        # Try to decode message
                        msg_data = bytes(buffer[i:i + expected_length])
                        message = PABotBaseMessage.decode(msg_data)

                        if message is not None:
                            # Valid message found - remove processed bytes
                            del buffer[:i + expected_length]
                            return message

            # Brief sleep to avoid busy-waiting
            time.sleep(0.001)

        return None

    def _send_request_and_wait(self, msg_type: int, payload: bytes = b'') -> Optional[PABotBaseMessage]:
        """
        Send a request and wait for acknowledgment.

        This implements the request-response pattern with automatic retransmission
        if no acknowledgment is received.

        Args:
            msg_type: Request message type
            payload: Request payload (default includes seqnum)

        Returns:
            Acknowledgment message or None if failed
        """
        # Prepend sequence number to payload
        seqnum_bytes = struct.pack('<I', self.seqnum)
        full_payload = seqnum_bytes + payload

        message = PABotBaseMessage(msg_type, full_payload)

        # Try sending with retries
        for attempt in range(5):
            self._send_message(message)

            # Wait for acknowledgment
            response = self._receive_message(timeout=RETRANSMIT_DELAY * 2)

            if response is not None:
                # Check if it's an ack for our message
                if response.msg_type in [MessageType.ACK_REQUEST,
                                        MessageType.ACK_REQUEST_I8,
                                        MessageType.ACK_REQUEST_I16,
                                        MessageType.ACK_REQUEST_I32,
                                        MessageType.ACK_REQUEST_DATA]:
                    # Verify sequence number matches
                    if len(response.payload) >= 4:
                        ack_seqnum = struct.unpack('<I', response.payload[:4])[0]
                        if ack_seqnum == self.seqnum:
                            self.seqnum += 1
                            return response

            # Wait before retry
            time.sleep(RETRANSMIT_DELAY)

        return None

    def connect(self) -> bool:
        """
        Establish connection with the PABotBase device.

        This sends a sequence number reset and waits for the device to
        acknowledge, confirming bidirectional communication.

        Returns:
            True if connection successful, False otherwise
        """
        # Send sequence number reset
        response = self._send_request_and_wait(MessageType.SEQNUM_RESET)

        if response is not None:
            self.connected = True
            return True

        self.connected = False
        return False

    def get_protocol_version(self) -> Optional[int]:
        """
        Query the PABotBase protocol version.

        Returns:
            Protocol version number or None if query failed
        """
        response = self._send_request_and_wait(MessageType.REQUEST_PROTOCOL_VERSION)

        if response and response.msg_type == MessageType.ACK_REQUEST_I32:
            if len(response.payload) >= 8:  # seqnum + data
                return struct.unpack('<I', response.payload[4:8])[0]

        return None

    def get_program_version(self) -> Optional[int]:
        """
        Query the firmware program version.

        Returns:
            Program version number or None if query failed
        """
        response = self._send_request_and_wait(MessageType.REQUEST_PROGRAM_VERSION)

        if response and response.msg_type == MessageType.ACK_REQUEST_I32:
            if len(response.payload) >= 8:
                return struct.unpack('<I', response.payload[4:8])[0]

        return None

    def get_program_id(self) -> Optional[int]:
        """
        Query the program ID (identifies device type).

        Returns:
            Program ID (e.g., 0x01 for Arduino Uno) or None if query failed
        """
        response = self._send_request_and_wait(MessageType.REQUEST_PROGRAM_ID)

        if response and response.msg_type == MessageType.ACK_REQUEST_I8:
            if len(response.payload) >= 5:  # seqnum + data
                return response.payload[4]

        return None

    def get_program_name(self) -> Optional[str]:
        """
        Query the human-readable program name.

        Returns:
            Program name string or None if query failed
        """
        response = self._send_request_and_wait(MessageType.REQUEST_PROGRAM_NAME)

        if response and response.msg_type == MessageType.ACK_REQUEST_DATA:
            if len(response.payload) > 4:
                # Skip seqnum, decode remaining as UTF-8 string
                try:
                    return response.payload[4:].decode('utf-8', errors='ignore').rstrip('\x00')
                except:
                    return None

        return None

    def get_queue_size(self) -> Optional[int]:
        """
        Query the device's command queue capacity.

        Returns:
            Queue size (typically 4 or more) or None if query failed
        """
        response = self._send_request_and_wait(MessageType.REQUEST_QUEUE_SIZE)

        if response and response.msg_type == MessageType.ACK_REQUEST_I8:
            if len(response.payload) >= 5:
                return response.payload[4]

        return None

    def stop_all_commands(self, wait_for_ack: bool = True) -> bool:
        """
        Immediately stop all running commands and clear the queue.

        Args:
            wait_for_ack: If False, fire-and-forget and advance the seqnum.

        Returns:
            True if stop command acknowledged (or sent), False otherwise
        """
        if wait_for_ack:
            response = self._send_request_and_wait(MessageType.REQUEST_STOP)
            return response is not None

        seqnum_bytes = struct.pack('<I', self.seqnum)
        message = PABotBaseMessage(MessageType.REQUEST_STOP, seqnum_bytes)
        self._send_message(message)
        self.seqnum += 1
        return True

    def interrupt_next_command(self, wait_for_ack: bool = True) -> bool:
        """
        Interrupt the current command to advance to the next queued command.

        Args:
            wait_for_ack: If False, fire-and-forget and advance the seqnum.

        Returns:
            True if interrupt command acknowledged (or sent), False otherwise
        """
        if wait_for_ack:
            response = self._send_request_and_wait(MessageType.REQUEST_NEXT_CMD_INTERRUPT)
            return response is not None

        seqnum_bytes = struct.pack('<I', self.seqnum)
        message = PABotBaseMessage(MessageType.REQUEST_NEXT_CMD_INTERRUPT, seqnum_bytes)
        self._send_message(message)
        self.seqnum += 1
        return True

    def send_controller_state(self,
                              buttons: Button = Button.NONE,
                              dpad: DPad = DPad.NONE,
                              left_x: int = STICK_CENTER,
                              left_y: int = STICK_CENTER,
                              right_x: int = STICK_CENTER,
                              right_y: int = STICK_CENTER,
                              duration_ms: int = 100,
                              timeout_s: float = 0.5,
                              debug_cb=None,
                              wait_for_ack: bool = True) -> bool:
        """
        Send a controller state command to hold for a specified duration.

        This is the main method for controlling the Switch. It sets all button
        and joystick states and holds them for the specified time.

        Args:
            buttons: Button flags to press (combine with | operator)
            dpad: D-pad position
            left_x: Left joystick X position (0x00-0xFF)
            left_y: Left joystick Y position
            right_x: Right joystick X position
            right_y: Right joystick Y position
            duration_ms: Time to hold this state in milliseconds

        Returns:
            True if command sent successfully, False otherwise

        Example:
            # Press A button for 100ms
            controller.send_controller_state(buttons=Button.A, duration_ms=100)

            # Press A+B while moving left stick right
            controller.send_controller_state(
                buttons=Button.A | Button.B,
                left_x=0xFF,
                duration_ms=200
            )
        """
        state = ControllerState(buttons, dpad, left_x, left_y, right_x, right_y)

        # Build command payload:
        #   4 bytes: sequence number
        #   2 bytes: duration in milliseconds (little-endian)
        #   7 bytes: controller state
        seqnum_bytes = struct.pack('<I', self.seqnum)
        duration_bytes = struct.pack('<H', min(duration_ms, 65535))
        state_bytes = state.encode()

        payload = seqnum_bytes + duration_bytes + state_bytes

        message = PABotBaseMessage(MessageType.COMMAND_NS_WIRED_CONTROLLER_STATE, payload)

        # Send command and optionally wait for ack
        self._send_message(message)
        if not wait_for_ack:
            self.seqnum += 1
            return True
        end_time = time.time() + max(0.05, float(timeout_s))

        while time.time() < end_time:
            remaining = end_time - time.time()
            response = self._receive_message(timeout=min(0.05, max(0.0, remaining)))
            if response is None:
                continue

            if debug_cb:
                debug_cb(response)

            if response.msg_type == MessageType.ACK_COMMAND:
                if len(response.payload) >= 4:
                    ack_seqnum = struct.unpack('<I', response.payload[:4])[0]
                    if ack_seqnum == self.seqnum:
                        self.seqnum += 1
                        return True
                    # Mismatched seqnum; keep waiting for the right ack.
                continue

            if response.msg_type == MessageType.REQUEST_COMMAND_FINISHED:
                if len(response.payload) >= 4:
                    finished_seqnum = struct.unpack('<I', response.payload[:4])[0]
                    if finished_seqnum == self.seqnum:
                        self.seqnum += 1
                        return True

        return False

    def press_button(self, button: Button, duration_ms: int = 100) -> bool:
        """
        Press a button (or combination) for the specified duration.

        Args:
            button: Button or button combination to press
            duration_ms: Duration to hold button(s) in milliseconds

        Returns:
            True if successful, False otherwise

        Example:
            controller.press_button(Button.A)  # Press A for 100ms
            controller.press_button(Button.A | Button.B, 200)  # Press A+B for 200ms
        """
        return self.send_controller_state(buttons=button, duration_ms=duration_ms)

    def move_joystick(self,
                     left_x: int = STICK_CENTER,
                     left_y: int = STICK_CENTER,
                     right_x: int = STICK_CENTER,
                     right_y: int = STICK_CENTER,
                     duration_ms: int = 100) -> bool:
        """
        Move joystick(s) to specified position(s) for the specified duration.

        Args:
            left_x: Left stick X (0x00=left, 0x80=center, 0xFF=right)
            left_y: Left stick Y (0x00=down, 0x80=center, 0xFF=up)
            right_x: Right stick X
            right_y: Right stick Y
            duration_ms: Duration to hold position in milliseconds

        Returns:
            True if successful, False otherwise

        Example:
            # Move left stick fully right for 500ms
            controller.move_joystick(left_x=0xFF, duration_ms=500)

            # Move right stick diagonally up-right
            controller.move_joystick(right_x=0xFF, right_y=0xFF, duration_ms=300)
        """
        return self.send_controller_state(
            left_x=left_x,
            left_y=left_y,
            right_x=right_x,
            right_y=right_y,
            duration_ms=duration_ms
        )

    def reset_to_neutral(self, duration_ms: int = 50) -> bool:
        """
        Release all buttons and center all joysticks.

        Args:
            duration_ms: Duration to hold neutral state

        Returns:
            True if successful, False otherwise
        """
        return self.send_controller_state(duration_ms=duration_ms)


# ============================================================================
# Example Usage
# ============================================================================

def example_usage():
    """
    Demonstrates basic usage of the PABotBase controller.

    This function shows how to:
    1. Connect to the device
    2. Verify firmware version
    3. Send button presses
    4. Control joysticks
    5. Combine buttons and joysticks
    """
    import sys

    # Open serial port (adjust port name for your system)
    # Windows: 'COM3', 'COM4', etc.
    # Linux: '/dev/ttyACM0', '/dev/ttyUSB0', etc.
    # macOS: '/dev/cu.usbmodem*'

    port_name = 'COM3'  # Change this to match your system

    try:
        # Open serial connection at 115200 baud
        ser = serial.Serial(port_name, BAUD_RATE, timeout=1)
        print(f"Opened serial port: {port_name}")

        # Create controller instance
        controller = PABotBaseController(ser)

        # Connect and verify
        print("Connecting to PABotBase...")
        if not controller.connect():
            print("ERROR: Failed to connect!")
            return

        print("Connected successfully!")

        # Query device information
        protocol_ver = controller.get_protocol_version()
        program_ver = controller.get_program_version()
        program_id = controller.get_program_id()
        program_name = controller.get_program_name()
        queue_size = controller.get_queue_size()

        print(f"Protocol Version: {protocol_ver}")
        print(f"Program Version: {program_ver}")
        print(f"Program ID: 0x{program_id:02X}" if program_id else "Program ID: Unknown")
        print(f"Program Name: {program_name}")
        print(f"Queue Size: {queue_size}")

        print("\n--- Testing Controller ---")

        # Example 1: Press A button
        print("Pressing A button...")
        controller.press_button(Button.A, duration_ms=100)
        time.sleep(0.5)

        # Example 2: Press A and B together
        print("Pressing A + B buttons...")
        controller.press_button(Button.A | Button.B, duration_ms=150)
        time.sleep(0.5)

        # Example 3: Move left joystick right
        print("Moving left joystick right...")
        controller.move_joystick(left_x=0xFF, left_y=0x80, duration_ms=500)
        time.sleep(0.5)

        # Example 4: Move left joystick in a circle
        print("Moving joystick in circle...")
        for angle in range(0, 360, 45):
            import math
            x = int(128 + 127 * math.cos(math.radians(angle)))
            y = int(128 + 127 * math.sin(math.radians(angle)))
            controller.move_joystick(left_x=x, left_y=y, duration_ms=100)

        # Example 5: Press A while moving joystick
        print("Pressing A while moving joystick right...")
        controller.send_controller_state(
            buttons=Button.A,
            left_x=0xFF,
            left_y=0x80,
            duration_ms=200
        )
        time.sleep(0.5)

        # Return to neutral
        print("Returning to neutral...")
        controller.reset_to_neutral()

        print("\nTest complete!")

    except serial.SerialException as e:
        print(f"Serial port error: {e}")
        print("Make sure the Arduino is connected and the port name is correct.")
    except KeyboardInterrupt:
        print("\nInterrupted by user")
    finally:
        if 'ser' in locals() and ser.is_open:
            ser.close()
            print("Serial port closed")


if __name__ == '__main__':
    # Run example if executed directly
    print("PABotBase Controller - Python Implementation")
    print("=" * 60)
    example_usage()
