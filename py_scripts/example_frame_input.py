"""
Example: Processing camera frames in custom Python scripts.

This demonstrates how to decode the $frame payload passed by run_python commands.
The frame is encoded as base64 PNG to allow subprocess communication without
shared memory.
"""
import base64
from io import BytesIO
from PIL import Image


def decode_frame(payload):
    """Decode a $frame payload into a PIL Image."""
    if not payload or payload.get("__frame__") != "png_base64":
        return None
    data = base64.b64decode(payload["data_b64"])
    img = Image.open(BytesIO(data)).convert("RGB")
    return img


def main(frame_payload, x, y):
    """Entry point called by run_python command with args: [$frame, x, y]."""
    img = decode_frame(frame_payload)
    # Process the image here...
    return True
