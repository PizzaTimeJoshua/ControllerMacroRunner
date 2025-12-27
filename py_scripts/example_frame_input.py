# in their py_scripts/*.py
import base64
from io import BytesIO
from PIL import Image

def decode_frame(payload):
    if not payload or payload.get("__frame__") != "png_base64":
        return None
    data = base64.b64decode(payload["data_b64"])
    img = Image.open(BytesIO(data)).convert("RGB")
    # optionally convert to numpy:
    # import numpy as np
    # arr = np.array(img)  # RGB
    return img

def main(frame_payload, x, y):
    img = decode_frame(frame_payload)
    # do stuff...
    return True
