"""
Audio module for Controller Macro Runner.
Handles audio device enumeration and passthrough.
"""

# Audio support (optional)
try:
    import pyaudio
    PYAUDIO_AVAILABLE = True
except ImportError:
    PYAUDIO_AVAILABLE = False
    pyaudio = None


def list_audio_devices():
    """
    Returns (input_devices, output_devices) as lists of (index, name) tuples.
    """
    if not PYAUDIO_AVAILABLE or pyaudio is None:
        return [], []

    try:
        p = pyaudio.PyAudio()
        inputs = []
        outputs = []

        for i in range(p.get_device_count()):
            try:
                info = p.get_device_info_by_index(i)
                name = info.get('name', f'Device {i}')
                max_in = info.get('maxInputChannels', 0)
                max_out = info.get('maxOutputChannels', 0)

                if max_in > 0:
                    inputs.append((i, name))
                if max_out > 0:
                    outputs.append((i, name))
            except Exception:
                continue

        p.terminate()
        return inputs, outputs
    except Exception:
        return [], []
