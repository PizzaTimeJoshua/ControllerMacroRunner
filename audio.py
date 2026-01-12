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
    Filters devices to match Windows Sound System by:
    - Preferring WASAPI devices (Windows Audio Session API)
    - Filtering out Microsoft Sound Mapper (legacy)
    - Removing duplicate device names
    - Excluding disabled or unavailable devices
    """
    if not PYAUDIO_AVAILABLE or pyaudio is None:
        return [], []

    try:
        p = pyaudio.PyAudio()

        # First pass: collect all devices with metadata
        all_devices = []
        host_api_count = p.get_host_api_count()
        wasapi_index = None

        # Find WASAPI host API (preferred on Windows)
        for i in range(host_api_count):
            try:
                host_info = p.get_host_api_info_by_index(i)
                if 'WASAPI' in host_info.get('name', '').upper():
                    wasapi_index = i
                    break
            except Exception:
                continue

        # Collect all devices
        for i in range(p.get_device_count()):
            try:
                info = p.get_device_info_by_index(i)
                name = info.get('name', f'Device {i}')
                max_in = info.get('maxInputChannels', 0)
                max_out = info.get('maxOutputChannels', 0)
                host_api = info.get('hostApi', -1)

                # Skip Microsoft Sound Mapper (legacy Windows device)
                if 'Microsoft Sound Mapper' in name:
                    continue

                # Skip disabled or unavailable devices
                if any(indicator in name.lower() for indicator in ['(unplugged)', '(disabled)', '(not present)']):
                    continue

                # Skip devices with no channels
                if max_in == 0 and max_out == 0:
                    continue

                all_devices.append({
                    'index': i,
                    'name': name,
                    'max_in': max_in,
                    'max_out': max_out,
                    'host_api': host_api,
                    'is_wasapi': host_api == wasapi_index
                })
            except Exception:
                continue

        # Second pass: filter to get unique devices, preferring WASAPI
        inputs = []
        outputs = []
        seen_input_names = set()
        seen_output_names = set()

        # Sort: WASAPI first, then others
        all_devices.sort(key=lambda d: (not d['is_wasapi'], d['index']))

        for device in all_devices:
            i = device['index']
            name = device['name']
            max_in = device['max_in']
            max_out = device['max_out']

            # Add input devices (avoid duplicates)
            if max_in > 0:
                clean_name = name.strip()
                if clean_name not in seen_input_names:
                    inputs.append((i, name))
                    seen_input_names.add(clean_name)

            # Add output devices (avoid duplicates)
            if max_out > 0:
                clean_name = name.strip()
                if clean_name not in seen_output_names:
                    outputs.append((i, name))
                    seen_output_names.add(clean_name)

        p.terminate()
        return inputs, outputs
    except Exception:
        return [], []
