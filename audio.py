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
    - Only showing WASAPI devices (Windows Audio Session API)
    - Filtering out legacy MME/DirectSound drivers
    - Removing duplicate device names
    - Excluding disabled, system, and virtual devices
    """
    if not PYAUDIO_AVAILABLE or pyaudio is None:
        return [], []

    try:
        p = pyaudio.PyAudio()

        # Find WASAPI host API index (required for Windows Sound System matching)
        host_api_count = p.get_host_api_count()
        wasapi_index = None

        for i in range(host_api_count):
            try:
                host_info = p.get_host_api_info_by_index(i)
                if 'WASAPI' in host_info.get('name', '').upper():
                    wasapi_index = i
                    break
            except Exception:
                continue

        # If WASAPI not found, return empty lists (Windows should have WASAPI)
        if wasapi_index is None:
            p.terminate()
            return [], []

        inputs = []
        outputs = []
        seen_input_names = set()
        seen_output_names = set()

        # Collect only WASAPI devices
        for i in range(p.get_device_count()):
            try:
                info = p.get_device_info_by_index(i)
                name = info.get('name', f'Device {i}')
                max_in = info.get('maxInputChannels', 0)
                max_out = info.get('maxOutputChannels', 0)
                host_api = info.get('hostApi', -1)

                # CRITICAL: Only accept WASAPI devices
                if host_api != wasapi_index:
                    continue

                # Skip devices with no channels
                if max_in == 0 and max_out == 0:
                    continue

                # Filter out legacy/system devices by name patterns
                skip_patterns = [
                    'Microsoft Sound Mapper',
                    'Primary Sound',  # Primary Sound Capture Driver, etc.
                    '@System32\\drivers\\',  # System driver devices
                    'Stereo Mix',  # Loopback devices
                ]
                if any(pattern in name for pattern in skip_patterns):
                    continue

                # Filter out disabled/unavailable devices
                if any(indicator in name.lower() for indicator in ['(unplugged)', '(disabled)', '(not present)']):
                    continue

                # Filter out numbered duplicate arrays (e.g., "Microphone Array 1", "Microphone Array 2")
                # These are internal channels that Windows Sound System doesn't show
                import re
                if re.search(r'\s+\d+\s*\(\s*\)$', name):  # Ends with " 1 ()" or similar
                    continue

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

            except Exception:
                continue

        p.terminate()
        return inputs, outputs
    except Exception:
        return [], []
