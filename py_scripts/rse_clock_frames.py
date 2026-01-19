import time
import datetime
# Convert real time to  frames for RSE Wall Clock

def real_time_to_frames(timestamp) -> int:
    time_struct = datetime.datetime.fromtimestamp(timestamp).timetuple()
    hours = time_struct.tm_hour
    minutes = time_struct.tm_min
    frames = 0
    if hours == 10:
        if minutes <= 10:
            frames = 6 * minutes
        if minutes > 10:
            frames = 60 + (minutes - 10) * 3
        if minutes > 30:
            frames = 125 + (minutes - 30) * 2
    else:
        if hours < 10:
            hours += 24
        frames = 185 + (hours - 11) * 60 + minutes - 5
    return frames
def main():
    current_time = time.time()
    frames = real_time_to_frames(current_time)
    return frames