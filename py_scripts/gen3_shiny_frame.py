from __future__ import annotations

A = 0x41C64E6D
C = 0x00006073
MASK32 = 0xFFFFFFFF


def lcg_step(state: int) -> int:
    """One Gen 3 LCRNG step (mod 2^32)."""
    return (state * A + C) & MASK32


def lcg_advance(state: int, n: int) -> int:
    """
    Advance the Gen 3 LCRNG by n steps using fast exponentiation.
    Returns the state after n steps.
    """
    if n < 0:
        raise ValueError("n must be >= 0")

    mul = A
    add = C
    acc_mul = 1
    acc_add = 0

    while n:
        if n & 1:
            acc_mul = (acc_mul * mul) & MASK32
            acc_add = (acc_add * mul + add) & MASK32

        add = (add * ((mul + 1) & MASK32)) & MASK32
        mul = (mul * mul) & MASK32
        n >>= 1

    return (acc_mul * state + acc_add) & MASK32


def get_high16(state: int) -> int:
    """Get the high 16 bits of the state."""
    return (state >> 16) & 0xFFFF


def is_shiny(tid: int, sid: int, pid: int) -> bool:
    """Check if a PID is shiny for given TID/SID."""
    pid_high = (pid >> 16) & 0xFFFF
    pid_low = pid & 0xFFFF
    return (tid ^ sid ^ pid_high ^ pid_low) < 8


def find_shiny_frame(
    tid: int,
    sid: int,
    seed: int,
    min_advances: int,
    max_search: int = 1000000,
) -> dict:
    """
    Find the nearest frame that generates a shiny Pokemon in Gen 3.
    Args:
        tid: Trainer ID (0-65535)
        sid: Secret ID (0-65535)
        seed: Initial RNG seed (32-bit)
        min_advances: Minimum number of advances before checking
        max_search: Maximum frames to search (default 1 million)

    Returns:
        Dictionary with frame number and PID info, or error message
    """


    tid = int(tid) & 0xFFFF
    sid = int(sid) & 0xFFFF
    seed = int(seed) & MASK32
    min_advances = int(min_advances)

    # Advance seed to min_advances
    state = lcg_advance(seed, min_advances)

    for frame in range(min_advances, min_advances + max_search):
        # Generate PID from current frame
        # Call 1: PID low
        state1 = lcg_step(state)
        pid_low = get_high16(state1)

        # Call 2: PID high
        state2 = lcg_step(state1)
        pid_high = get_high16(state2)

        pid = (pid_high << 16) | pid_low

        if is_shiny(tid, sid, pid):
            return {
                "frame": frame,
                "pid": pid,
                "pid_hex": f"0x{pid:08X}",
                "advances_from_min": frame - min_advances,
            }

        # Advance to next frame
        state = lcg_step(state)

    return {
        "error": "No shiny frame found within search limit",
        "searched": max_search,
        "last_frame_checked": min_advances + max_search - 1,
    }


def main(tid=0, sid=0, seed=0, min_advances=0):
    """
    Find the nearest shiny frame for Gen 3 Pokemon.

    Args:
        tid: Trainer ID (0-65535)
        sid: Secret ID (0-65535)
        seed: Initial RNG seed (32-bit hex string or int)
        min_advances: Minimum advances before searching

    Returns:
        Dictionary with shiny frame info or error
    """
    # Handle hex string for seed
    if isinstance(seed, str):
        seed = seed.strip()
        if seed.lower().startswith("0x"):
            seed = int(seed, 16)
        else:
            seed = int(seed)

    shiny_info = find_shiny_frame(
        tid=int(tid),
        sid=int(sid),
        seed=seed,
        min_advances=int(min_advances),
    )
    return shiny_info