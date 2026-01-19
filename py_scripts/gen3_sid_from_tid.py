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
            # compose: acc(x) = acc_mul*x + acc_add; then apply T(x)=mul*x+add
            acc_mul = (acc_mul * mul) & MASK32
            acc_add = (acc_add * mul + add) & MASK32

        # square the transform: T(T(x)) = (mul^2)*x + add*(mul+1)
        add = (add * ((mul + 1) & MASK32)) & MASK32
        mul = (mul * mul) & MASK32
        n >>= 1

    return (acc_mul * state + acc_add) & MASK32


def sid_from_tid_and_advances(
    tid: int,
    advances: int,
    *,
    advances_are_zero_based: bool = True,
    extra_calls: int = 0,
) -> int:
    """
    Compute SID from TID and an RNG advance index.

    Matches common Gen 3 tooling conventions where:
      - Initial RNG state is 0x0000TID.
      - The RNG value used is the upper 16 bits of the *advanced* state.

    Indexing:
      - If advances_are_zero_based=True:
          advances=0 means "the first RNG call after seeding".
          steps = advances + 1
      - If False:
          advances=1 means "the first RNG call after seeding".
          steps = advances

    extra_calls:
      - Add this if the SID is taken from a later RNG call on the same moment/frame.
        (Example: if you know the advance for the *first* call but SID uses the *second*,
         set extra_calls=1.)

    Returns:
      SID as an int in [0, 65535].
    """
    if not (0 <= tid <= 0xFFFF):
        raise ValueError("TID must be a 16-bit value (0..65535).")
    if advances < 0 or extra_calls < 0:
        raise ValueError("advances and extra_calls must be >= 0.")

    seed0 = tid  # 0x0000TID (32-bit state with leading zeros)

    steps = advances + extra_calls + (1 if advances_are_zero_based else 0)
    state = lcg_advance(seed0, steps)
    sid = (state >> 16) & 0xFFFF
    return sid


def main(tid, advances):
    if tid == "" or tid.isnumeric() == False:
        return []
    seed = int(tid)
    possible_sids = []
    possible_sids.append(sid_from_tid_and_advances(seed, int(advances)))
    possible_sids.append(sid_from_tid_and_advances(seed, int(advances) +1))
    possible_sids.append(sid_from_tid_and_advances(seed, int(advances) -1))
    return possible_sids