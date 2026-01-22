"""
Gen 3 Frame Finder - Determine which RNG frame was hit based on Pokemon stats

Given a Pokemon's actual stats, nature, and the RNG parameters, this script
finds the most likely frame that was hit within a search window.

Arguments:
    tid: Trainer ID (0-65535) - can be int or string
    sid: Secret ID (0-65535) - can be int or string
    seed: Initial RNG seed (hex string like "0x5A0" or int)
    target_frame: The frame you were aiming for - can be int or string
    frame_window: How many frames +/- to check - can be int or string
    method: 1 or 4 (Pokemon generation method) - can be int or string
    pokemon: Pokemon name (e.g., "Mudkip") - fuzzy matched
    level: Pokemon's level - can be int or string
    nature: Pokemon's nature (e.g., "Adamant") - fuzzy matched, or "" to ignore
    stats: List of [HP, Atk, Def, SpA, SpD, Spe]
           - Each stat can be int, string, or "" / None to skip that stat

Returns:
    dict with best_frame, pokemon_data, and candidates list
"""

# Gen 3 LCRNG constants
MULT = 0x41C64E6D
ADD = 0x6073
MASK = 0xFFFFFFFF

# Nature modifiers: [+stat_index, -stat_index] (1=Atk, 2=Def, 3=SpA, 4=SpD, 5=Spe)
# None means neutral
NATURES = {
    "Hardy": None, "Lonely": (1, 2), "Brave": (1, 5), "Adamant": (1, 3), "Naughty": (1, 4),
    "Bold": (2, 1), "Docile": None, "Relaxed": (2, 5), "Impish": (2, 3), "Lax": (2, 4),
    "Timid": (5, 1), "Hasty": (5, 2), "Serious": None, "Jolly": (5, 3), "Naive": (5, 4),
    "Modest": (3, 1), "Mild": (3, 2), "Quiet": (3, 5), "Bashful": None, "Rash": (3, 4),
    "Calm": (4, 1), "Gentle": (4, 2), "Sassy": (4, 5), "Careful": (4, 3), "Quirky": None
}

NATURE_LIST = [
    "Hardy", "Lonely", "Brave", "Adamant", "Naughty",
    "Bold", "Docile", "Relaxed", "Impish", "Lax",
    "Timid", "Hasty", "Serious", "Jolly", "Naive",
    "Modest", "Mild", "Quiet", "Bashful", "Rash",
    "Calm", "Gentle", "Sassy", "Careful", "Quirky"
]

# Base stats: [HP, Atk, Def, SpA, SpD, Spe]
BASE_STATS = {
    # Starters
    "Bulbasaur": [45, 49, 49, 65, 65, 45],
    "Ivysaur": [60, 62, 63, 80, 80, 60],
    "Venusaur": [80, 82, 83, 100, 100, 80],
    "Charmander": [39, 52, 43, 60, 50, 65],
    "Charmeleon": [58, 64, 58, 80, 65, 80],
    "Charizard": [78, 84, 78, 109, 85, 100],
    "Squirtle": [44, 48, 65, 50, 64, 43],
    "Wartortle": [59, 63, 80, 65, 80, 58],
    "Blastoise": [79, 83, 100, 85, 105, 78],
    "Chikorita": [45, 49, 65, 49, 65, 45],
    "Bayleef": [60, 62, 80, 63, 80, 60],
    "Meganium": [80, 82, 100, 83, 100, 80],
    "Cyndaquil": [39, 52, 43, 60, 50, 65],
    "Quilava": [58, 64, 58, 80, 65, 80],
    "Typhlosion": [78, 84, 78, 109, 85, 100],
    "Totodile": [50, 65, 64, 44, 48, 43],
    "Croconaw": [65, 80, 80, 59, 63, 58],
    "Feraligatr": [85, 105, 100, 79, 83, 78],
    "Treecko": [40, 45, 35, 65, 55, 70],
    "Grovyle": [50, 65, 45, 85, 65, 95],
    "Sceptile": [70, 85, 65, 105, 85, 120],
    "Torchic": [45, 60, 40, 70, 50, 45],
    "Combusken": [60, 85, 60, 85, 60, 55],
    "Blaziken": [80, 120, 70, 110, 70, 80],
    "Mudkip": [50, 70, 50, 50, 50, 40],
    "Marshtomp": [70, 85, 70, 60, 70, 50],
    "Swampert": [100, 110, 90, 85, 90, 60],
    # Legendaries
    "Rayquaza": [105, 150, 90, 150, 90, 95],
    "Kyogre": [100, 100, 90, 150, 140, 90],
    "Groudon": [100, 150, 140, 100, 90, 90],
    "Latias": [80, 80, 90, 110, 130, 110],
    "Latios": [80, 90, 80, 130, 110, 110],
    "Regirock": [80, 100, 200, 50, 100, 50],
    "Regice": [80, 50, 100, 100, 200, 50],
    "Registeel": [80, 75, 150, 75, 150, 50],
    "Deoxys": [50, 150, 50, 150, 50, 150],
    "Jirachi": [100, 100, 100, 100, 100, 100],
    "Mew": [100, 100, 100, 100, 100, 100],
    "Mewtwo": [106, 110, 90, 154, 90, 130],
    "Lugia": [106, 90, 130, 90, 154, 110],
    "Ho-Oh": [106, 130, 90, 110, 154, 90],
    "Celebi": [100, 100, 100, 100, 100, 100],
    # Common Gen 3
    "Ralts": [28, 25, 25, 45, 35, 40],
    "Kirlia": [38, 35, 35, 65, 55, 50],
    "Gardevoir": [68, 65, 65, 125, 115, 80],
    "Abra": [25, 20, 15, 105, 55, 90],
    "Kadabra": [40, 35, 30, 120, 70, 105],
    "Alakazam": [55, 50, 45, 135, 95, 120],
    "Beldum": [40, 55, 80, 35, 60, 30],
    "Metang": [60, 75, 100, 55, 80, 50],
    "Metagross": [80, 135, 130, 95, 90, 70],
    "Bagon": [45, 75, 60, 40, 30, 50],
    "Shelgon": [65, 95, 100, 60, 50, 50],
    "Salamence": [95, 135, 80, 110, 80, 100],
    "Larvitar": [50, 64, 50, 45, 50, 41],
    "Pupitar": [70, 84, 70, 65, 70, 51],
    "Tyranitar": [100, 134, 110, 95, 100, 61],
    "Dratini": [41, 64, 45, 50, 50, 50],
    "Dragonair": [61, 84, 65, 70, 70, 70],
    "Dragonite": [91, 134, 95, 100, 100, 80],
    "Eevee": [55, 55, 50, 45, 65, 55],
    "Magikarp": [20, 10, 55, 15, 20, 80],
    "Gyarados": [95, 125, 79, 60, 100, 81],
    "Feebas": [20, 15, 20, 10, 55, 80],
    "Milotic": [95, 60, 79, 100, 125, 81],
    "Wynaut": [95, 23, 48, 23, 48, 23],
    "Wobbuffet": [190, 33, 58, 33, 58, 33],
    "Pichu": [20, 40, 15, 35, 35, 60],
    "Pikachu": [35, 55, 40, 50, 50, 90],
    "Raichu": [60, 90, 55, 90, 80, 110],
    # Gift Pokemon
    "Castform": [70, 70, 70, 70, 70, 70],
    "Lileep": [66, 41, 77, 61, 87, 23],
    "Anorith": [45, 95, 50, 40, 50, 75],
    "Zigzagoon": [38, 30, 41, 30, 41, 60],
    "Wurmple": [45, 45, 35, 20, 30, 20],
    "Poochyena": [35, 55, 35, 30, 30, 35],
    "Lotad": [40, 30, 30, 40, 50, 30],
    "Seedot": [40, 40, 50, 30, 30, 30],
    "Taillow": [40, 55, 30, 30, 30, 85],
    "Wingull": [40, 30, 30, 55, 30, 85],
    "Surskit": [40, 30, 32, 50, 52, 65],
    "Shroomish": [60, 40, 60, 40, 60, 35],
    "Slakoth": [60, 60, 60, 35, 35, 30],
    "Nincada": [31, 45, 90, 30, 30, 40],
    "Whismur": [64, 51, 23, 51, 23, 28],
    "Makuhita": [72, 60, 30, 20, 30, 25],
    "Azurill": [50, 20, 40, 20, 40, 20],
    "Skitty": [50, 45, 45, 35, 35, 50],
    "Aron": [50, 70, 100, 40, 40, 30],
    "Electrike": [40, 45, 40, 65, 40, 65],
    "Plusle": [60, 50, 40, 85, 75, 95],
    "Minun": [60, 40, 50, 75, 85, 95],
    "Volbeat": [65, 73, 55, 47, 75, 85],
    "Illumise": [65, 47, 55, 73, 75, 85],
    "Roselia": [50, 60, 45, 100, 80, 65],
    "Gulpin": [70, 43, 53, 43, 53, 40],
    "Carvanha": [45, 90, 20, 65, 20, 65],
    "Wailmer": [130, 70, 35, 70, 35, 60],
    "Numel": [60, 60, 40, 65, 45, 35],
    "Torkoal": [70, 85, 140, 85, 70, 20],
    "Spoink": [60, 25, 35, 70, 80, 60],
    "Spinda": [60, 60, 60, 60, 60, 60],
    "Trapinch": [45, 100, 45, 45, 45, 10],
    "Cacnea": [50, 85, 40, 85, 40, 35],
    "Swablu": [45, 40, 60, 40, 75, 50],
    "Zangoose": [73, 115, 60, 60, 60, 90],
    "Seviper": [73, 100, 60, 100, 60, 65],
    "Lunatone": [70, 55, 65, 95, 85, 70],
    "Solrock": [70, 95, 85, 55, 65, 70],
    "Barboach": [50, 48, 43, 46, 41, 60],
    "Corphish": [43, 80, 65, 50, 35, 35],
    "Baltoy": [40, 40, 55, 40, 70, 55],
    "Shuppet": [44, 75, 35, 63, 33, 45],
    "Duskull": [20, 40, 90, 30, 90, 25],
    "Tropius": [99, 68, 83, 72, 87, 51],
    "Chimecho": [65, 50, 70, 95, 80, 65],
    "Absol": [65, 130, 60, 75, 60, 75],
    "Snorunt": [50, 50, 50, 50, 50, 50],
    "Spheal": [70, 40, 50, 55, 50, 25],
    "Clamperl": [35, 64, 85, 74, 55, 32],
    "Relicanth": [100, 90, 130, 45, 65, 55],
    "Luvdisc": [43, 30, 55, 40, 65, 97],
}


def parse_int(value, default=None):
    """
    Parse a value as an integer. Handles:
    - Integers (returned as-is)
    - Strings (parsed, including hex with 0x prefix)
    - Empty strings or None (returns default)
    """
    if value is None or value == "":
        return default
    if isinstance(value, int):
        return value
    if isinstance(value, str):
        value = value.strip()
        if value == "":
            return default
        if value.lower().startswith("0x"):
            return int(value, 16)
        return int(value)
    return int(value)


def normalize_stats(stats):
    """
    Normalize a stats list. Each stat can be:
    - An integer (kept as-is)
    - A string number (parsed to int)
    - Empty string or None (kept as None to indicate "skip")

    Returns list of 6 values (int or None)
    """
    result = []
    for i in range(6):
        if i < len(stats):
            val = stats[i]
            if val is None or val == "":
                result.append(None)
            elif isinstance(val, int):
                result.append(val)
            elif isinstance(val, str):
                val = val.strip()
                if val == "":
                    result.append(None)
                else:
                    result.append(int(val))
            else:
                result.append(int(val))
        else:
            result.append(None)
    return result


def fuzzy_match_nature(input_nature):
    """
    Fuzzy match a nature name. Handles:
    - Case insensitivity (ADAMANT, adamant, AdAmAnT -> Adamant)
    - Prefix matching (ada -> Adamant, tim -> Timid)
    - Partial matching with edit distance for typos

    Returns (matched_nature, confidence) or (None, 0) if no match
    """
    if not input_nature:
        return None, 0

    input_lower = input_nature.lower().strip()

    # Exact match (case insensitive)
    for nature in NATURE_LIST:
        if nature.lower() == input_lower:
            return nature, 1.0

    # Prefix match
    prefix_matches = [n for n in NATURE_LIST if n.lower().startswith(input_lower)]
    if len(prefix_matches) == 1:
        return prefix_matches[0], 0.9

    # Substring match (input is contained in nature name)
    substring_matches = [n for n in NATURE_LIST if input_lower in n.lower()]
    if len(substring_matches) == 1:
        return substring_matches[0], 0.8

    # Edit distance for typo tolerance
    def edit_distance(s1, s2):
        if len(s1) < len(s2):
            return edit_distance(s2, s1)
        if len(s2) == 0:
            return len(s1)

        prev_row = range(len(s2) + 1)
        for i, c1 in enumerate(s1):
            curr_row = [i + 1]
            for j, c2 in enumerate(s2):
                insertions = prev_row[j + 1] + 1
                deletions = curr_row[j] + 1
                substitutions = prev_row[j] + (c1 != c2)
                curr_row.append(min(insertions, deletions, substitutions))
            prev_row = curr_row
        return prev_row[-1]

    # Find best edit distance match
    best_match = None
    best_distance = float('inf')

    for nature in NATURE_LIST:
        dist = edit_distance(input_lower, nature.lower())
        if dist < best_distance:
            best_distance = dist
            best_match = nature

    # Accept if edit distance is small relative to string length
    max_allowed = max(2, len(input_lower) // 3)
    if best_distance <= max_allowed:
        confidence = 1.0 - (best_distance / len(best_match))
        return best_match, max(0.5, confidence)

    return None, 0


def lcrng_next(seed):
    """Advance LCRNG by one step"""
    return (seed * MULT + ADD) & MASK


def lcrng_advance(seed, n):
    """Advance LCRNG by n steps using fast exponentiation"""
    if n < 0:
        raise ValueError("n must be >= 0")

    mul = MULT
    add = ADD
    acc_mul = 1
    acc_add = 0

    while n:
        if n & 1:
            acc_mul = (acc_mul * mul) & MASK
            acc_add = (acc_add * mul + add) & MASK
        add = (add * ((mul + 1) & MASK)) & MASK
        mul = (mul * mul) & MASK
        n >>= 1

    return (acc_mul * seed + acc_add) & MASK


def get_pokemon_method1(seed):
    """
    Generate Pokemon using Method 1
    Returns (pid, ivs, nature, seed_after)
    """
    # PID generation
    seed = lcrng_next(seed)
    pid_low = (seed >> 16) & 0xFFFF
    seed = lcrng_next(seed)
    pid_high = (seed >> 16) & 0xFFFF
    pid = (pid_high << 16) | pid_low

    # IV generation
    seed = lcrng_next(seed)
    iv1 = (seed >> 16) & 0xFFFF
    seed = lcrng_next(seed)
    iv2 = (seed >> 16) & 0xFFFF

    # Extract IVs: HP, Atk, Def from iv1; Spe, SpA, SpD from iv2
    hp_iv = iv1 & 0x1F
    atk_iv = (iv1 >> 5) & 0x1F
    def_iv = (iv1 >> 10) & 0x1F
    spe_iv = iv2 & 0x1F
    spa_iv = (iv2 >> 5) & 0x1F
    spd_iv = (iv2 >> 10) & 0x1F

    ivs = [hp_iv, atk_iv, def_iv, spa_iv, spd_iv, spe_iv]
    nature = NATURE_LIST[pid % 25]

    return pid, ivs, nature, seed


def get_pokemon_method4(seed):
    """
    Generate Pokemon using Method 4
    Returns (pid, ivs, nature, seed_after)
    """
    # PID generation
    seed = lcrng_next(seed)
    pid_low = (seed >> 16) & 0xFFFF
    seed = lcrng_next(seed)
    pid_high = (seed >> 16) & 0xFFFF
    pid = (pid_high << 16) | pid_low

    # Method 4: Skip one RNG call before IVs
    seed = lcrng_next(seed)

    # IV generation
    seed = lcrng_next(seed)
    iv1 = (seed >> 16) & 0xFFFF
    seed = lcrng_next(seed)
    iv2 = (seed >> 16) & 0xFFFF

    # Extract IVs
    hp_iv = iv1 & 0x1F
    atk_iv = (iv1 >> 5) & 0x1F
    def_iv = (iv1 >> 10) & 0x1F
    spe_iv = iv2 & 0x1F
    spa_iv = (iv2 >> 5) & 0x1F
    spd_iv = (iv2 >> 10) & 0x1F

    ivs = [hp_iv, atk_iv, def_iv, spa_iv, spd_iv, spe_iv]
    nature = NATURE_LIST[pid % 25]

    return pid, ivs, nature, seed


def is_shiny(pid, tid, sid):
    """Check if PID is shiny for given TID/SID"""
    return ((pid >> 16) ^ (pid & 0xFFFF) ^ tid ^ sid) < 8


def calc_stat(base, iv, level, is_hp, nature_mod=1.0):
    """Calculate a Pokemon's stat at level with 0 EVs"""
    if is_hp:
        return ((2 * base + iv) * level // 100) + level + 10
    else:
        return int((((2 * base + iv) * level // 100) + 5) * nature_mod)


def get_nature_modifiers(nature_name):
    """Get stat modifiers for a nature. Returns list of 6 multipliers (HP always 1.0)"""
    mods = [1.0, 1.0, 1.0, 1.0, 1.0, 1.0]  # HP, Atk, Def, SpA, SpD, Spe
    nature_effect = NATURES.get(nature_name)
    if nature_effect:
        plus_idx, minus_idx = nature_effect
        mods[plus_idx] = 1.1
        mods[minus_idx] = 0.9
    return mods


def calc_all_stats(base_stats, ivs, level, nature_name):
    """Calculate all 6 stats for a Pokemon"""
    mods = get_nature_modifiers(nature_name)
    stats = []
    for i in range(6):
        stat = calc_stat(base_stats[i], ivs[i], level, i == 0, mods[i])
        stats.append(stat)
    return stats


def calc_stat_range(base_stats, level):
    """
    Calculate the minimum and maximum possible stats for a Pokemon.
    Returns (min_stats, max_stats) - each is a list of 6 values.

    Min: 0 IV, negative nature (0.9x for non-HP)
    Max: 31 IV, positive nature (1.1x for non-HP)
    """
    min_stats = []
    max_stats = []

    for i in range(6):
        base = base_stats[i]
        if i == 0:  # HP - no nature modifier
            min_stat = calc_stat(base, 0, level, True, 1.0)
            max_stat = calc_stat(base, 31, level, True, 1.0)
        else:  # Other stats - can have nature modifier
            min_stat = calc_stat(base, 0, level, False, 0.9)
            max_stat = calc_stat(base, 31, level, False, 1.1)

        min_stats.append(min_stat)
        max_stats.append(max_stat)

    return min_stats, max_stats


def validate_stats_in_range(stats, min_stats, max_stats):
    """
    Validate that each stat is within the possible range.
    Stats outside the range are set to None (ignored).
    Returns (validated_stats, out_of_range_indices).
    """
    validated = []
    out_of_range = []

    for i, stat in enumerate(stats):
        if stat is None:
            validated.append(None)
        elif stat < min_stats[i] or stat > max_stats[i]:
            validated.append(None)
            out_of_range.append(i)
        else:
            validated.append(stat)

    return validated, out_of_range


def stats_match_score(predicted_stats, actual_stats):
    """
    Calculate how well predicted stats match actual stats.
    Skips any stat where actual is None (not provided).
    Returns (exact_matches, total_diff, stats_compared)
    """
    exact = 0
    diff = 0
    compared = 0

    for p, a in zip(predicted_stats, actual_stats):
        if a is None:
            continue  # Skip this stat
        compared += 1
        if p == a:
            exact += 1
        diff += abs(p - a)

    return exact, diff, compared


def main(tid, sid, seed, target_frame, frame_window, method, pokemon, level, nature, stats):
    """
    Find the most likely frame hit based on Pokemon stats.

    Args:
        tid: Trainer ID
        sid: Secret ID
        seed: Initial RNG seed (hex string like "0xABCD1234" or int)
        target_frame: Frame you aimed for
        frame_window: +/- frames to search
        method: 1 or 4
        pokemon: Pokemon name
        level: Pokemon's level
        nature: Actual nature of the Pokemon
        stats: Actual stats [HP, Atk, Def, SpA, SpD, Spe]

    Returns:
        dict with results
    """
    # Parse all numeric parameters (support strings)
    tid = parse_int(tid, 0)
    sid = parse_int(sid, 0)
    seed = parse_int(seed, 0) & MASK
    target_frame = parse_int(target_frame, 0)
    frame_window = parse_int(frame_window, 100)
    method = parse_int(method, 1)
    level = parse_int(level, 5)

    # Normalize stats list (handle strings and empty values)
    stats = normalize_stats(stats)

    # Get base stats
    pokemon_title = pokemon.title()
    if pokemon_title not in BASE_STATS:
        return {"error": f"Unknown Pokemon: {pokemon}. Add it to BASE_STATS."}

    base_stats = BASE_STATS[pokemon_title]

    # Validate stats are within possible range for this Pokemon/level
    min_stats, max_stats = calc_stat_range(base_stats, level)
    stats, out_of_range = validate_stats_in_range(stats, min_stats, max_stats)

    # Fuzzy match nature (if match fails, ignore nature filter)
    matched_nature, confidence = fuzzy_match_nature(nature)
    nature_title = matched_nature  # None if no match - will skip nature filtering

    # Select method
    get_pokemon = get_pokemon_method1 if method == 1 else get_pokemon_method4

    # Calculate frame range
    start_frame = max(0, target_frame - frame_window)
    end_frame = target_frame + frame_window

    # Advance seed to start frame
    current_seed = lcrng_advance(seed, start_frame)

    candidates = []

    for frame in range(start_frame, end_frame + 1):
        # Generate Pokemon at this frame
        pid, ivs, frame_nature, _ = get_pokemon(current_seed)

        # Check if nature matches (or skip filter if nature_title is None)
        if nature_title is None or frame_nature == nature_title:
            # Calculate predicted stats
            predicted_stats = calc_all_stats(base_stats, ivs, level, frame_nature)

            # Score the match
            exact_matches, total_diff, stats_compared = stats_match_score(predicted_stats, stats)

            # Check shininess
            shiny = is_shiny(pid, tid, sid)

            candidates.append({
                "frame": frame,
                "offset": frame - target_frame,
                "pid": f"0x{pid:08X}",
                "ivs": ivs,
                "nature": frame_nature,
                "predicted_stats": predicted_stats,
                "exact_matches": exact_matches,
                "stats_compared": stats_compared,
                "total_diff": total_diff,
                "shiny": shiny,
                "seed": f"0x{current_seed:08X}"
            })

        # Advance to next frame
        current_seed = lcrng_next(current_seed)

    # Sort by exact matches (desc), then total diff (asc), then distance from target (asc)
    candidates.sort(key=lambda x: (-x["exact_matches"], x["total_diff"], abs(x["offset"])))

    if not candidates:
        return {
            "error": "No candidates found with matching nature in frame window",
            "frames_searched": end_frame - start_frame + 1
        }

    best = candidates[0]

    return {
        "best_frame": best["frame"],
        "best_offset": best["offset"],
        "best_pid": best["pid"],
        "best_ivs": best["ivs"],
        "best_nature": best["nature"],
        "best_predicted_stats": best["predicted_stats"],
        "actual_stats": stats,
        "stats_out_of_range": out_of_range,
        "stat_ranges": {"min": min_stats, "max": max_stats},
        "exact_matches": best["exact_matches"],
        "stats_compared": best["stats_compared"],
        "total_diff": best["total_diff"],
        "is_shiny": best["shiny"],
        "seed_at_frame": best["seed"],
        "method": method,
        "pokemon": pokemon_title,
        "level": level,
        "nature_input": nature,
        "nature_matched": nature_title,
        "nature_confidence": confidence,
        "candidates_count": len(candidates),
        "top_candidates": candidates[:10]
    }


# For testing outside of macro runner
if __name__ == "__main__":
    # Example: Finding frame for a Mudkip
    # Demonstrates: string numbers, fuzzy nature, skipped stats (empty string)
    result = main(
        tid="10141",           # String number
        sid="57319",           # String number
        seed="0x0",            # Hex string
        target_frame="7586",   # String number
        frame_window="20",    # String number
        method="1",            # String number
        pokemon="Mudkip",
        level="5",             # String number
        nature="LAPIS",         # Fuzzy matched (Random Word)
        stats=["20", "12", "11", "", "10", "9"]  # HP, Atk, Def, SpA, SpD, Spe
    )

    import json
    print(json.dumps(result, indent=2))
