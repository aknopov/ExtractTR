import logging as log
import math
import re
from typing import Any

N_A = "N/A"


def col_name_to_idx(s: str):
    if isinstance(s, str):
        s = s.strip().upper()
    if len(s) == 1:
        return ord(s) - ord('A') + 1
    if len(s) == 2:
        return (ord(s[0]) - ord('A') + 1) * 26 + (ord(s[1]) - ord('A') + 1)

    log.warning("Invalid column data: %s", s)
    return 0

def col_idx_to_name(i: int):
    if i < 27:
        return chr(ord('A') + i - 1)
    if i < 26**2 -  1:
        c1 = chr(ord('A') + (i - 1) // 26 - 1)
        c2 = chr(ord('A') + (i - 1) % 26)
        return c1 + c2

    return None

def convert_na(value: Any):
    # Convert None or NaN to 'N/A'
    if value is None or value == '#NUM!' or (isinstance(value, float) and math.isnan(value)):
        return N_A
    return value


def may_be_convert(v: Any) -> Any:
    if isinstance(v, str):
        v = v.strip()
        if is_number(v) or is_range(v):
            v = convert_value(v)
    return v


# Deal with either single value or a range
# return NaN if conversion impossible
def convert_value(s: str) -> float:
    if is_number(s):
        return float(s)
    if is_range(s):
        return convert_range(s)

    log.warning("Can't parse %s to number or range", s)
    return math.nan


def is_number(s: str) -> bool:
    return bool(s=='nan' or re.match(r'^[+-]?\d+(>?\.\d+)?$', s) is not None)


def is_range(s: str) -> bool:
    return bool(re.match(r'^[+-]?\d+(>?\.\d+)? *- *[+-]?\d+(>?\.\d+)?$', s) is not None)


def remove_units(s: str) -> str:
    if re.match(r'^[+-]?\d+(>?\.\d+)?\D*$', s) is not None:
        #  find the last non-digit and truncate by it
        for i in range(len(s) - 1, -1, -1):
            if s[i].isdigit():
                return s[0:i+1]
    return None


def convert_range(s: str) -> float:
    parts = s.split("-", 4)
    start = math.nan
    end = math.nan
    try:
        match len(parts):
            case 2:  # simple range
                start = float(parts[0].strip())
                end = float(parts[1].strip())
            case 3: # one value is negative
                if parts[0] == "":
                    start = -float(parts[1].strip())
                    end = float(parts[2].strip())
                else:
                    start = float(parts[0].strip())
                    end = -float(parts[2].strip())
            case 4: # two negative values
                start = -float(parts[1].strip())
                end = -float(parts[3].strip())
            case _:
                log.warning("Invalid range format: %s", s)

        return (start + end) / 2.0
    except ValueError:
        log.warning("Invalid range format: %s", s)

    return math.nan
