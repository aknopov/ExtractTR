import math
import os
from openpyxl import Workbook
from openpyxl import load_workbook
import re
import logging as log

INPUT_SHEET = "Input"
REPORT_SHEET = "Report"
DILATION_SHEET = "Dilation"
EIEUR_SHEET = "Ei-Eur"
RFKN_SHEET = "Rf-K-n"
FLAC_SHEET = "FLAC1"
OUTPUT_SHEET = "Summary by Type"
N_A = "N/A"

MAPPINGS = [
    # sheet name, src cell 1, [src cell 2,] dest col
    {"sheet": INPUT_SHEET, "in1": "D6", "out": "B", "offset": 0},
    {"sheet": INPUT_SHEET, "in1": "D8", "out": "D", "offset": 0},
    {"sheet": INPUT_SHEET, "in1": "I10", "out": "F", "offset": 0},
    {"sheet": INPUT_SHEET, "in1": "D9", "out": "G", "offset": 0},
    {"sheet": INPUT_SHEET, "in1": "I6", "out": "H", "offset": 0},
    {"sheet": REPORT_SHEET, "in1": "F21", "out": "L", "offset": 0},
    {"sheet": REPORT_SHEET, "in1": "G21", "out": "L", "offset": 1},
    {"sheet": REPORT_SHEET, "in1": "H21", "out": "L", "offset": 2},
    {"sheet": REPORT_SHEET, "in1": "F20", "out": "T", "offset": 0},
    {"sheet": REPORT_SHEET, "in1": "G20", "out": "T", "offset": 1},
    {"sheet": REPORT_SHEET, "in1": "H20", "out": "T", "offset": 2},
    {"sheet": REPORT_SHEET, "in1": "Q171", "out": "AG", "offset": 0},
    {"sheet": REPORT_SHEET, "in1": "Q171", "out": "AG", "offset": 1},
    {"sheet": REPORT_SHEET, "in1": "Q171", "out": "AG", "offset": 2},
    {"sheet": REPORT_SHEET, "in1": "R171", "out": "AH", "offset": 0},
    {"sheet": REPORT_SHEET, "in1": "R171", "out": "AH", "offset": 1},
    {"sheet": REPORT_SHEET, "in1": "R171", "out": "AH", "offset": 2},
    {"sheet": REPORT_SHEET, "in1": "Q163", "in2": "Q194", "out": "AJ", "offset": 0},
    {"sheet": REPORT_SHEET, "in1": "Q163", "in2": "Q194", "out": "AJ", "offset": 1},
    {"sheet": REPORT_SHEET, "in1": "Q163", "in2": "Q194", "out": "AJ", "offset": 2},
    # ? {"sheet": INPUT_SHEET, "D11", "AN", "offset": 0}
    {"sheet": DILATION_SHEET, "in1": "V4", "out": "AQ", "offset": 0},
    {"sheet": DILATION_SHEET, "in1": "V30", "out": "AQ", "offset": 1},
    {"sheet": DILATION_SHEET, "in1": "V51", "out": "AQ", "offset": 2},
    {"sheet": DILATION_SHEET, "in1": "V11", "out": "AR", "offset": 0},
    {"sheet": DILATION_SHEET, "in1": "V37", "out": "AR", "offset": 1},
    {"sheet": DILATION_SHEET, "in1": "V58", "out": "AR", "offset": 2},
    {"sheet": REPORT_SHEET, "in1": "F45", "out": "AS", "offset": 0},
    {"sheet": REPORT_SHEET, "in1": "G45", "out": "AS", "offset": 1},
    {"sheet": REPORT_SHEET, "in1": "H45", "out": "AS", "offset": 2},
    {"sheet": REPORT_SHEET, "in1": "F39", "out": "BB", "offset": 0},
    {"sheet": REPORT_SHEET, "in1": "G39", "out": "BB", "offset": 1},
    {"sheet": REPORT_SHEET, "in1": "H39", "out": "BB", "offset": 2},
    {"sheet": EIEUR_SHEET, "in1": "Y2", "out": "BC", "offset": 0},
    {"sheet": EIEUR_SHEET, "in1": "Y24", "out": "BC", "offset": 1},
    {"sheet": EIEUR_SHEET, "in1": "Y49", "out": "BC", "offset": 2},
    {"sheet": EIEUR_SHEET, "in1": "Y2", "out": "BD", "offset": 0},
    {"sheet": EIEUR_SHEET, "in1": "Y19", "out": "BE", "offset": 0},
    {"sheet": EIEUR_SHEET, "in1": "Y42", "out": "BE", "offset": 1},
    {"sheet": EIEUR_SHEET, "in1": "Y67", "out": "BE", "offset": 2},
    {"sheet": EIEUR_SHEET, "in1": "Y9", "out": "BH", "offset": 0},
    {"sheet": EIEUR_SHEET, "in1": "Y31", "out": "BH", "offset": 1},
    {"sheet": EIEUR_SHEET, "in1": "Y57", "out": "BH", "offset": 2},
    {"sheet": RFKN_SHEET, "in1": "G56", "out": "BS", "offset": 0},
    {"sheet": RFKN_SHEET, "in1": "G57", "out": "BT", "offset": 0},
    {"sheet": RFKN_SHEET, "in1": "G75", "out": "BU", "offset": 0},
    {"sheet": RFKN_SHEET, "in1": "G76", "out": "BV", "offset": 0},
    {"sheet": FLAC_SHEET, "in1": "AD22", "out": "BW", "offset": 0},
    {"sheet": FLAC_SHEET, "in1": "AD23", "out": "BW", "offset": 1},
    {"sheet": FLAC_SHEET, "in1": "AD24", "out": "BW", "offset": 2},
    {"sheet": INPUT_SHEET, "in1": "D11", "out": "CG", "offset": 0},
    {"sheet": INPUT_SHEET, "in1": "D11", "out": "CG", "offset": 1},
    {"sheet": INPUT_SHEET, "in1": "D11", "out": "CG", "offset": 2},
]

MERGE_COLS = [
    "B", "D", "F", "G", "H", "AE", "AN", "AP", "BS", "BT", "BU", "BV", "BX"
]

def extract(source, destination):
    log.info(f"Extracting data from `{source}` to `{destination}` ...")
    wb_in = load_workbook(source, data_only=True)
    wb_out = load_workbook(destination)

    # print_mappings()

    assert OUTPUT_SHEET == wb_out.sheetnames[wb_out._active_sheet_index]
    last_row = wb_out.active.max_row
    log.info(f"Extracting data in '{destination}' to row {last_row}")

    for mapping in MAPPINGS:
        copy_one_value(wb_in, wb_out, last_row, mapping)

    for col in MERGE_COLS:
        merge_cells(wb_out, col, last_row)

    os.rename(destination, new_file_name(destination))
    wb_out.save(destination)
    log.info("Done extracting")


def new_file_name(fname):
    base, ext = os.path.splitext(fname)
    return base + ".org" + ext


# def print_mappings():
#     for mapping in MAPPINGS:
#         entry = mapping["sheet"] + " (" + mapping["in1"]
#         if "in2" in mapping:
#             entry += "+" + mapping["in2"]
#         entry += ") -> " + mapping["out"]
#         print(entry)


def copy_one_value(wb_in, wb_out, last_row, mapping):
    v = may_be_convert(wb_in[mapping["sheet"]][mapping["in1"]].value)

    if "in2" in mapping:
        v2 = may_be_convert(wb_in[mapping["sheet"]][mapping["in2"]].value)
        if not math.isnan(v) and not math.isnan(v2):
            v = (v + v2) / 2.0

    wb_out.active.cell(
        row=last_row + mapping["offset"] + 1,
        column=convert_column(mapping["out"]),
        value=convert_na(v)
    )


def merge_cells(wb_out, col, last_row):
    colIdx = convert_column(col)
    wb_out.active.merge_cells(start_row=last_row + 1, end_row=last_row + 3, start_column=colIdx, end_column=colIdx)


def convert_column(s):
    if isinstance(s, str):
        s = s.strip().upper()
    if len(s) == 1:
        return ord(s) - ord('A') + 1
    elif len(s) == 2:
        return (ord(s[0]) - ord('A') + 1) * 26 + (ord(s[1]) - ord('A') + 1)
    else:
        log.error(f"Invalid column format: {s}")
        return 0

def convert_na(value):
    # Convert None or NaN to 'N/A'
    if value is None or (isinstance(value, float) and math.isnan(value)):
        return N_A
    return value


def may_be_convert(v):
    if isinstance(v, str):
        v = v.strip()
        if is_number(v) or is_range(v):
            v = convert_value(v)
    return v


# Deal with either single value or a range
# return NaN if conversion impossible
def convert_value(s):
    if is_number(s):
        return float(s)
    elif is_range(s):
        return convert_range(s)
    else:
        log.error(f"Can't parse {s} to number or range")
        return math.nan


def is_number(s):
    return True if s=='nan' or re.match(r"^[+-]?\d+(>?\.\d+)?$", s) is not None else False


def is_range(s):
    return True if re.match(r'^[+-]?\d+(>?\.\d+)? - [+-]?\d+(>?\.\d+)?$', s) is not None else False


def convert_range(s):
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
                log.error(f"Invalid range format: {s}")

        return (start + end) / 2.0
    except ValueError:
        log.error(f"Invalid range format: {s}")

    return math.nan
