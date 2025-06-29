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
    # sheet name, src cell, dest col
    {"sheet": INPUT_SHEET, "in1": "D6", "out": "B"},
    {"sheet": INPUT_SHEET, "in1": "D8", "out": "D"},
    {"sheet": INPUT_SHEET, "in1": "I10", "out": "F"},
    {"sheet": INPUT_SHEET, "in1": "D9", "out": "G"},
    {"sheet": INPUT_SHEET, "in1": "I6", "out": "H"},
    {"sheet": REPORT_SHEET, "in1": "F21", "out": "L"},
    {"sheet": REPORT_SHEET, "in1": "G21", "out": "L"},
    {"sheet": REPORT_SHEET, "in1": "H21", "out": "L"},
    {"sheet": REPORT_SHEET, "in1": "F20", "out": "T"},
    {"sheet": REPORT_SHEET, "in1": "G20", "out": "T"},
    {"sheet": REPORT_SHEET, "in1": "H20", "out": "T"},
    {"sheet": REPORT_SHEET, "in1": "Q171", "out": "AG"},
    {"sheet": REPORT_SHEET, "in1": "Q171", "out": "AG"},
    {"sheet": REPORT_SHEET, "in1": "Q171", "out": "AG"},
    {"sheet": REPORT_SHEET, "in1": "R171", "out": "AH"},
    {"sheet": REPORT_SHEET, "in1": "R171", "out": "AH"},
    {"sheet": REPORT_SHEET, "in1": "R171", "out": "AH"},
    # for values - take average
    {"sheet": REPORT_SHEET, "in1": "Q163", "in2": "Q194", "out": "AJ"},
    {"sheet": REPORT_SHEET, "in1": "Q163", "in2": "Q194", "out": "AJ"},
    {"sheet": REPORT_SHEET, "in1": "Q163", "in2": "Q194", "out": "AJ"},
    # ? {"sheet": INPUT_SHEET, "D11", "AN"}
    {"sheet": DILATION_SHEET, "in1": "V4", "out": "AQ"},
    {"sheet": DILATION_SHEET, "in1": "V30", "out": "AQ"},
    {"sheet": DILATION_SHEET, "in1": "V51", "out": "AQ"},
    {"sheet": DILATION_SHEET, "in1": "V11", "out": "AR"},
    {"sheet": DILATION_SHEET, "in1": "V37", "out": "AR"},
    {"sheet": DILATION_SHEET, "in1": "V58", "out": "AR"},
    {"sheet": REPORT_SHEET, "in1": "F45", "out": "AS"},
    {"sheet": REPORT_SHEET, "in1": "G45", "out": "AS"},
    {"sheet": REPORT_SHEET, "in1": "H45", "out": "AS"},
    {"sheet": REPORT_SHEET, "in1": "F39", "out": "BB"},
    {"sheet": REPORT_SHEET, "in1": "G39", "out": "BB"},
    {"sheet": REPORT_SHEET, "in1": "H39", "out": "BB"},
    {"sheet": EIEUR_SHEET, "in1": "Y2", "out": "BC"},
    {"sheet": EIEUR_SHEET, "in1": "Y24", "out": "BC"},
    {"sheet": EIEUR_SHEET, "in1": "Y49", "out": "BC"},
    {"sheet": EIEUR_SHEET, "in1": "Y2", "out": "BD"},
    {"sheet": EIEUR_SHEET, "in1": "Y19", "out": "BE"},
    {"sheet": EIEUR_SHEET, "in1": "Y42", "out": "BE"},
    {"sheet": EIEUR_SHEET, "in1": "Y67", "out": "BE"},
    {"sheet": EIEUR_SHEET, "in1": "Y9", "out": "BH"},
    {"sheet": EIEUR_SHEET, "in1": "Y31", "out": "BH"},
    {"sheet": EIEUR_SHEET, "in1": "Y57", "out": "BH"},
    {"sheet": RFKN_SHEET, "in1": "G56", "out": "BS"},
    {"sheet": RFKN_SHEET, "in1": "G57", "out": "BT"},
    {"sheet": FLAC_SHEET, "in1": "AD22", "out": "BW"},
    {"sheet": FLAC_SHEET, "in1": "AD23", "out": "BW"},
    {"sheet": FLAC_SHEET, "in1": "AD24", "out": "BW"},
    {"sheet": INPUT_SHEET, "in1": "D11", "out": "CG"},
]


def extract(source, destination):
    log.info(f"Extracting data from `{source}` to `{destination}` ...")
    wb_in = load_workbook(source, data_only=True)
    wb_out = load_workbook(destination)

    # print_mappings()

    assert OUTPUT_SHEET == wb_out.sheetnames[wb_out._active_sheet_index]
    last_row = wb_out.active.max_row
    log.info(f"'{destination}' has {last_row} rows")

    wb_out.save(new_file_name(destination))
    log.info("Done extracting")


# def printCell(workbook, sheet, cellId):
#     cell = workbook[sheet][cellId]
#     print(f"Cell value: {cell.value}")

def new_file_name(fname):
    base, ext = os.path.splitext(fname)
    return base + ".new" + ext


def print_mappings():
    for mapping in MAPPINGS:
        entry = mapping["sheet"] + " (" + mapping["in1"]
        if "in2" in mapping:
            entry += "+" + mapping["in2"]
        entry += ") -> " + mapping["out"]
        print(entry)


def copy_one_value(mapping):
    if "in2" in mapping:
        print(f"Mapping has second value")


# Deal with either single value or a range
# return NaN if conversion impossible
def get_value(s):
    if is_number(s):
        return float(s)
    elif is_range(s):
        return mid_range(s)
    else:
        log.error(f"Can't parse {s} to number or range")
        return math.nan


def is_number(s):
    return True if s=='nan' or re.match(r"^[+-]?\d+(>?\.\d+)?$", s) is not None else False


def is_range(s):
    return True if re.match(r'^[+-]?\d+(>?\.\d+)? - [+-]?\d+(>?\.\d+)?$', s) is not None else False


def mid_range(s):
    parts = s.split("-", 4)
    start = math.nan
    end = math.nan
    try:
        match len(parts):
            case 2:  # simple range
                start = float(parts[0].strip())
                end = float(parts[1].strip())
            case 3:
                if parts[0] == "":
                    start = -float(parts[1].strip())
                    end = float(parts[2].strip())
                else:
                    start = float(parts[0].strip())
                    end = -float(parts[2].strip())
            case 4:
                start = -float(parts[1].strip())
                end = -float(parts[3].strip())
            case _:
                log.error(f"Invalid range format: {s}")

        return (start + end) / 2.0
    except ValueError:
        log.error(f"Invalid range format: {s}")

    return math.nan
