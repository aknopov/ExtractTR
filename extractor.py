from openpyxl import Workbook
from openpyxl import load_workbook
import re

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
    print(f"Extracting data from {source} to {destination}...")  # UC
    wb_in = load_workbook(source, data_only=True)
    wb_out = load_workbook(destination)
    assert OUTPUT_SHEET == wb_out.sheetnames[wb_out._active_sheet_index]

    # print_mappings()

    # printCell(wb_in, INPUT_SHEET, "D6")
    # printCell(wb_in, INPUT_SHEET, "D7")
    # printCell(wb_in, INPUT_SHEET, "I7")

    # wb_out.save()


# def printCell(workbook, sheet, cellId):
#     cell = workbook[sheet][cellId]
#     print(f"Cell value: {cell.value}")


def print_mappings():
    for mapping in MAPPINGS:
        entry = mapping["sheet"] + " (" + mapping["in1"]
        if "in2" in mapping:
            entry += "+" + mapping["in2"]
        entry += ") -> " + mapping["out"]
        print(entry)


def get_last_row(workbook):
    return 1  # UC


def copy_one_value(mapping):
    if "in2" in mapping:
        print(f"Mapping has second value")


def is_number(s):
    return False if re.match(r'^[+-]?\d(>?\.\d+)?$', s) is None else True

def is_range(s):
    return False if re.match(r'^[+-]?\d(>?\.\d+)? - [+-]?\d(>?\.\d+)?$', s) is None else True

def mid_range(s):
    # split range string
    # return average
    return 0.0