import math
import os
from openpyxl import load_workbook, Worksheet
import logging as log
import converter as cnv

INPUT_SHEET = "Input"
REPORT_SHEET = "Report"
DILATION_SHEET = "Dilation"
EIEUR_SHEET = "Ei-Eur"
RFKN_SHEET = "Rf-K-n"
FLAC_SHEET = "FLAC1"
OUTPUT_SHEET = "Summary by Type"
SORT_COLUMN=cnv.col_name_to_idx("CG")

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
    {"sheet": REPORT_SHEET, "in1": "Q170", "out": "AH", "offset": 0},
    {"sheet": REPORT_SHEET, "in1": "Q170", "out": "AH", "offset": 1},
    {"sheet": REPORT_SHEET, "in1": "Q170", "out": "AH", "offset": 2},
    {"sheet": REPORT_SHEET, "in1": "Q163", "in2": "Q194", "out": "AJ", "offset": 0},
    {"sheet": REPORT_SHEET, "in1": "Q163", "in2": "Q194", "out": "AJ", "offset": 1},
    {"sheet": REPORT_SHEET, "in1": "Q163", "in2": "Q194", "out": "AJ", "offset": 2},
    {"sheet": REPORT_SHEET, "in1": "O164", "in2": "O195", "out": "AI", "offset": 0},
    {"sheet": REPORT_SHEET, "in1": "O164", "in2": "O195", "out": "AI", "offset": 1},
    {"sheet": REPORT_SHEET, "in1": "O164", "in2": "O195", "out": "AI", "offset": 2},
    {"sheet": INPUT_SHEET, "in1": "D11", "out": "AN", "offset": 0},
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
    "B", "D", "F", "G", "H", "AN", "BS", "BT", "BU", "BV"
]

def extract_file(source, destination):
    log.info(f"Extracting data from `{source}` to `{destination}` ...")
    wb_out = load_workbook(destination)
    assert OUTPUT_SHEET == wb_out.sheetnames[wb_out._active_sheet_index]

    extract_one(source, destination, wb_out)

    os.rename(destination, new_file_name(destination))
    wb_out.save(destination)
    log.info("Done file extracting")


def extract_dir(source, destination):
    log.info(f"Extracting data from `{source}` to `{destination}` ...")
    wb_out = load_workbook(destination)
    assert OUTPUT_SHEET == wb_out.sheetnames[wb_out._active_sheet_index]

    # Get list of *.xlsx files
    excel_files = list_excel_files(source)
    if len(excel_files) == 0:
        log.warninig(f"No excel files found in ${source}")
        return
    
    for fn in excel_files:
        extract_one(fn, destination, wb_out)

    os.rename(destination, new_file_name(destination))
    wb_out.save(destination)
    log.info("Done directory extracting")


def list_excel_files(dir_path):
    excel_files = []
    for fn in os.listdir(dir_path):
        full_path = os.path.join(dir_path, fn)
        if fn.endswith('.xlsx') and os.path.isfile(full_path):
            excel_files.append(full_path)

    return excel_files


def extract_one(source, destination, wb_out):
    wb_in = load_workbook(source, data_only=True)

    last_row = wb_out.active.max_row
    log.info(f"Extracting data from '{source}' to '{destination}' from row {last_row} ...")

    for mapping in MAPPINGS:
        copy_one_value(wb_in, wb_out, last_row, mapping)

    for col in MERGE_COLS:
        merge_cells(wb_out, col, last_row)


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
    v = cnv.may_be_convert(wb_in[mapping["sheet"]][mapping["in1"]].value)

    if "in2" in mapping:
        v2 = cnv.may_be_convert(wb_in[mapping["sheet"]][mapping["in2"]].value)
        if not math.isnan(v) and not math.isnan(v2):
            v = (v + v2) / 2.0

    v_conv = cnv.convert_na(v)
    if v_conv == cnv.N_A:
        log.warning(f"Couldn't extract value from sheet '{mapping["sheet"]}', cell '{mapping["in1"]}'")

    wb_out.active.cell(
        row = last_row + mapping["offset"] + 1,
        column = cnv.col_name_to_idx(mapping["out"]),
        value=v_conv
    )


def merge_cells(wb_out, col, last_row):
    colIdx = cnv.col_name_to_idx(col)
    wb_out.active.merge_cells(start_row=last_row + 1, end_row=last_row + 3, start_column=colIdx, end_column=colIdx)


def cell_sorter(c):
    return c[SORT_COLUMN]
    # return t[1] + " " + t[0][::-1]


# From https://stackoverflow.com/questions/44767554/sorting-with-openpyxl
# Copying format: https://stackoverflow.com/questions/45433425/how-to-move-cell-range-in-openpyxl-with-its-properties-hyperlinks-formatting-e
def sort_rows(ws: Worksheet, row_start: int, row_end: int, col_start: int, col_end: int, sorter=None):
    """ Sorts given rows of the sheet
        row_start   First row to be sorted
        row_end     Last row to be sorted (default last row)
        col_start   Start column to be considered in sort
        col_end     End column to be considered in sort
        sorter      Function that accepts a tuple of values and returns a sortable key
    """

    bottom = ws.max_row
    right = cnv.col_idx_to_name(ws.max_column)

    cell_rows = {}
    cols = range(col_start, col_end + 1)
    for row in range(row_start, row_end + 1):
        key = []
        for col in cols:
            key.append(ws.cell(row, col).value)
        cell_rows[key] = cell_rows.get(key, set()).union({row})

    order = sorted(cell_rows, key=sorter)

    ws.move_range(f"A{row_start}:{right}{row_end}", bottom)

    dest = row_start
    for src_key in order:
        for row in cell_rows[src_key]:
            src = row + bottom
            dist = dest - src
            ws.move_range(f"A{src}:{right}{src}", dist)
            dest += 1