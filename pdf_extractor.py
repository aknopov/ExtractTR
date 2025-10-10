import os
import logging as log
import math
from typing import Any
import fitz # PyMuPDF
import converter as cnv
import xcl_extractor as xcl

MAPPINGS = [
    # PDF page, key string (lower case), num values, dest col, offsets (negative - absolute)
    {"page": 0, "key": "client:", "num": 1, "out": "B", "offset": [0]},
    {"page": 0, "key": "soil classification:", "num": 1, "out": "F", "offset": [0]},
    {"page": 0, "key": "borehole/sample no.:", "num": 3, "out": "G", "offset": [0]},
    {"page": 0, "key": "sample depth (m):", "num": 3, "out": "H", "offset": [0]},
    {"page": 0, "key": "sample depth (ft):", "num": 3, "out": "H", "offset": [0]},
    {"page": 0, "key": "sample depth:", "num": 3, "out": "H", "offset": [0]},
    {"page": 0, "key": "liquid limit:", "num": 1, "out": "I", "offset": [0]},
    {"page": 0, "key": "plastic limit:", "num": 1, "out": "J", "offset": [0]},
    {"page": 0, "key": "plastic index:", "num": 1, "out": "K", "offset": [0]},
    {"page": 0, "key": "initial water content, (%)", "num": 3, "out": "L", "offset": [0, 1, 2]},
    {"page": 0, "key": "dry unit weight (kn/m3)", "num": 3, "out": "T", "offset": [0, 1, 2]},
    {"page": 0, "key": "void ratio", "num": 3, "out": "U", "offset": [0, 1, 2]},
    # {"page": 0, "key": "", "num": 1, "out": "", "offset": [0]},
]


def extract_file(source: str, destination: str):
    log.info("Extracting data from file `%s` to `%s` ...", source, destination)

    with fitz.open(source) as doc:
        wb_out = xcl.open_workbook(destination)
        _extract_one(doc, wb_out)


def extract_dir(source: str, destination: str):
    log.info("Extracting data from directory '%s' to '%s' ...", source, destination)

    pdf_files = _list_pdf_files(source)
    if len(pdf_files) == 0:
        log.warning("No PDF files found in '%s'", source)
        return

    wb_out = xcl.open_workbook(destination)
    start_row = xcl.max_row(wb_out) + 1
    for source in pdf_files:
        with fitz.open(source) as doc:
            _extract_one(doc, wb_out)

    end_row = xcl.max_row(wb_out)
    xcl.sort_rows(wb_out, start_row, end_row)

    xcl.save_workbook(wb_out, destination)
    log.info("Done directory extracting")

def _list_pdf_files(dir_path: str):
    pdf_files = []
    for fn in os.listdir(dir_path):
        full_path = os.path.join(dir_path, fn)
        if full_path.lower().endswith(".pdf") and os.path.isfile(full_path):
            pdf_files.append(full_path)

    return pdf_files

def _extract_one(doc: fitz.Document, wb_out: xcl.ExcelWorkbook):
    last_row = xcl.max_row(wb_out)
    log.info("Extracting data from '%s' from row %d ...", doc.name, last_row)

    for mapping in MAPPINGS:
        _copy_one_value(doc, wb_out, last_row, mapping)

    for col in xcl.MERGE_COLS:
        xcl.merge_cells(wb_out, col, last_row)


def _copy_one_value(doc: fitz.Document, wb_out: xcl.ExcelWorkbook, last_row: int, mapping: Any):
    page_num = mapping["page"]
    key = mapping["key"]
    num_vals =  mapping["num"]
    column = mapping["out"]
    offset = mapping["offset"]

    key_words = key.lower().split()
    page = doc[page_num]
    page_boxes = page.get_text("words", sort=True)

    # As per https://pymupdf.readthedocs.io/en/latest/textpage.html#TextPage.extractWORDS
    page_words = [row[4] for row in page_boxes]

    num_keys = len(key_words)
    num_words = len(page_words)

    for i in range(num_words - num_keys + 1):
        if [s.lower() for s in page_words[i : i + num_keys]] == key_words:
            values = page_words[i + num_keys : i + num_keys + num_vals]
            if num_vals > 1 and len(offset) == 1:
                xcl.insert_one_value(_merge_vals(values), wb_out, last_row + 1, column)
            else:
                for j, val in enumerate(values):
                    xcl.insert_one_value(val, wb_out, last_row + offset[j] + 1, column)

            break


# Major cases:
# 1. three values are strings with '/' separator
# 2. thee values are numbers with optional units separated with '-'
# 3. first value is number with optional unit
def _merge_vals(values: list) -> Any:
    if len(values) == 3:
        if values[1] == '/':
            return values[0] + '_' + values[2]
        else:
            val1 = cnv.convert_value(cnv.remove_units(values[0]))
            val2 = cnv.convert_value(cnv.remove_units(values[2]))
            if values[1] == '-' and not math.isnan(val1) and not math.isnan(val2):
                return (val1 + val2) / 2
            elif not math.isnan(val1):
                return val1

    return "".join(values)
