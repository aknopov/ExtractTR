import os
import logging as log
import math
from typing import Any
import fitz # PyMuPDF
import converter as cnv
import xcl_extractor as xcl

MAPPINGS = [
    # PDF page, key string (lower case), num values, dest col[, offsets (if num > 1))
    # Keys are sorted in order of PDF words
    {"page": 0, "key": "client:", "till": "borehole/sample no.:", "num": 1, "out": "B"},
    {"page": 0, "key": "borehole/sample no.:", "till": "sample type:",  "num": 1, "out": "G"},
    {"page": 0, "key": "sample depth (m):", "till": "soil classification:", "num": 1, "out": "H"},
    {"page": 0, "key": "sample depth:", "till": "soil classification:", "num": 1, "out": "H"},#?
    {"page": 0, "key": "soil classification:", "till": "liquid limit:", "num": 1, "out": "F"},
    {"page": 0, "key": "liquid limit:", "num": 1, "out": "I"},
    {"page": 0, "key": "plastic limit:", "num": 1, "out": "J"},
    {"page": 0, "key": "plastic index:", "num": 1, "out": "K"},#?
    {"page": 0, "key": "initial water content, (%)", "num": 3, "out": "L"},
    {"page": 0, "key": "void ratio", "num": 3, "out": "U"},
    {"page": 0, "key": "dry unit weight (kn/m3)", "num": 3, "out": "T"},
]


def extract_file(source: str, wb_out: xcl.ExcelWorkbook) -> bool:
    log.info("Extracting data from file `%s` to `%s` ...", source, wb_out.filename)

    with fitz.open(source) as doc:
        _extract_one(doc, wb_out)

    log.info("Done PDF extracting")
    return True


def extract_dir(source: str, wb_out: xcl.ExcelWorkbook) -> bool:
    log.info("Extracting data from PDF files in directory '%s' to '%s' ...", source, wb_out.filename)

    pdf_files = _list_pdf_files(source)
    if len(pdf_files) == 0:
        log.warning("No PDF files found in '%s'", source)
        return False

    for source in pdf_files:
        with fitz.open(source) as doc:
            _extract_one(doc, wb_out)

    log.info("Done PDF extracting")
    return True

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

    page = doc[page_num]
    page_boxes = page.get_text("words", sort=True)
    # As per https://pymupdf.readthedocs.io/en/latest/textpage.html#TextPage.extractWORDS
    page_words = [row[4] for row in page_boxes]

    key_words = key.split()
    start_idx = _key_index(key_words, page_words, 0) #UC start_idx can be passed down
    if start_idx == -1:
        return
    start_idx += len(key_words)

    end_idx = start_idx + num_vals
    if "till" in mapping:
        till_words = mapping["till"].split()
        idx = _key_index(till_words, page_words, start_idx + 1)
        # idx = _key_index(till_words, page_words[start_idx:])
        if idx != -1:
            end_idx = idx

    values = page_words[start_idx : end_idx]

    if num_vals == 1 and len(values) > 1:
        xcl.insert_one_value(_merge_vals(values), wb_out, last_row + 1, column)
    else:
        for j, val in enumerate(values):
            xcl.insert_one_value(cnv.remove_units(val), wb_out, last_row + j + 1, column)


def _key_index(key_words: list, all_words: list, start_idx: int) -> int:
    num_keys = len(key_words)
    num_words = len(all_words)
    for i in range(start_idx, num_words - num_keys + 1):
        if [s.lower() for s in all_words[i : i + num_keys]] == key_words:
            return i

    log.warning("Words '%s' not found in the document", ",".join(key_words))
    return -1

# Major cases:
# 1. three values are strings with '/' separator
# 2. thee values are numbers with optional units separated with '-'
# 3. first value is number with optional unit
def _merge_vals(values: list) -> Any:
    if len(values) == 3:
        if values[1] == '/':
            return values[0] + '_' + values[2]
        else:
            val1 = cnv.remove_units(values[0])
            val2 = cnv.remove_units(values[2])
            if values[1] == '-' and not math.isnan(val1) and not math.isnan(val2):
                return (val1 + val2) / 2
            elif not math.isnan(val1):
                return val1

    return " ".join(values)
