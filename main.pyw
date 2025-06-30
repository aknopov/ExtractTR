import os
import tkinter as tk
from tkinter import filedialog, messagebox
import extractor as ex
import logging as log
from logging.handlers import RotatingFileHandler
from pathlib import Path

source_dir = ""
source_file = ""
destination_file = ""
label_width = 70


def start_app_window():
    global root
    global source_f_label
    global source_d_label
    global destination_label
    global extract_button

    configure_logging()

    root = tk.Tk()
    root.title("Soil Test Data Extractor")
    root.geometry("500x260")
    root.resizable(False, False)
    root.protocol("WM_DELETE_WINDOW", on_closing)

    frame = tk.Frame(root)
    source_f_button = tk.Button(
        frame,
        text="Source File",
        width=10,
        command=select_source_file
    )
    source_d_button = tk.Button(
        frame,
        text="Source Directory",
        width=15,
        command=select_source_dir
    )
    destination_button = tk.Button(
        frame,
        text="Destination File",
        width=15,
        command=select_destination_file,
    )
    extract_button = tk.Button(
        root,
        text="Extract",
        state="disabled",
        command=extract,
    )
    source_f_label = tk.Label(
        frame, borderwidth=1, width=label_width, relief="groove", anchor="w"
    )
    source_d_label = tk.Label(
        frame, borderwidth=1, width=label_width, relief="groove", anchor="w"
    )
    destination_label = tk.Label(
        frame, borderwidth=1, width=label_width, relief="groove", anchor="w"
    )

    root.columnconfigure(0, weight=1)
    frame.grid(row=0, column=0, padx=10, pady=10, sticky="w")

    frame.columnconfigure(0, weight=1)
    source_d_button.grid(row=0, column=0, sticky="w")
    source_d_label.grid(row=1, column=0, sticky="w", pady=7)
    source_f_button.grid(row=2, column=0, sticky="w", pady=10)
    source_f_label.grid(row=3, column=0, sticky="w")
    destination_button.grid(row=4, column=0, sticky="w", pady=10)
    destination_label.grid(row=5, column=0, sticky="w")

    extract_button.grid(row=1, column=0, padx=30, pady=15)

    root.mainloop()


def on_closing():
    # UC cleanup and do not close before completing
    root.destroy()  # Destroy the Tkinter window


def select_source_file():
    global source_file
    global source_f_label
    source_file = open_file_dialog()
    source_f_label.configure(text=source_file)
    source_f_label.update()
    enable_extract()


def select_source_dir():
    global source_dir
    global source_d_label
    source_dir = open_dir_dialog()
    source_d_label.configure(text=source_dir)
    source_d_label.update()
    enable_extract()


def select_destination_file():
    global destination_file
    global destination_label
    destination_file = open_file_dialog()
    destination_label.configure(text=destination_file)
    destination_label.update()
    enable_extract()


def extract():
    global source_file
    global source_dir

    if source_file != "" and source_dir != "":
        messagebox.showerror("Error",
            "Either source directory or file should be specified.\nUnselect one with Cancel button in a dialog.")
        return

    disable_extract()

    if source_file != "":
        ex.extract_file(source_file, destination_file)
    else:
        ex.extract_dir(source_dir, destination_file)

    source_file = ""
    source_dir = ""
    source_f_label.configure(text=source_file)
    source_f_label.update()
    source_d_label.configure(text=source_dir)
    source_d_label.update()

def open_dir_dialog():
    lnx_path = filedialog.askdirectory(
        title="Select a Folder with Spreadsheets",
        initialdir=".",
    )
    return os.path.normpath(lnx_path) if lnx_path != "" else ""

def open_file_dialog():
    lnx_path = filedialog.askopenfilename(
        title="Select a Spreadsheet",
        initialdir=".",
        filetypes=[
            ("Excel files", "*.xlsx"),
            ("All files", "*.*"),
        ],
    )
    return os.path.normpath(lnx_path) if lnx_path != "" else ""


def enable_extract():
    global source_file
    global source_dir
    global destination_file
    global extract_button
    if (source_file != "" or source_dir != "") and destination_file != "":
        extract_button.config(state="normal")
        extract_button.update()


def disable_extract():
    global extract_button
    extract_button.config(state="disabled")
    extract_button.update()


def configure_logging():
    log_name = Path.home().joinpath("extractr.log").absolute()
    rot_handler = RotatingFileHandler(log_name, maxBytes=5 * 1024 * 1024, backupCount=5)
    log.basicConfig(
        level=log.INFO,
        format="%(asctime)s - %(levelname)s - %(message)s",
        handlers=[rot_handler],
    )


start_app_window()
