import os
import sys
import tkinter as tk
from tkinter import filedialog, messagebox
import logging as log
from logging.handlers import RotatingFileHandler
from pathlib import Path
import extractor as ex


BUTTON_WIDTH = 15
LABEL_WIDTH = 70


class ExtractTRApp:
    """
    A class responsible for initializing GUI components and getting  source and destination paths.
    """
    def __init__(self):
        self.source_dir = ""
        self.source_file = ""
        self.destination_file = ""

        self.root = None
        self.source_f_label = None
        self.source_d_label = None
        self.destination_label = None
        self.extract_button = None

    def start_app_window(self):
        self.configure_logging()

        self.root = tk.Tk()
        self.root.title("Soil Test Data Extractor")
        self.root.geometry("500x260")
        self.root.resizable(False, False)
        self.root.iconbitmap(self.get_resource_path('res/extract.ico'))
        self.root.protocol("WM_DELETE_WINDOW", self.on_closing)

        frame = tk.Frame(self.root)
        source_f_button = tk.Button(
            frame,
            text="Source File",
            width=BUTTON_WIDTH,
            command=self.select_source_file
        )
        source_d_button = tk.Button(
            frame,
            text="Source Directory",
            width=BUTTON_WIDTH,
            command=self.select_source_dir
        )
        destination_button = tk.Button(
            frame,
            text="Destination File",
            width=BUTTON_WIDTH,
            command=self.select_destination_file,
        )
        self.extract_button = tk.Button(
            self.root,
            text="Extract",
            state="disabled",
            command=self.extract,
        )
        self.source_f_label = tk.Label(
            frame, borderwidth=1, width=LABEL_WIDTH, relief="groove", anchor="w"
        )
        self.source_d_label = tk.Label(
            frame, borderwidth=1, width=LABEL_WIDTH, relief="groove", anchor="w"
        )
        self.destination_label = tk.Label(
            frame, borderwidth=1, width=LABEL_WIDTH, relief="groove", anchor="w"
        )

        self.root.columnconfigure(0, weight=1)
        frame.grid(row=0, column=0, padx=10, pady=10, sticky="w")

        frame.columnconfigure(0, weight=1)
        source_d_button.grid(row=0, column=0, sticky="w")
        self.source_d_label.grid(row=1, column=0, sticky="w", pady=7)
        source_f_button.grid(row=2, column=0, sticky="w", pady=10)
        self.source_f_label.grid(row=3, column=0, sticky="w")
        destination_button.grid(row=4, column=0, sticky="w", pady=10)
        self.destination_label.grid(row=5, column=0, sticky="w")

        self.extract_button.grid(row=1, column=0, padx=30, pady=15)

        self.root.mainloop()


    def get_resource_path(self, rel_path):
        base_path = ''
        try:
            # PyInstaller stores data files in a tmp folder referred to as _MEIPASS
            base_path = sys._MEIPASS
        except AttributeError:
            pass

        path = os.path.join(base_path, rel_path)

        # If the path if exists
        return path if os.path.exists(path) else None


    def on_closing(self):
        self.root.destroy()

    def select_source_file(self):
        self.source_file = self.open_file_dialog()
        self.source_f_label.configure(text=self.source_file)
        self.source_f_label.update()
        self.enable_extract()

    def select_source_dir(self):
        self.source_dir = self.open_dir_dialog()
        self.source_d_label.configure(text=self.source_dir)
        self.source_d_label.update()
        self.enable_extract()

    def select_destination_file(self):
        self.destination_file = self.open_file_dialog()
        self.destination_label.configure(text=self.destination_file)
        self.destination_label.update()
        self.enable_extract()

    def extract(self):
        if self.source_file != "" and self.source_dir != "":
            messagebox.showerror("Error",
                "Either source directory or file should be specified.\n" \
                "Unselect one with Cancel button in a dialog.")
            return

        self.pre_extract_ui()

        if self.source_file != "":
            ex.extract_file(self.source_file, self.destination_file)
        else:
            ex.extract_dir(self.source_dir, self.destination_file)

        self.post_extract_ui()

    def open_dir_dialog(self):
        lnx_path = filedialog.askdirectory(
            title="Select a Folder with Spreadsheets",
            initialdir=".",
        )
        return os.path.normpath(lnx_path) if lnx_path != "" else ""

    def open_file_dialog(self):
        lnx_path = filedialog.askopenfilename(
            title="Select a Spreadsheet",
            initialdir=".",
            filetypes=[
                ("Excel files", "*.xls*"),
                ("All files", "*.*"),
            ],
        )
        return os.path.normpath(lnx_path) if lnx_path != "" else ""

    def pre_extract_ui(self):
        self.root.config(cursor="watch")
        self.root.update()
        self.disable_extract()

    def post_extract_ui(self):
        self.source_file = ""
        self.source_dir = ""
        self.source_f_label.configure(text=self.source_file)
        self.source_f_label.update()
        self.source_d_label.configure(text=self.source_dir)
        self.source_d_label.update()
        self.root.config(cursor="arrow")
        self.root.update()

    def enable_extract(self):
        if (self.source_file != "" or self.source_dir != "") and self.destination_file != "":
            self.extract_button.config(state="normal")
            self.extract_button.update()

    def disable_extract(self):
        self.extract_button.config(state="disabled")
        self.extract_button.update()

    def configure_logging(self):
        log_name = Path.home().joinpath("extractr.log").absolute()
        rot_handler = RotatingFileHandler(log_name, maxBytes=5 * 1024 * 1024, backupCount=5)
        log.basicConfig(
            level=log.INFO,
            format="%(asctime)s - %(levelname)s - %(message)s",
            handlers=[rot_handler],
        )

app = ExtractTRApp()
app.start_app_window()
