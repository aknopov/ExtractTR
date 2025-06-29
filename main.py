import tkinter as tk
from tkinter import filedialog
import extractor as ex
from extractor import extract

source_file = ""
destination_file = ""
label_width = 70

def start_app_window():
    global root
    global source_label
    global destination_label
    global extract_button

    root = tk.Tk()
    root.title("Soil Test Data Extractor")
    root.geometry("500x200")
    root.resizable(False, False)
    root.protocol("WM_DELETE_WINDOW", on_closing)

    frame = tk.Frame(root)
    source_button = tk.Button(
        frame,
        text="Source File", 
        width=10, 
        command=select_source_file
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
    source_label = tk.Label(
        frame, borderwidth=1, width=label_width, relief="groove", anchor="w"
    )
    destination_label = tk.Label(
        frame, borderwidth=1, width=label_width, relief="groove", anchor="w"
    )

    root.columnconfigure(0, weight=1) 
    frame.grid(row=0, column=0, padx=10, pady=10, sticky="w")
    
    frame.columnconfigure(0, weight=1) 
    source_button.grid(row=0, column=0, sticky="w")
    source_label.grid(row=1, column=0, sticky="w", pady=7)
    destination_button.grid(row=2, column=0, sticky="w", pady=10)
    destination_label.grid(row=3, column=0, sticky="w")

    extract_button.grid(row=1, column=0, padx=30, pady=15)

    root.mainloop()


def on_closing():
    # UC do not close before completing
    root.destroy()  # Destroy the Tkinter window


def select_source_file():
    global source_file
    global source_label
    source_file = open_file_dialog()
    source_label.configure(text=source_file)
    source_label.update()
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
    disable_extract()

    ex.extract(source_file, destination_file)

    source_file = ""
    source_label.configure(text=source_file)
    source_label.update()

def open_file_dialog():
    return filedialog.askopenfilename(
        title="Select a Spreadsheet",
        initialdir=".",  # Optional: Set initial directory
        filetypes=[
            ("Excel files", "*.xlsx"),
            ("All files", "*.*"),
        ],
    )


def enable_extract():
    global source_file
    global destination_file
    global extract_button
    if source_file != "" and destination_file != "":
        extract_button.config(state="normal")
        extract_button.update()

def disable_extract():
    global extract_button
    extract_button.config(state="disabled")
    extract_button.update()

start_app_window()
