import tkinter as tk
from tkinter import filedialog

source_file = ""
destination_file = ""
widget_width = 20

# Google: "python create GUI with file open dialog"

def start_app_window():
    global root
    global source_label
    global destination_label
    root = tk.Tk()
    root.title("Soil Test Data Extractor")
    root.geometry("350x200")
    root.protocol("WM_DELETE_WINDOW", on_closing)

    source_button = tk.Button(
        root, text="Source File", width=widget_width, command=select_source_file
    )
    destination_button = tk.Button(
        root,
        text="Destination File",
        width=widget_width,
        command=select_destination_file,
    )
    # extract_button = tk.Button(
    #     root,
    #     text="Extract",
    #     command=extract,
    # )
    source_label = tk.Label(root, borderwidth=1, width=widget_width, relief="groove")
    destination_label = tk.Label(root, borderwidth=1, width=widget_width, relief="groove")

    source_button.grid(row=0, column=0, padx=10, pady=10)
    destination_button.grid(row=0, column=1, padx=10, pady=10)
    source_label.grid(row=1, column=0, padx=10, pady=10)
    destination_label.grid(row=1, column=1, padx=10, pady=10)
    # extract_button.pack(tk.CENTER, padx=10, pady=10)

    root.mainloop()


def on_closing():
    # UC do not close before completing
    root.destroy()  # Destroy the Tkinter window


def select_source_file():
    global source_file
    global source_label
    source_file = open_file_dialog()
    source_label.configure(text = source_file)
    source_label.update()


def select_destination_file():
    global destination_file
    global destination_label
    destination_file = open_file_dialog()
    destination_label.configure(text = destination_file)
    destination_label.update()

def extract():
    print("Starting extraction...")

def open_file_dialog():
    return filedialog.askopenfilename(
        title="Select a Spreadsheet",
        initialdir=".",  # Optional: Set initial directory
        filetypes=[
            ("Excel files", "*.xlsx"),
            ("All files", "*.*"),
        ],
    )


start_app_window()
