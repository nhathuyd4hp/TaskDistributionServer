# --- Main.py ---
import argparse
import logging
import os
import shutil
import threading
import time
import tkinter as tk
from tkinter import ttk

import pandas as pd
from config_access_token import token_file  # noqa

# Excel Check Engine
from excel_check import Excel_check
from Nasiwak import Bot_Update
from progress_reporter import add_result, clear_report, save_report

# Replace with your actual file path
file_path = os.path.join(os.getcwd(), "Access_token", "Access_token.txt")
logging.info(f"file path for text file is: {file_path}")
# Open and read the file
with open(file_path, "r", encoding="utf-8") as file:
    content = file.read()
logging.info(f"Extracted text from .txt file is: {content}")

# 🚀 Update Check
REPO_OWNER = "Nasiwak"
REPO_NAME = "Zenbu_bot"
CURRENT_VERSION = "v3.0"
ACCESS_TOKEN = content

Bot_Update(REPO_OWNER, REPO_NAME, CURRENT_VERSION, ACCESS_TOKEN)

# 📦 Global Variables
Ankens_bango_builder = {}
result = []

# 📋 GUI Functions


def clear_placeholder(event, entry, placeholder):
    if entry.get() == placeholder:
        entry.delete(0, tk.END)
        entry.config(foreground="black")


def restore_placeholder(event, entry, placeholder):
    if entry.get() == "":
        entry.insert(0, placeholder)
        entry.config(foreground="gray")


def load_values():

    if text_area.compare("end-1c", "==", "1.0"):
        text_area.insert(tk.END, f"{'Anken Bango':^30} {'Builder ID':^30}\n")
        text_area.insert(tk.END, "-" * 80 + "\n")


def read_excel_file(file_path):
    try:
        df = pd.read_excel(file_path, dtype=str)
        if "Anken Number" in df.columns and "Builder Code" in df.columns:
            anken_data = df["Anken Number"]
            builder_data = df["Builder Code"]

            if text_area.compare("end-1c", "==", "1.0"):
                text_area.insert(tk.END, f"{'Anken Bango':^30} {'Builder ID':^30}\n")
                text_area.insert(tk.END, "-" * 80 + "\n")

            for anken, builder in zip(anken_data, builder_data):
                Ankens_bango_builder[anken] = builder
                text_area.insert(tk.END, f"{anken:^30} {builder:^30}\n")
        else:
            logging.error("Excel must have 'Anken Number' and 'Builder Code' columns.")
            root.quit()
            root.destroy()
    except Exception as e:
        logging.error(f"Failed to read the Excel file:\n{str(e)}")
        root.quit()
        root.destroy()


def Start_Excel_check():
    global Ankens_bango_builder

    text_area.delete(1.0, tk.END)
    text_area.insert(tk.END, f"{'Anken Bango':^30} {'Builder ID':^30} {'Status':^20}\n")
    text_area.insert(tk.END, "-" * 80 + "\n")

    Ankens_bango_builder_sorted = dict(sorted(Ankens_bango_builder.items(), key=lambda item: item[1]))
    clear_report()  # 🧹 Clear previous session report

    zenbu = Excel_check()

    for anken, builder in Ankens_bango_builder_sorted.items():
        try:
            folder_path = os.path.join("Ankens", anken)
            if not os.path.exists(folder_path):
                zenbu.driver.maximize_window()
                zenbu.data_fetching(anken, builder)
            else:
                zenbu.Display(f"📂 {anken} already present. Skipping download.")
                continue
        except Exception as e:
            text_area.insert(tk.END, f"{anken:^30} {builder:^30} {'❌':^20}\n")
            add_result(anken, builder, "Failed")  # 📦 Report Failed
            print(f"Error: {e}")
        else:
            text_area.insert(tk.END, f"{anken:^30} {builder:^30} {'✅':^20}\n")
            add_result(anken, builder, "Success")  # 📦 Report Success

    zenbu.driver.quit()
    Ankens_bango_builder.clear()

    save_report()  # 💾 Save the full Progress_Report at the end


def empty_ankens_folder():
    Ankens_folder = "Ankens"
    if os.path.exists(Ankens_folder):
        shutil.rmtree(Ankens_folder)
        logging.info(f"Deleted contents of {Ankens_folder} folder.")
    os.makedirs(Ankens_folder)
    logging.info(f"Created an empty {Ankens_folder} folder.")
    time.sleep(1)


def threaded_excel_check():
    thread = threading.Thread(target=Start_Excel_check)
    thread.start()

    monitor_thread(thread)


def monitor_thread(thread):
    if thread.is_alive():
        root.after(5000, lambda: monitor_thread(thread))
    else:
        root.quit()
        root.destroy()


class TextHandler(logging.Handler):
    """Custom logging handler that writes to a Tkinter Text widget."""

    def __init__(self, text_widget):
        logging.Handler.__init__(self)
        self.text_widget = text_widget

    def emit(self, record):
        msg = self.format(record)

        def append():
            self.text_widget.configure(state="normal")
            self.text_widget.insert(tk.END, msg + "\n")
            self.text_widget.configure(state="disabled")
            self.text_widget.yview(tk.END)  # Auto-scroll to latest log

        self.text_widget.after(0, append)


def parse_args():
    parser = argparse.ArgumentParser()
    parser.add_argument(
        "--file",
        required=True,
    )
    return parser.parse_args()


if __name__ == "__main__":
    args = parse_args()

    root = tk.Tk()
    root.title("Zenbu Bot")
    root.geometry("1000x700")
    root.configure(bg="#d6e0f0")

    # 🌟 CLEAR OLD ANKENS FOLDER
    empty_ankens_folder()

    # GUI Components
    logo_frame = tk.Frame(root, bg="#d6e0f0")
    logo_frame.pack(fill="x", padx=10, pady=5)

    title_label = tk.Label(root, text="Zenbu Bot", font=("Segoe UI Emoji", 18, "bold"), bg="#d6e0f0")
    title_label.pack(pady=(10, 5))

    entry_frame = tk.Frame(root, bg="#d6e0f0")
    entry_frame.pack(pady=10)

    excel_frame = tk.Frame(root, bg="#d6e0f0")
    excel_frame.pack(pady=10)

    excel_file_entry = ttk.Entry(excel_frame, width=90)
    excel_file_entry.grid(row=0, column=1, padx=10)
    excel_file_entry.insert(0, args.file)
    excel_file_entry.state(["readonly"])

    text_area = tk.Text(root, height=15, width=80, bg="#e6eefc", font=("Segoe UI Emoji", 10))
    text_area.pack(pady=10)

    # After text_area creation
    text_handler = TextHandler(text_area)

    start_button = ttk.Button(root, text="START", width=10, command=threaded_excel_check)
    start_button.pack(pady=10)

    footer_label = tk.Label(
        root, text=f"Nasiwak Services India Pvt Ltd  V{CURRENT_VERSION}", bg="#d6e0f0", fg="#333333"
    )
    footer_label.pack(side="bottom", pady=(5, 0))

    root.after(5000, read_excel_file, args.file)

    root.after(10000, read_excel_file, start_button.invoke)

    root.mainloop()
