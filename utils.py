# Standard Python
import os
import sys
import webbrowser
# Third party
from tkinter import messagebox
import requests
import pandas as pd

def check_version(current_version,api_link,repo_link):
    response = requests.get(api_link)
    if response.json()["name"] != current_version:
        messagebox.showerror("Version Error", "You are using an outdated version of this program.\nPlease update to the latest version by downloading the latest release.")
        webbrowser.open(repo_link)

def resource_path(relative_path):
    """ Get the absolute path to the resource, works for dev and PyInstaller """
    try:
        # PyInstaller creates a temp folder and stores the path in _MEIPASS
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")

    return os.path.join(base_path, relative_path)

def get_excel_file_path():
        if getattr(sys, 'frozen', False):
        # If the application is run as a bundle (PyInstaller executable)
            base_path = os.path.dirname(sys.executable)
        else:
        # If run in a normal Python environment (development)
            base_path = os.path.dirname(os.path.abspath(__file__))

        file_path = os.path.join(base_path, 'process_inputs.xlsx')

        if os.path.exists(file_path):
            return file_path
        else:
            messagebox.showerror("Error", "Excel file 'process_inputs.xlsx' not found in the program directory.")
            return

def read_excel_sheets(file_path):
    try:
        excel_file = pd.ExcelFile(file_path)
        config_df = excel_file.parse('Config')
        filelist_df = excel_file.parse('FileList')
        return config_df.values.tolist(), filelist_df.values.tolist()
    except FileNotFoundError:
        messagebox.showerror("Error", f"No {file_path} file found.")
        sys.exit(1)
    except Exception as e:
        messagebox.showerror("Error", f"Error reading the Excel file: {e}")
        sys.exit(1)