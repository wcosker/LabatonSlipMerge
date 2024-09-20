import os

from pypdf import PdfWriter
from tqdm import tqdm
import tkinter as tk
from tkinter import messagebox
import pandas as pd

def run_merges(config, file_list, scrub_metadata, output_text):
    # Clear the text box before starting
    output_text.delete('1.0', tk.END)

    num_lines = len(file_list)
    largest_file_size = 0
    largest_file_path = ""
    
    # Merging process
    for row in tqdm(file_list,total=num_lines):
        file_name = str(row[0]) + ".pdf"
        treatment_type = row[1]

        for config_row in config:
            merger = PdfWriter()

            if config_row[0] == treatment_type:
                for input_path in config_row[2:]:
                    if not pd.isna(input_path):
                        pdf_path = os.path.join(input_path, file_name)

                        if os.path.exists(pdf_path):
                            merger.append(pdf_path)
                        else:
                            messagebox.showerror("Error", f"No valid PDF path found for '{pdf_path}'")
                            return

                output_path = config_row[1]
                if os.path.exists(output_path):
                    merged_pdf_path = os.path.join(output_path, file_name)
                    merger.write(merged_pdf_path)
                    merger.close()

                    file_size = os.path.getsize(merged_pdf_path)
                    if file_size > largest_file_size:
                        largest_file_path = merged_pdf_path.replace("\\","/")
                        largest_file_size = file_size
                else:
                    messagebox.showerror("Error", f"No valid output path exists for '{output_path}'")
                    return

    output_text.insert(tk.END, f"Success! The largest PDF file written was '{largest_file_path}' with a size of {largest_file_size / (1024 * 1024):.2f} MB.\n", 'success')
