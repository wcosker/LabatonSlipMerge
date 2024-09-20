# Standard Python
import os
# Third party
from pypdf import PdfWriter
import tkinter as tk
from tkinter import messagebox
from tkinter import ttk  # Import ttk for Progressbar widget
import pandas as pd

def run_merges(config, file_list, scrub_metadata, output_text):
    # Clear the text box before starting
    output_text.delete('1.0', tk.END)

    num_lines = len(file_list)
    largest_file_size = 0
    largest_file_path = ""
    
    # Create a separate pop-up window for the progress bar
    progress_window = tk.Toplevel()
    progress_window.title("Merging Progress")
    progress_window.geometry("300x100")
    
    # Create a progress bar widget
    progress_bar = ttk.Progressbar(progress_window, orient='horizontal', mode='determinate', length=280)
    progress_bar.pack(pady=20)
    progress_bar['maximum'] = num_lines  # Set the maximum value to the number of files
    progress_label = tk.Label(progress_window, text="Merging PDF files...")
    progress_label.pack(pady=10)

    # Update the GUI to show the new window
    progress_window.update()

    # Merging process
    for i, row in enumerate(file_list):
        file_name = str(row[0]) + ".pdf"
        treatment_type = row[1]

        has_valid_treatment = False
        for config_row in config:
            merger = PdfWriter()

            if config_row[0] == treatment_type:
                has_valid_treatment = True
                for input_path in config_row[2:]:
                    if not pd.isna(input_path):
                        pdf_path = os.path.join(input_path, file_name)

                        if os.path.exists(pdf_path):
                            merger.append(pdf_path)
                        else:
                            pdf_path = pdf_path.replace("\\", "/")
                            messagebox.showerror("Error", f"No valid PDF path found for '{pdf_path}'")
                            progress_window.destroy()  # Close the progress window if there's an error
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
                    output_path = output_path.replace("\\", "/")
                    messagebox.showerror("Error", f"No valid output path exists for '{output_path}'")
                    progress_window.destroy()  # Close the progress window if there's an error
                    return

        if not has_valid_treatment:
            messagebox.showerror("Error", f"No valid treatment type found in the Config for '{file_name}' with treatment type '{treatment_type}'.")
            progress_window.destroy()  # Close the progress window if there's an error
            return

        # Update the progress bar and the GUI
        progress_bar['value'] = i + 1
        progress_window.update_idletasks()

    # Destroy the progress window when the merging is complete
    progress_window.destroy()

    # Display success message in the output_text box
    output_text.insert(tk.END, f"Success! The largest PDF file written was '{largest_file_path}' with a size of {largest_file_size / (1024 * 1024):.2f} MB.\n", 'success')
