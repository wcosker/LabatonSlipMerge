# Standard Python
import os
# Third party
from pypdf import PdfWriter,PdfReader
import tkinter as tk
from tkinter import messagebox
from tkinter import ttk  # Import ttk for Progressbar widget
import pandas as pd

def run_merges(config, file_list, preserve_metadata, output_text):
    # Clear the text box before starting
    output_text.delete('1.0', tk.END)

    num_lines = len(file_list)
    output_files = []
    
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

    metadata = {}
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
                        pdf_path = os.path.join("Working Directory/"+input_path, file_name)

                        if os.path.exists(pdf_path):
                            merger.append(pdf_path)
                            if preserve_metadata:
                                reader = PdfReader(pdf_path)
                                if reader.metadata:
                                    metadata.update(reader.metadata)
                                    merger.add_metadata(metadata)
                        else:
                            pdf_path = pdf_path.replace("\\", "/")
                            pdf_path = pdf_path.replace("Working Directory/","")
                            messagebox.showerror("Error", f"No valid PDF path found for '{pdf_path}'")
                            progress_window.destroy()  # Close the progress window if there's an error
                            return

                output_path = "Working Directory/"+config_row[1]
                if os.path.exists(output_path):
                    merged_pdf_path = os.path.join(output_path, file_name)
                    merger.write(merged_pdf_path)
                    merger.close()

                    file_size = os.path.getsize(merged_pdf_path)
                    output_files.append((
                        merged_pdf_path.replace("\\", "/").replace("Working Directory/", ""),
                        file_size
                    ))
                else:
                    output_path = output_path.replace("\\", "/")
                    output_path = output_path.replace("Working Directory/","")
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

    # Sort the output files by size in descending order
    output_files.sort(key=lambda x: x[1], reverse=True)

    # Display success message in the output_text box
    output_text.insert(tk.END, "Success! \n\nMerged PDF files in descending order of size:\n", 'success')
    for file_path, size in output_files:
        output_text.insert(tk.END, f"    - {file_path}: {size / (1024 * 1024):.2f} MB\n", 'success')
