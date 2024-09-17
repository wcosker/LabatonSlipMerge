import os
import pandas as pd
import sys
from pypdf import PdfWriter
from tqdm import tqdm
import tkinter as tk
from tkinter import messagebox, scrolledtext, ttk
from PIL import Image, ImageTk

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

def validate_paths_and_explain(config, file_list, output_text):
    output_text.tag_config('error', foreground='#FF0000')  # Red for errors
    output_text.tag_config('success', foreground='#006400')  # Green for success
    output_text.tag_config('heading', foreground='#062984', font=('Helvetica', 12, 'bold'))

    output_text.delete('1.0', tk.END)
    errors = []
    explanations = {}

    for file_row in file_list:
        file_name = str(file_row[0]) + ".pdf"
        treatment_type = file_row[1]

        for config_row in config:
            if config_row[0] == treatment_type:
                output_path = config_row[1]
                input_paths = [input_path for input_path in config_row[2:] if not pd.isna(input_path)]

                for input_path in input_paths:
                    pdf_path = os.path.join(input_path, file_name)

                    if not os.path.exists(pdf_path):
                        pdf_path = pdf_path.replace("\\", "/")
                        errors.append(f"Error: No valid PDF found for '{pdf_path}' within treatment type '{treatment_type}'.")

    if errors:
        output_text.insert(tk.END, "Validation failed with the following errors:\n", 'heading')
        for error in errors:
            output_text.insert(tk.END, f"- {error}\n", 'error')
    else:
        output_text.insert(tk.END, "Validation passed successfully!\n", 'success')

    output_text.insert(tk.END, "\nPDF merging process summary by Treatment Type:\n", 'heading')
    for config_row in config:
        treatment_type = config_row[0]
        output_path = config_row[1]
        input_paths = [input_path for input_path in config_row[2:] if not pd.isna(input_path)]

        output_text.insert(tk.END, f"\nTreatment Type: '{treatment_type}'\n", 'heading')
        output_text.insert(tk.END, f"  Output path: {output_path}\n")
        output_text.insert(tk.END, "  Input paths being merged (in order):\n")
        for input_path in input_paths:
            output_text.insert(tk.END, f"    - {input_path}\n")

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


def main_gui():
    def get_excel_file_path():
        program_dir = os.path.dirname(os.path.abspath(__file__))
        file_path = os.path.join(program_dir, 'process_inputs.xlsx')
        if os.path.exists(file_path):
            return file_path
        else:
            messagebox.showerror("Error", "Excel file 'process_inputs.xlsx' not found in the program directory.")
            return None

    def on_run_merge():
        file_path = get_excel_file_path()
        if file_path:
            config, filelist = read_excel_sheets(file_path)
            scrub_metadata = scrub_var.get()

            # Run the merge process
            run_merges(config, filelist, scrub_metadata, output_text)

    def on_validate():
        file_path = get_excel_file_path()
        if file_path:
            config, filelist = read_excel_sheets(file_path)
            validate_paths_and_explain(config, filelist, output_text)

    # GUI setup
    root = tk.Tk()
    root.title("Labaton Slip Merge Program")
    root.geometry("800x600")

    # Define colors
    bg_color = "#bfd1fc"
    button_color = "#062984"
    text_color = "#002C54"
    header_color = "#dceeff"

    # Set background color
    root.configure(bg=bg_color)

    # Set window icon
    icon_path = "LKS Secondary Logo Blue.png"  # Replace with your actual icon file path
    icon_image = Image.open(icon_path)
    icon_size = (32, 32)  # Typical size for window icons
    icon_image.thumbnail(icon_size, Image.LANCZOS)
    icon_photo = ImageTk.PhotoImage(icon_image)
    root.iconphoto(False, icon_photo)

    # Frame for buttons and options
    frame = tk.Frame(root, padx=20, pady=20, bg=bg_color)
    frame.pack(fill=tk.X)

    # Buttons and checkboxes
    run_button = tk.Button(frame, text="Run Merge", command=on_run_merge, width=20, bg=button_color, fg="white", font=('Helvetica', 12, 'bold'))
    run_button.grid(row=0, column=0, padx=10, pady=10)

    scrub_var = tk.BooleanVar()
    scrub_checkbox = tk.Checkbutton(frame, text="Scrub PDF Metadata", variable=scrub_var, font=('Helvetica', 12), bg=bg_color, fg=text_color)
    scrub_checkbox.grid(row=1, column=0, padx=10, pady=10, sticky='w')  # Positioned below the Run Merge button

    validate_button = tk.Button(frame, text="Validate Config", command=on_validate, width=20, bg=button_color, fg="white", font=('Helvetica', 12, 'bold'))
    validate_button.grid(row=0, column=1, padx=10, pady=10)

    # Scrolled Text box for output display
    output_text = scrolledtext.ScrolledText(root, width=95, height=25, wrap=tk.WORD, font=('Helvetica', 12), bg=header_color, fg=text_color)
    output_text.pack(padx=20, pady=10)

    root.mainloop()

if __name__ == "__main__":
    main_gui()
