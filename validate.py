# Standard Python
import os
# Third party
import tkinter as tk
import pandas as pd

def validate_paths_and_explain(config, file_list, output_text):
    output_text.tag_config('error', foreground='#FF0000')  # Red for errors
    output_text.tag_config('success', foreground='#006400')  # Green for success
    output_text.tag_config('heading', foreground='#062984', font=('Helvetica', 12, 'bold'))

    output_text.delete('1.0', tk.END)
    errors = []

    for file_row in file_list:
        file_name = str(file_row[0]) + ".pdf"
        treatment_type = file_row[1]

        has_valid_treatment = False
        for config_row in config:        
            if config_row[0] == treatment_type:
                has_valid_treatment = True
                output_path = config_row[1]
                input_paths = [input_path for input_path in config_row[2:] if not pd.isna(input_path)]

                for input_path in input_paths:
                    pdf_path = os.path.join("Working Directory/"+input_path, file_name)

                    if not os.path.exists(pdf_path):
                        pdf_path = pdf_path.replace("\\", "/")
                        pdf_path = pdf_path.replace("Working Directory/","")
                        errors.append(f"Error: No valid PDF found for '{pdf_path}' within treatment type '{treatment_type}'.")
        
        if not has_valid_treatment:
            errors.append(f"Error: No valid treatment type found in the Config for '{file_name}' with treatment type '{treatment_type}'.")

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