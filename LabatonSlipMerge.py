import os
import pandas as pd
import sys
from pypdf import PdfWriter
from tqdm import tqdm # loop counter

def read_excel_sheets(file_path):
    try:
        # Load the Excel file
        excel_file = pd.ExcelFile(file_path)
        
        # Read the sheets into DataFrames
        config_df = excel_file.parse('Config')
        filelist_df = excel_file.parse('FileList')
        
        return config_df, filelist_df
    except FileNotFoundError:
        input(f"No {file_path} file found. Press any button to close...")
        sys.exit(1)
    except Exception as e:
        input(f"Error reading the Excel file: {e}. Press any button to close...")
        sys.exit(1)

def validate_paths_and_explain(config_df, filelist_df):
    # Convert DataFrames to list of lists
    config = config_df.values.tolist()

    # Store errors and explanations
    errors = []
    explanations = {}

    # Iterate through each row in the Config tab to gather unique treatment configurations
    for config_row in config:
        if len(config_row) < 2:
            errors.append(f"Config row for treatment type '{config_row[0]}' is missing the output folder.")
            continue

        treatment_type = config_row[0]
        output_path = config_row[1]
        
        # Gather input paths for this configuration
        input_paths = [input_path for input_path in config_row[2:] if not pd.isna(input_path)]
        
        # Validate input and output paths
        for input_path in input_paths:
            if not os.path.exists(input_path):
                errors.append(f"Input path does not exist: {input_path}")

        if not os.path.exists(output_path):
            errors.append(f"Output path does not exist: {output_path}")

        # Store a general explanation for this treatment type's configuration
        if treatment_type not in explanations:
            explanations[treatment_type] = []
        
        explanations[treatment_type].append({
            "input_paths": input_paths,
            "output_path": output_path
        })

    # Display errors if any
    if errors:
        print("Validation failed with the following errors:")
        for error in errors:
            print(f"- {error}")
    else:
        print("Validation passed successfully!")

    # Provide a general explanation of the merging process
    print("\nPDF merging process summary by Treatment Type:")
    for treatment_type, configs in explanations.items():
        total_inputs = sum(len(config['input_paths']) for config in configs)
        
        print(f"\nTreatment Type: '{treatment_type}'")
        print(f"  Total # of PDFs for '{treatment_type}' treatment type: {total_inputs}")
        
        for idx, config in enumerate(configs, 1):
            print(f"  Merge configuration {idx}:")
            print(f"    Merging order (top to bottom):")
            
            # Display the merging order for this configuration
            for i, input_path in enumerate(config['input_paths'], 1):
                print(f"      {i}. {input_path}")
            
            print(f"    Output folder: {config['output_path']}\n")

config_df, filelist_df = read_excel_sheets('process_inputs.xlsx')
validate_paths_and_explain(config_df,filelist_df)
sys.exit(1)
    
# Remove headers from both DataFrames
config = config_df.values.tolist()
fileList = filelist_df.values.tolist()

num_lines = len(fileList)

largest_file_size = 0
largest_file_path = ""

for row in tqdm(fileList,total=num_lines):
    file_name = str(row[0])+".pdf"
    treatment_type = row[1]
    # for each config row in the config file
    for config_row in config:
        merger = PdfWriter()

        # if the config name is the current row's treatment type
        if config_row[0] == treatment_type:
            # go through the config row and grab each input
            for input_path in config_row[2:]:
                # if the config row is not blank
                if not pd.isna(input_path):
                    #input pdf needs to exist
                    if os.path.exists(input_path+"/"+file_name):
                        merger.append(input_path+"/"+file_name)
                    else:
                        input(f"\nError: No valid PDF path found for treatment type '{treatment_type}' and input '{input_path}/{file_name}'. \nPress any button to close...")
                        sys.exit(1)
            # output path needs to exist
            if (os.path.exists(config_row[1])):
                merger.write(config_row[1]+"/"+file_name)
                merger.close()
            else:
                input(f"\nError: No valid output path exists for '{config_row[1]}'. Press any button to close...")
            file_size = os.path.getsize(config_row[1]+"/"+file_name)
            # Check for the largest file size that was outputted
            if file_size > largest_file_size:
                largest_file_path = (config_row[1]+"/"+file_name)
                largest_file_size = file_size
print(f"Success! The largest PDF file written was '{largest_file_path}' with a size of {largest_file_size/(1024*1024):.2f} megabytes.")
# input("Press any button to close...")