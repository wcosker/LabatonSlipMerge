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
        
        return config_df.values.tolist(), filelist_df.values.tolist()
    except FileNotFoundError:
        input(f"No {file_path} file found. Press Enter to close...")
        sys.exit(1)
    except Exception as e:
        input(f"Error reading the Excel file: {e}. Press Enter to close...")
        sys.exit(1)

def validate_paths_and_explain(config, file_list):
        # Store errors and explanations
        errors = []
        explanations = {}

        # Iterate through each row in the FileList
        for file_row in file_list:
            file_name = str(file_row[0]) + ".pdf"  # Get the file name with the .pdf extension
            treatment_type = file_row[1]  # Get the treatment type from the FileList

            # Iterate through the Config to find the matching treatment type
            for config_row in config:
                if config_row[0] == treatment_type:  # If treatment type matches
                    output_path = config_row[1]  # The output path from the Config row
                    input_paths = [input_path for input_path in config_row[2:] if not pd.isna(input_path)]  # Get the input paths

                    # Check if the file exists in any of the input paths
                    for input_path in input_paths:
                        pdf_path = os.path.join(input_path, file_name)

                        if not os.path.exists(pdf_path):
                            pdf_path = pdf_path.replace("\\","/")
                            errors.append(f"Error: No valid PDF found for filepath '{pdf_path}' within treatment type '{treatment_type}'.")

        # Display errors if any
        if errors:
            print("\nValidation failed with the following errors:")
            for error in errors:
                print(f"- {error}")
        else:
            print("\nValidation passed successfully!")

        input("\nPress Enter to continue to config explanation....")
        # Provide a general explanation of the merging process by treatment type
        print("\nPDF merging process summary by Treatment Type:")
        for config_row in config:
            treatment_type = config_row[0]
            output_path = config_row[1]
            input_paths = [input_path for input_path in config_row[2:] if not pd.isna(input_path)]

            print(f"\nTreatment Type: '{treatment_type}'")
            print(f"  Output path: {output_path}")
            print(f"  Input paths being merged (in order):")
            for input_path in input_paths:
                print(f"    - {input_path}")

def run_merges(config,fileList,isScrub):
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
                            input(f"\nError: No valid PDF path found for treatment type '{treatment_type}' and input '{input_path}/{file_name}'. \nPress Enter to close...")
                            sys.exit(1)
                # output path needs to exist
                if (os.path.exists(config_row[1])):
                    merger.write(config_row[1]+"/"+file_name)
                    merger.close()
                else:
                    input(f"\nError: No valid output path exists for '{config_row[1]}'. Press Enter to close...")
                file_size = os.path.getsize(config_row[1]+"/"+file_name)
                # Check for the largest file size that was outputted
                if file_size > largest_file_size:
                    largest_file_path = (config_row[1]+"/"+file_name)
                    largest_file_size = file_size
    print(f"Success! The largest PDF file written was '{largest_file_path}' with a size of {largest_file_size/(1024*1024):.2f} megabytes.")
    input("Press Enter to close...")

def main():
    config, filelist = read_excel_sheets('process_inputs.xlsx')
    # Welcome message and instructions (printed once)
    print("Welcome to the Labaton Slip Merge program.")
    print("Please select an option:")
    print("1. Run the merge process")
    print("2. Validate inputs and explain the config file")
    print("Press 'e' to exit the program.")

    # Prompt the user for their choice with input validation
    user_choice = None
    while user_choice not in ["1", "2", "e"]:
        user_choice = input("Enter your choice (1, 2, or 'e' to exit): ").lower()
        if user_choice not in ["1", "2", "e"]:
            print("Invalid choice. Please enter 1, 2, or 'e' to exit.")

    # Execute based on the user's choice
    if user_choice == "e":
        print("Exiting the program. Goodbye!")
        return

    elif user_choice == "1":
        # Ask if the user wants to scrub metadata
        scrub_metadata = None
        while scrub_metadata not in ["y", "n", "e"]:
            scrub_metadata = input("Would you like to scrub PDF metadata before merging? (y/n or 'e' to exit): ").lower()
            if scrub_metadata not in ["y", "n", "e"]:
                print("Invalid choice. Please enter 'y' for yes, 'n' for no, or 'e' to exit.")

        if scrub_metadata == "e":
            print("Exiting the program. Goodbye!")
            return
        elif scrub_metadata == "y":
            # Run the merge process
            run_merges(config, filelist,True)
        else:
            run_merges(config, filelist,False)

    elif user_choice == "2":
        # Validate inputs and explain the config file
        validate_paths_and_explain(config, filelist)


if __name__ == "__main__":
    main()