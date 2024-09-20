# Labaton Slip Merge

## Installation
1. **Download the Executable**: Obtain the Labaton Slip Merge executable file from 
https://github.com/wcosker/LabatonSlipMerge/releases 
(Make sure to download the latest release).
2. **Create the Excel File**: You will need an Excel file named `process_inputs.xlsx`. Make sure this file is located in the same folder as the executable.

## Preparing the Excel File
The Excel file should contain two tabs:

### Tab 1: **FileList**
This tab should list the PDF files you want to merge and their exhibit numbers. Example:

| Exhibit Number  | Treatment Type |
|-----------------|----------------|
| 1               | Public         |
| 2               | Sealed         |
| 3               | Redacted       |

### Tab 2: **Config**
This tab defines how the files will be merged and where to save the final document. Example:

| Treatment Type | Output Path          | Input Path 1        | Input Path 2       | Input Path x... |
|----------------|----------------------|---------------------|--------------------|-----------------|
| Public         | output/public        | input/GenericSlips  | input/Exhibits     | (Path to input) |
| Sealed         | output/sealed        | input/Exhibits      |                    |                 |

## How to Use the Tool
1. **Run the Application**: Double-click the `Labaton_Slip_Merge.exe` file to launch the application.
2. **Validate Paths**: Click the "Validate Config" button to ensure that all file paths in your Excel sheet are correct and for an explanation of the config file for a further sanity check.
3. **Run the Merge**: Click the "Run Merge" button to start merging the PDF files.
4. **Check Output**: Your merged PDFs will be saved in the locations specified in the Config tab of your Excel file.

## Troubleshooting
If you encounter any errors, the application will display messages to help you identify what went wrong. Common issues include:
- Missing PDF files.
- Incorrect output paths.
- Problems reading the Excel file.

## Support
For help or questions, please reach out to Will Cosker at wcosker@labaton.com
