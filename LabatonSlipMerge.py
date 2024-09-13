import os
import csv
import sys
from pypdf import PdfWriter
from tqdm import tqdm # loop counter

try:
    config_file = csv.reader(open('config.csv',newline=''))
except:
    input("No config.csv file found. Press any button to close...")
    sys.exit(1)

try:
    fileList = csv.reader(open('filelist.csv',newline=''))
except:
    input("No filelist.csv file found. Press any button to close...")
    sys.exit(1)

num_lines = sum(1 for row in open('fileList.csv','r'))-1

# remove headers from both files
next(fileList)
next(config_file)

largest_file_size = 0
largest_file_path = ""

config = list(config_file)
for row in tqdm(fileList,total=num_lines):
    file_name = row[0]+".pdf"
    treatment_type = row[1]
    # for each config row in the config file
    for config_row in config:
        merger = PdfWriter()

        # if the config name is the current row's treatment type
        if config_row[0] == treatment_type:
            # go through the config row and grab each input
            for input_path in config_row[2:]:
                # if the config row is not blank
                if input_path:
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