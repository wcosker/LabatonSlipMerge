import shutil
import csv
import sys
import traceback
from pypdf import PdfWriter
import glob # get csv and pdf list
from tqdm import tqdm # loop counter


# Append slip cover to PDF
def append_pdf(master_slip,row):
    merger.append(master_slip[int(row[1])-1])
    merger.append("Input/Raw Exhibits/" + row[0])

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

config = list(config_file)

for row in tqdm(fileList,total=num_lines):
    merger = PdfWriter()
    # for each config row in the config file
    for configs in config:
        # if the config name is the current row's treatment type
        if configs[0] == row[1]:
            # go through the config row and grab each input
            for x in range(len(configs)):
                # if the config row is not blank
                if configs[x] != '' and x != 0 and x != 1:
                    merger.append(configs[x]+"/"+row[0])
                if x == len(configs)-1:
                    merger.write(configs[1]+"/"+row[0])
                    merger.close()
    '''
    merger = PdfWriter()
    # If document is not subject to request to seal
    if int(row[2]) == 1:
        shutil.copyfile("Input/Raw Exhibits/"+row[0],"Output/Sealed Filing/" + row[0])
        shutil.copyfile("Input/Raw Exhibits/"+row[0],"Output/Public Version/" + row[0])
    # If document is fully redacted
    elif int(row[2]) == 2:
        shutil.copyfile(master_sealing[int(row[1])-1],"Output/Sealing Motion/" + row[0])
        shutil.copyfile(master_public[int(row[1])-1],"Output/Public Version/" + row[0])
        append_pdf(master_slip=master_sealed,row=row)
        merger.write("Output/Sealed Filing/" + row[0])
    # If document is partially redacted
    elif int(row[2]) == 3:
        if (int(row[3]) == 0):
            append_pdf(master_slip=master_sealed,row=row)
            merger.write("Output/Sealed Filing/" + row[0])
        elif (int(row[3]) == 1):
            # Merger sealing motion and public motion files
            append_pdf(master_slip=master_sealing,row=row)
            merger.write("Output/Sealing Motion/" + row[0])
            merger.close()
            merger = PdfWriter()
            append_pdf(master_slip=master_public,row=row)
            merger.write("Output/Public Version/" + row[0])
        else:
            print("\nSkipping row containing: " + row[0] + " file name. Incorrect Partial Redaction value.")
    else:
        print("\nSkipping row containing: " + row[0] + " file name. Incorrect ExhibitType")
    merger.close()
    '''

# input("\nSuccess! Press any button to close...")