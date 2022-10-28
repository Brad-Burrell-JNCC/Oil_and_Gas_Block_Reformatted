# Name: OilAndGasBlockReformatted.py
# Author: Bradley Burrell - October 2022
# Description: This script reformats the GS_closeout workbook to GS_closeout_corrected to be imported to Marine Noise
#              Register
# Version: 1.0
# REQUIREMENTS:
#  1. Python
#   1.1. Version 3.10
#  2. Modules
#   2.1. os
#   2.2. pandas
#   2.3. csv
#   2.4. datetime
#   2.5. sys

import os
import pandas as pd
import csv
import datetime
import sys

# Hard coded path to excel. Could be replaced with sys agru or similar
spreadsheet_path = "Update\\Path\\To\\Spreadsheet"

# Check spreadsheet is valid, if so code will run, else it will be exited.
file_exists = os.path.exists(spreadsheet_path)
if file_exists is True:
    print('Found Spreadsheet to reformat')
else:
    print('Spreadsheet Not Found, please check path on Line 23 is correct, Varible "spreadsheet_path"')
    sys.exit(1)

# Converts Excel workbook to pandas dataframe. This may need to be adjusted for new worksheets.
dataframe = pd.read_excel(spreadsheet_path,
                          sheet_name='Seismic Log',
                          header=13,
                          skiprows=0,
                          nrows=27)

# Renames First cell
dataframe.rename(columns={'Unnamed: 0': 'Oil and Gas Block'}, inplace=True)

# Deletes empty cells
columns_to_delete = []
for col in dataframe.columns:
    if str(col).startswith('Unnamed'):
        columns_to_delete.append(str(col))
for colum_to_delete in columns_to_delete:
    dataframe.drop(colum_to_delete, axis=1, inplace=True)

# Exports dataframe to csv for easy of manipulation
dataframe.to_csv('temp.csv')

# Opens the Temp csv
with open('temp.csv', newline='') as csvfile:
    # Creates Reader for CVS
    reader = csv.reader(csvfile, delimiter=',', quotechar='|')
    # Pull headers in the first line form cvs.
    headers = next(reader)
    # Create output CSV
    with open('Oil and Gas Block.csv', 'w', encoding='UTF8', newline='') as f:
        # creates writer for csv
        writer = csv.writer(f)
        # headers to csv file
        out_headers = ['Oil and Gas Block', 'Dates']
        # Writes headers to file
        writer.writerow(out_headers)
        # Iterates through the rows in the cvs
        for row in reader:
            # Pull Oil and Gas block ID
            og_block = row[1]
            # Craetes Variable for sort dates
            dates = ''
            # Count to pull date from header
            count = 0
            # Iterates through the columns in the row
            for value in row:
                # If the value is x then the date is pull from the header
                if value == 'x':
                    # Creates a date variable with pull the date from header and reformat to dd-mm-yyyy
                    date = datetime.datetime.strptime(headers[count][:10], "%Y-%m-%d").strftime("%d-%m-%Y")
                    # Adds add to date variable
                    dates = "{}, {}".format(dates, date)
                # Count is increase by 1 (next iteration will be the nextr colum along).
                count = count + 1
            # Once all columns have been checked, result added to a list to write
            out_row = [og_block, dates[2:]]
            # Out row is write to CSV
            writer.writerow(out_row)
# temp csv removed.
os.remove('temp.csv')

