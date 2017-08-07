import openpyxl
import csv
import os
from numpy import median
from numpy import average
from copy import copy

# Config vars

# Overall
PROTOCOL_WB = "2017_07_18_Cecilia Practice Array_Protocol.xlsx"
OUTPUT_WB = "2017_07_18_Cecilia Practice Array_Results.xlsx"
NUM_INPUT = 20 # number of samples
NUM_BLOCKS = 16 # number of blocks on slide
SAVE_ENABLED = True

# Results File Info
DATA_COL = 'Z'
FLAG_COL = 'A'
NAME_COL = 'G'
BLOC_COL = 'D'
FIRST_ROW_DATA = 34

# Protocol File Info
SAMPLE_COL = 'B'
SECOND_COL = 'E'
SAMPLE_ROW = 20

# ===== FUNCTIONS ======

# returns True if variable can be cast as Float
def is_number(s):
    try:
        float(s)
        return True
    except ValueError:
        return False
    except TypeError:
        return False

# creates copy of sheet [num] in specified workbook
# sets all digit values < 1 to 1
# inserts new sheet after copy of specified sheet with name + " floored"
def sheetfloor(wb,num):
    og = wb.worksheets[num]
    ws = wb.create_sheet(og.title+" floored", num+1)
    for j in range(len(og['A'])):
        for i in range(len(og[1])):
            value = og.cell(row = j+1, column = i+1).value
            font = copy(og.cell(row = j+1, column = i+1).font)
            ws.cell(row = j+1, column = i+1).font = font
            if is_number(value) and float(value) < 1:
                ws.cell(row = j+1, column = i+1).value = 1
            else:
                ws.cell(row = j+1, column = i+1).value = value



# ===== SCRIPT =====

# Set up WB's
wb_protocol = openpyxl.load_workbook(PROTOCOL_WB)
os.chdir("Results File")

# Combine all data into one .xlsx book with each sample on a page

wb_combined = openpyxl.Workbook()
for i in range(NUM_INPUT):
    if(i > len(wb_combined.sheetnames) - 1):
       wb_combined.create_sheet()
    print("loading sheet " + str(i))
    ws = wb_combined.worksheets[i]
    curnum = str(i+1).zfill(len(str(NUM_INPUT)))
    # everything else is 1-indexed and 0-padded

    ws.title = curnum
    with open("plate"+curnum+".txt") as data:
        reader = csv.reader(data, delimiter='\t')
        for row in reader:
            ws.append(row)
if(SAVE_ENABLED):
    wb_combined.save("data_combined.xlsx")
    print("data_combined.xlsx saved")

os.chdir("..")

# create one sheet in a book with all relevant data, flagged samples removed

wb_working = openpyxl.Workbook()
ws = wb_working.worksheets[0]
ws.title = "raw medians"


for i in range(len(wb_combined.worksheets)):
    currsheet = wb_combined.worksheets[i]
    print("condensing sheet " + str(i))
    j = 1 #side note: I hate 1-indexing
    # get ready to see it a lot
    while(currsheet[DATA_COL+str(j+FIRST_ROW_DATA)].value is not None):
        if(i==0):
            ws['A'+str(j+1)] = (currsheet[NAME_COL+str(j+FIRST_ROW_DATA)].value) + "_" + (currsheet[BLOC_COL+str(j+FIRST_ROW_DATA)].value)
        if(currsheet[FLAG_COL+str(j+FIRST_ROW_DATA)].value == '-100'):
            ws.cell(row = j+1, column = i+2).value = 'NA'
        else:
            ws.cell(row = j+1, column = i+2).value = currsheet[DATA_COL+str(j+FIRST_ROW_DATA)].value
        j = j+1

# add sample names for column titles

print("Adding sample names from Protocol File")
for i in range(NUM_INPUT):
    ws.cell(row = 1, column = i+2).value = wb_protocol['Protocol'][SAMPLE_COL + str(SAMPLE_ROW+i)].value

# performing median confinement

print("Consolidating identical analytes")
data = dict()
for i in range(len(ws['A'])-1):
    key = ws['A'+str(i+2)].value
    if key not in data.keys():
        data[key] = [list() for a in range(NUM_INPUT)]
        print("Adding analyte ID "+key+"...")
    for j in range(NUM_INPUT):
        data[key][j].append(ws.cell(row = i+2, column = j+2).value)

print("Creating new worksheet")
ws = wb_working.create_sheet("median medians")

print("Adding sample and secondary names from Protocol File")
for i in range(NUM_INPUT):
    ws.cell(row = 1, column = i+2).value = wb_protocol['Protocol'][SAMPLE_COL + str(SAMPLE_ROW+i)].value + "_" + wb_protocol['Protocol'][SECOND_COL + str(SAMPLE_ROW+i)].value

print("Calculating median values")
i = 2
for key in data.keys():
    ws.cell(row = i, column = 1).value = key
    print("Finding median of "+key+"...")
    for j in range(NUM_INPUT):
        values = [int(a) for a in data[key][j] if a != 'NA']
        if values == list():
            ws.cell(row = i, column = j+2).value = 'NA'
        else:
            ws.cell(row = i, column = j+2).value = median(values)
    i = i+1

print("Setting floor to 1")
sheetfloor(wb_working, 1)

print("Consolidating identical samples")
ws = wb_working.worksheets[2]

data = dict()
num_analytes = len(ws['A'])-1
for i in range(NUM_INPUT):
    key = ws.cell(row = 1, column = i+2).value
    if key not in data.keys():
        data[key] = [list() for a in range(num_analytes)]
        print("Adding sample ID "+key+"...")
    for j in range(num_analytes):
        data[key][j].append(ws.cell(row = j+2, column = i+2).value)

print("Creating new worksheet")
ws = wb_working.create_sheet("mean samples")

print("Adding analyte names from previous sheet")
for i in range(num_analytes):
    ws.cell(column = 1, row = i+2).value = wb_working.worksheets[2].cell(column = 1, row = i+2).value

print("Calculating average values")
i = 2
for key in data.keys():
    ws.cell(row = 1, column = i).value = key
    print("Averaging values for "+key+"...")
    for j in range(num_analytes):
        values = [float(a) for a in data[key][j] if a != 'NA']
        if values == list():
            ws.cell(row = j+2, column = i).value = 'NA'
        else:
            ws.cell(row = j+2, column = i).value = average(values)
    i = i+1


print("Subtracting PBS per block and removing block identifiers")
num_samples = len(ws[1]) - 1
data = [dict() for a in range(NUM_BLOCKS)]
for i in range(len(ws['A'])-1):
    ID = ws.cell(column = 1, row = i+2).value
    block = int(ID.split("_")[1]) - 1
    key = ID.split("_")[0]
    data[block][key] = [0 for a in range(num_samples)]
    print("Adding analyte ID "+key+" to block "+str(block)+"...")
    for j in range(num_samples):
        data[block][key][j] = ws.cell(row = i+2, column = j+2).value

print("Creating new worksheet")
ws = wb_working.create_sheet("PBS corrected")

print("Adding sample names from previous sheet")
for i in range(num_samples):
    ws.cell(column = i+2, row = 1).value = wb_working.worksheets[3].cell(column = i+2, row = 1).value

currrow = 2
for block in range(16):
    for key in data[block].keys():
        ws.cell(column = 1, row = currrow).value = key
        for i in range(len(data[block][key])):
            curval = data[block][key][i]
            curPBS = data[block]["PBS"][i]
            if(curval=='NA'):
                ws.cell(column = i+2, row = currrow).value = 'NA'
            elif(curPBS=='NA'):
                # if the PBS value is invalid, leave as is and change format
                ws.cell(column = i+2, row = currrow).value = curval
                ws.cell(column = i+2, row = currrow).font = openpyxl.styles.Font(italic=True)
            else:
                ws.cell(column = i+2, row = currrow).value = curval - curPBS
        currrow = currrow + 1

print("Setting floor to 1")
sheetfloor(wb_working, 4)

print("Subtracting blanks from matching secondary")
ws = wb_working.worksheets[5]

data = dict()
styles = dict()
for i in range(num_samples):
    ID = ws.cell(row = 1, column = i+2).value
    sample = ID.split("_")[0]
    secondary = ID.split("_")[1]
    if secondary not in data.keys():
        data[secondary]=dict()
        styles[secondary]=dict()
    for col in ws.iter_cols(min_row=2, min_col=i+2,max_col=i+2):
        data[secondary][sample]=[cell.value for cell in col]
        styles[secondary][sample]=[cell.font for cell in col] #need to be sure to copy styles at this point

print("Creating new worksheet")
ws = wb_working.create_sheet("blank subtracted")

print("Adding analyte names from previous sheet")
for i in range(num_analytes):
    ws.cell(column = 1, row = i+2).value = wb_working.worksheets[5].cell(column = 1, row = i+2).value

currcol = 2
for secondary in data.keys():
    for sample in data[secondary].keys():
        ws.cell(row = 1, column = currcol).value = sample
        for i in range(len(data[secondary][sample])):
            curval = data[secondary][sample][i]
            curblank = data[secondary]['Blank'][i]
            if(curval=='NA'):
                ws.cell(column = currcol, row = i+2).value = 'NA'
            elif(curblank=='NA'):
                ws.cell(column = currcol, row = i+2).value = curval
               # ws.cell(column = currcol, row = i+2).font = openpyxl.Font(bold=True)
            else:
                ws.cell(column = currcol, row = i+2).value = curval - curblank
            ws.cell(column = currcol, row = i+2).font = copy(styles[secondary][sample][i])
        currcol = currcol + 1

print("Setting floor to 1")
sheetfloor(wb_working, 6)
        
if(SAVE_ENABLED):
    wb_working.save("Results_Normalization_Process.xlsx")  
    print("Saved file 'Results_Normalization_Process.xlsx'")

    while(len(wb_working.worksheets) > 1):
        wb_working.remove_sheet(wb_working.worksheets[0])
    wb_working.worksheets[0].title = "master"
    wb_working.save(OUTPUT_WB)  
    print("Saved file '" + OUTPUT_WB + "'")


