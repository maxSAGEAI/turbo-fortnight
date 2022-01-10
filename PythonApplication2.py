# Test for Logging Tenter 7 Data take 2
# Written by Max Savage

import xlsxwriter
import time
from os import system
import sys
import numpy

# User Experience
system("title " + "Tenter 7 Report File Converter Tool - MTS")
print("Tenter 7 Report File Converter Tool - v1.0\n")
file = input("Name of report file: ")
file = file + ".txt"

try:
    f = open(file)
except (OSError, IOError) as e:
    print('\033[91m' + "\nFile not found! Make sure you are typing the correct file name.")
    print("Exiting...\n" + '\033[0m')
    time.sleep(2)
    sys.exit(1)


print("\nConverting... ",  end = "")    

# Initialize Variables
lines = []
timestamp = "00.00"
entryyds = "00.00"
entrycpi = "00.00"
entryypm = "00.00"
exityds = "00.00"
exitcpi = "00.00"
exitypm = "00.00"
offeed = "00.00"
target = "00.00"
mode = "00.00"
# ---
date = "00.00"
frame = "00.00"
shift = "00.00"
user = "00.00"
lot = "00.00"
style = "00.00"
src = "00.00"

for line in f:
    lines.append(line)


# Iterating through the header
line = 0
char = 0
flag = 0

while line < 4:
    if line == 0:
        line = line + 1
    if line == 1:
        data = lines[line].split()
        if len(data) >= 6:
            date = data[1]
            frame = data[3]
            shift = data[5]
        else:
            flag = 1
        line = line + 1
    if line == 2:
        data1 = lines[line].split()
        if len(data1) >= 6:
            user = data1[1]
            dataS = lines[line].split("\t")
            lot = dataS[3]
            style = dataS[5]
        else:
            flag = 1
        line = line + 1
    if line == 3:
        data2 = lines[line].split()
        src = data2[1]
        line = line + 1

# Create new workbook
workbook = "tenter7_" + style.strip() + "_" + date + ".xlsx"
wb = xlsxwriter.Workbook(workbook)
ws = wb.add_worksheet()
ws.set_column(0,20,12.5)

# Iterating through the information
line = 7
char = 0

while line >= 7 and line < len(lines):
    data3 = lines[line].split()
    if len(data3) < 9:
        ws.write('A'+str(line), lines[line])
    else:
        timestamp = data3[0]
        entryyds = data3[1]
        entrycpi = data3[2]
        entryypm = data3[3]
        exityds = data3[4]
        exitcpi = data3[5]
        exitypm = data3[6]
        offeed = data3[7]
        target = data3[8]
        mode = data3[9]

        # Write to workbook
        ws.write('A'+str(line), timestamp)
        if entryyds == "_":
            ws.write('B'+str(line), entryyds)
        else:
            ws.write_number('B'+str(line), float(entryyds))
        if entrycpi == "_":
            ws.write('C'+str(line), entrycpi)
        else:
            ws.write_number('C'+str(line), float(entrycpi))
        if entryypm == "_":
            ws.write('D'+str(line), entryypm)
        else:
            ws.write_number('D'+str(line),float( entryypm))
        if exityds == "_":
            ws.write('E'+str(line), exityds)
        else:
            ws.write_number('E'+str(line), float(exityds))
        if exitcpi == "_":
            ws.write('F'+str(line), exitcpi)
        else:
            ws.write_number('F'+str(line), float(exitcpi))
        if exitypm == "_":
            ws.write('G'+str(line), exitypm)
        else:
            ws.write_number('G'+str(line), float(exitypm))
        if offeed == "_":
            ws.write('H'+str(line), offeed)
        else:
            ws.write_number('H'+str(line), float(offeed))
        if target == "_":
            ws.write('I'+str(line), target)
        else:
            ws.write_number('I'+str(line), float(target))
        ws.write('J'+str(line), mode)

        # Reset
        timestamp = ""
        entryyds = ""
        entrycpi = ""
        entryypm = ""
        exityds = ""
        exitcpi = ""
        exitypm = ""
        offeed = ""
        target = ""
        mode = ""
    line = line + 1


# Write header
ws.write('A1', lines[0])

ws.write('A2', "Report Date:")
ws.write('B2', date)
ws.write('D2', "Frame:")
ws.write('E2', frame)
ws.write('G2', "Shift:")
ws.write_number('H2', float(shift))

ws.write('A3', "User:")
ws.write('B3', user)
ws.write('D3', "Lot:")
ws.write('E3', lot)
ws.write('G3', "Style:")
ws.write('H3', style)

ws.write('A4', "Source:")
ws.write('B5', src)

ws.write('C5', "Entry")
ws.write('F5', "Exit")
ws.write('A6', "Timestamp")
ws.write('B6', "Yards")
ws.write('C6', "CPI")
ws.write('D6', "YPM")
ws.write('E6', "Yards")
ws.write('F6', "CPI")
ws.write('G6', "YPM")
ws.write('H6', "OFFEED")
ws.write('I6', "Target")
ws.write('J6', "Mode")

wb.close()

if flag == 1:
    print('\033[91m' + "[WARN: File was missing some header information]\n"+'\033[0m')

time.sleep(2)
print("[DONE]")
time.sleep(1)