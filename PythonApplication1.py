# Test for Logging Tenter 7 Data
# Written by Max Savage

import xlsxwriter
import time
from os import system
import sys

# User Experience
system("title " + "Tenter 7 Report File Converter - MTS")
print("Tenter 7 Report File Converter\n")
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
# ---
date = ""
frame = ""
shift = ""
user = ""
lot = ""
style = ""
src = ""

for line in f:
    lines.append(line)


# Iterating through the header
line = 0
char = 0

while line < 4:
    if line == 0:
        line = line + 1
        char = -1
    if line == 1:
        if char >= 16 and char < 36:
            date = date + lines[line][char]
        if char >= 42 and char < 55:
            frame = frame + lines[line][char]
        if char >= 68 and char < 83:
            shift = shift + lines[line][char]
        if char > 83:
            line = line + 1
            char = -1
    if line == 2:
        if char >= 16 and char < 36:
            user = user + lines[line][char]
        if char >= 42 and char < 60:
            lot = lot + lines[line][char]
        if char >= 68 and char < 80:
            style = style + lines[line][char]   
        if char > 83:
            line = line + 1
            char = -1
    if line == 3:
        line = line + 1
        char = -1
    char = char + 1 

# Create new workbook
workbook = "tenter7_" + style.strip() + "_" + date.strip() + ".xlsx"
wb = xlsxwriter.Workbook(workbook)
ws = wb.add_worksheet()
ws.set_column(0,20,12.5)

# Iterating through the information
line = 7
char = 0

while line >= 7 and line < len(lines):
    if char <= 7: # Timestamp
        timestamp = timestamp + lines[line][char]
    if char > 7 and char <= 17: # Entry Yards
        entryyds = entryyds + lines[line][char]
    if char > 17 and char <= 24: # Entry CPI
        entrycpi = entrycpi + lines[line][char]
    if char > 24 and char <= 33: # Entry YPM
        entryypm = entryypm + lines[line][char]
    if char > 33 and char <= 41: # Exit Yards
        exityds = exityds + lines[line][char]
    if char > 41 and char <= 48: # Exit CPI
        exitcpi = exitcpi + lines[line][char]
    if char > 48 and char <= 57: # Exit YPM
        exitypm = exitypm + lines[line][char]
    if char > 57 and char <= 64: # Overfeed
        offeed = offeed + lines[line][char]
    if char > 64 and char <= 71: # Target
        target = target + lines[line][char]
    if char > 71 and char < 78: # Mode
        mode = mode + lines[line][char]
    if char == 78:

        # Write to workbook
        ws.write('A'+str(line), timestamp.strip())
        ws.write('B'+str(line), entryyds.strip())
        ws.write('C'+str(line), entrycpi.strip())
        ws.write('D'+str(line), entryypm.strip())
        ws.write('E'+str(line), exityds.strip())
        ws.write('F'+str(line), exitcpi.strip())
        ws.write('G'+str(line), exitypm.strip())
        ws.write('H'+str(line), offeed.strip())
        ws.write('I'+str(line), target.strip())
        ws.write('J'+str(line), mode.strip())

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
        char = -1
    char = char + 1


# Write header
ws.write('A1', lines[0].strip())

ws.write('A2', "Report Date:")
ws.write('B2', date.strip())
ws.write('D2', "Frame:")
ws.write('E2', frame.strip())
ws.write('G2', "Shift:")
ws.write('H2', shift.strip())

ws.write('A3', "User:")
ws.write('B3', user.strip())
ws.write('D3', "Lot:")
ws.write('E3', lot.strip())
ws.write('G3', "Style:")
ws.write('H3', style.strip())

ws.write('A4', lines[3].strip())

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

time.sleep(2)
print("[DONE]")
time.sleep(1)