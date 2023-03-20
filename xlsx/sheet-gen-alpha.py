
import os
import openpyxl

path        = r"D:\Users\25384\Desktop\xlsx"

os.chdir(path)


# Return the value of data type is workbook
workbook    = openpyxl.load_workbook('analogue-addendum.xlsx')
sheet       = workbook.active


# ****** Closed interval ******
Yni         = 2
Ynk         = 7
# ***** ***** ***** ***** *****



temp_name   = sheet[str('A' + str(Yni))]

print('temp_name: ' + str(temp_name.value))

# -----------------------------


for index in range(Ynk-1):
    index += Yni
    print('Index: ' + str(index) + '\tName: ' + sheet[str('D' + str(index))].value)


