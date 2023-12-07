import pandas as pd
import openpyxl
import csv
import sys
import os

from openpyxl.styles import Font
from openpyxl.utils import get_column_letter

"""
Данный скрипт помогает импортировать (импортнутые) .htm/.csv файлы
в .xlsx формат.

Page, NULL(not use!), Group, Item, Value
"""

SUM_INFO_DEF = "Суммарная информация"
END_INFO_DEF = "Имя компьютера"
ROWS_MAX_DEF = 256
FILE_NAME_DEF = sys.argv[1]
FILES_DIR_DEF = "./"

"""
Проверяет является ли массив частью "Суммарной информации"
Возвращает сконвертированный его вариант
"""
curr_group = "null"
def convertInfo(group, arr):
    global curr_group
    if group != curr_group:
        curr_group = group
        return False
    # Group, Item, Value
    new_arr = [curr_group, arr[3], arr[4]]
    return new_arr

"""
Возвращает всю информацию о файле для импорта в .xslx
 # ri - счестчик строк
 # arr_info - сконвертированный массив с нужным данными
 # result - возвращаемая инфа о файле
"""
def getFileInfo(file_name):
    with open (file_name, "r") as file:
        reader = csv.reader(file)
        result = [] # читаем это в writeExcel
        ri = 0
        n_group = ""
        can_do = False
        for row in reader:
            ri=ri+1
            if ri > ROWS_MAX_DEF:
                break
            if len(row) <= 0:
                continue
            if row[0] == SUM_INFO_DEF:
                can_do = True
                continue
            if can_do != True:
                continue
            if row[0] == END_INFO_DEF:
                print(row[0])
                break
            if row[2] != "":
                n_group = row[2]
            arr_info = convertInfo(n_group, row)
            if arr_info == False:
                continue
            result.append(arr_info)
        return [file_name, result]

def toFrame(data):
    result = []
    old_group = ""
    for arr in data:
        print(arr)
        print("\n")
        if old_group != arr[0]:
            result.append([arr[0], "", ""])
            old_group = arr[0]
            continue
        result.append(["", arr[1], arr[2]])
    return pd.DataFrame(result)

def writeExcel(data, name):
    with pd.ExcelWriter(name, engine="openpyxl") as writer:
        for data_f in data:
            sheet_n = data_f[0]
            # writing
            data_fc = toFrame(data_f[1])
            data_fc.to_excel(writer, sheet_name=sheet_n, index=False)
""" SHIT CODE
            # styling
            workbook = writer.book
            sheet = workbook[sheet_n]
            font = Font(name="Arial", size=24)
            for column in sheet.columns:
                letter = get_column_letter(column[0].column)
                sheet.column_dimensions[letter].font = font
"""
           
def getAllFiles(folder_path):
    extension = ".csv"
    file_list = []
    for file in os.listdir(folder_path):
        if file.endswith(extension):
            file_list.append(file)
    return file_list
           
def main():
    file_arrs = []
    for file in getAllFiles(FILES_DIR_DEF):
        file_arrs.append(getFileInfo(file))
        
    if len(file_arrs) == 0:
        print("Can't find .csv files!")
        return True
        
    writeExcel(file_arrs, FILE_NAME_DEF)
    print("\n Done! "+ FILE_NAME_DEF +" was created!")
main()

