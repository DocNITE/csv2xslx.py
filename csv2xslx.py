import pandas as pd
import openpyxl
import csv
import sys
import os

from openpyxl.styles import Font
from openpyxl.utils import get_column_letter

"""
Данный скрипт помогает импортировать CSV файлы в единый XSLS, 
с использованием ТОЛЬКО суммарной информации.

Page,Device(No use!),Group,ItemID(Not use!),Item,Value
"""

SUM_INFO_DEF = "Суммарная информация"
ROWS_MAX_DEF = 256
FILE_NAME_DEF = sys.argv[1]
FILES_DIR_DEF = "./"

# os.system("py run.py")

"""
Проверяет является ли массив частью "Суммарной информации"
Возвращает сконвертированный его вариант
"""
def isSumInfo(arr):
    if arr[0] != SUM_INFO_DEF:
        return False
    # Group, Item, Value
    new_arr = [arr[2], arr[4], arr[5]]
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
        for row in reader:
            ri=ri+1
            if ri > ROWS_MAX_DEF:
                break
            arr_info = isSumInfo(row)
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