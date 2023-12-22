import pandas as pd
import openpyxl
import csv
import sys
import os

from openpyxl.styles import Border, Side, Font, Alignment
from openpyxl.utils import get_column_letter

"""
Данный скрипт помогает импортировать CSV файлы в единый XSLS, 
с использованием ТОЛЬКО суммарной информации.

Page,Device(No use!),Group,ItemID(Not use!),Item,Value

TODO: Сделать игнор для Ввод, DMI, некоторых полей.
Затем добавить автотитульник от "Имя компьютера"
Также название для sheet должно идти от "Имя компьютера"
"""

SUM_INFO_DEF = "Суммарная информация"
FONT_DEF = "Times New Roman"
ROWS_MAX_DEF = 256
MAIN_TITLE_SIZE_DEF = 20
TITLE_SIZE_DEF = 14
CONTENT_SIZE_DEF = 12
A_ROW_WIDTH_DEF = 11
B_ROW_WIDTH_DEF = 30
C_ROW_WIDTH_DEF = 50
IGNORE_TITLES = ["DMI", "Ввод"]
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

"""
Конвертирует полученный массив в excel data для таблицы.
"""
def toFrame(data):
    result = []
    result.append(["", "", ""])
    old_group = ""
    for arr in data:
        print(arr)
        print("\n")
        if old_group != arr[0]:
            result.append([arr[0]+":", "", ""])
            old_group = arr[0]
            # continue
        result.append(["", arr[1], arr[2]])
    return pd.DataFrame(result)

"""
Записывает данные в excel. Also - добавляет стили и все такое
"""
def writeExcel(data, name):
    with pd.ExcelWriter(name, engine="openpyxl") as writer:
        for data_f in data:
            sheet_n = data_f[0] # todo: shoulde be use PC name from report
            # writing
            data_fc = toFrame(data_f[1])
            data_fc.to_excel(writer, sheet_name=sheet_n, index=False)
            #todo: also - add title support from pc name 
            # styling
            workbook = writer.book
            sheet = workbook[sheet_n]
            sheet.column_dimensions['A'].width = A_ROW_WIDTH_DEF
            sheet.column_dimensions['B'].width = B_ROW_WIDTH_DEF
            sheet.column_dimensions['C'].width = C_ROW_WIDTH_DEF
            border_style = Side(border_style="thick")
            border = Border(top=border_style, left=border_style, right=border_style, bottom=border_style)
            alignment = Alignment(horizontal='left', vertical='center', wrap_text=True)
            alignment_a = Alignment(horizontal='left')
            # formating B column 
            for cell in sheet['B']:
                cell.font = Font(name=FONT_DEF, size=CONTENT_SIZE_DEF)
                cell.border = border
                cell.alignment = alignment
            # formating C column 
            for cell in sheet['C']:
                cell.font = Font(name=FONT_DEF, size=CONTENT_SIZE_DEF)
                cell.border = border
                cell.alignment = alignment
            # formating A column 
            ai = 1      # A cells iterate
            s_ai = 0    # start empty A cells iterate
            for cell in sheet['A']:
                if cell.value != "":
                    if s_ai != 0:
                        # merge for vertical cells 
                        sheet.merge_cells("A"+str(s_ai)+":A"+str(ai-1))
                        # restore start A interate
                        s_ai = 0
                    # merge for current horizontal title line
                    sheet.merge_cells("A"+str(ai)+":C"+str(ai))
                else:
                    if s_ai == 0:
                        s_ai = ai
                cell.font = Font(name=FONT_DEF, size=TITLE_SIZE_DEF)
                cell.border = border
                cell.alignment = alignment_a
                ai = ai + 1
            if s_ai != 0:
                sheet.merge_cells("A"+str(s_ai)+":A"+str(ai-1))
                s_ai = 0
            alignment_title = Alignment(horizontal='center')
            border_title = Border(bottom=border_style)
            sheet['A1'].value = "There should b ename"
            sheet.merge_cells('A1:C1')
            sheet['A1'].font = Font(name=FONT_DEF, size=MAIN_TITLE_SIZE_DEF)
            sheet['A1'].border = border_title
            sheet['A1'].alignment = alignment_title
            workbook.save(name)
           
"""
Ищет все нужные файлы для бандла
"""
def getAllFiles(folder_path):
    extension = ".csv"
    file_list = []
    for file in os.listdir(folder_path):
        if file.endswith(extension):
            file_list.append(file)
    return file_list
           
"""
Это база
"""
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