import openpyxl as op
import openpyxl.styles as st
from openpyxl.utils import get_column_letter

import excel_format
def createNewTable(spisok, workSheet, path):
    font = st.Font(
        size=8,
        bold=False,
        italic=False,
        color="000000",
        name="Arial"
    )
    alignment = st.Alignment(horizontal='center', vertical = 'center')

    for row in range(1, 40 + len(spisok)):
        for col in range(1, 40):
            workSheet.cell(row = row, column = col).font = font
            workSheet.cell(row = row, column = col).alignment = alignment


    pathShablon = f"{path}\\Шаблон.xlsx"
    wbShablon = op.open(filename = pathShablon, data_only = False)
    ws = wbShablon['Настройки']
    row = 1
    spr = dict()
    while ws.cell(row = row, column = 1).value is not None:
        cl = ws.cell(row = row, column = 1).value
        clValue = ws.cell(row = row, column = 2).value
        clValue = float(clValue)
        if cl == "ФИО":
            spr.update({"FIO" : clValue})
        elif cl == "Табельный номер":
            spr.update({'Tabel' : clValue})
        elif cl == "Премия":
            spr.update({'Prem' : clValue})
        elif cl == "Должность":
            spr.update({'Dol' : clValue})
        elif cl == "Дата":
            spr.update({'Data' : clValue})
        elif cl == "Итого":
            spr.update({'Itog' : clValue})
        row += 1


    workSheet.column_dimensions['A'].width = spr['FIO'] #фамилия имя отчество
    workSheet.column_dimensions['B'].width = spr['Tabel'] #Табельный номер
    workSheet.column_dimensions['C'].width = spr['Prem'] #Премия
    workSheet.column_dimensions['D'].width = spr['Dol'] #Должность
    col = 4
    for i in range(15):
        col += 1
        letter = get_column_letter(col)
        workSheet.column_dimensions[letter].width = spr['Data'] #Даты

    col += 1
    letter = get_column_letter(col)
    workSheet.column_dimensions[letter].width = spr['Itog']

    n = excel_format.getMonth(arg='count')
    for i in range(16, n + 1):
        col += 1
        letter = get_column_letter(col)
        workSheet.column_dimensions[letter].width = spr['Data'] #Даты

    col += 1
    letter = get_column_letter(col)
    workSheet.column_dimensions[letter].width = spr['Itog']

    #Объединение ячеек
    workSheet.merge_cells('B11:C11')

    col += 31 - n
    letter = get_column_letter(col)
    workSheet.merge_cells(f'E11:{letter}11')

    workSheet.merge_cells('A11:A12')
    workSheet.merge_cells('D11:D12')

    #Заполняем шапку таблицы
    workSheet.cell(row = 11, column = 1).value = "Фамилия, имя, отчество"
    workSheet.cell(row = 11, column = 1).alignment = st.Alignment(wrap_text=True,
                                                                  horizontal='center', vertical='center')
    workSheet.cell(row = 11, column = 2).value = "Учетный номер"
    workSheet.cell(row = 12, column = 2).value = "Табельный номер"
    workSheet.cell(row = 12, column = 3).value = "Премия"
    workSheet.cell(row = 11, column = 4).value = "Должность (профессия)"
    workSheet.cell(row = 11, column = 4).alignment = st.Alignment(wrapText=True,
                                                                  horizontal='center', vertical = 'center')
    workSheet.cell(row = 11, column = 5).value = "Числа месяца"
    col = 4
    for i in range(n):
        col += 1
        if i == 15:
            workSheet.cell(row = 12, column = col).value = "Итого дней (часов) явок (неявок) с 1 по 15"
            workSheet.cell(row = 12, column = col).alignment = st.Alignment(wrap_text=True)
            col += 1
        workSheet.cell(row = 12, column = col).value = i + 1

    workSheet['B12'].alignment = st.Alignment(textRotation=90, horizontal='center', vertical = 'center')
    workSheet['C12'].alignment = st.Alignment(textRotation=90, horizontal='center', vertical = 'center')

